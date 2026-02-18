/*
Project Name: FMX Equipment Import non-Gem
Project Version: 2.00
Filename: ImportExport.gs
File Version: 2.01
Chat link: [Insert Link]
*/

/**
 * @file ImportExport.gs
 * @description Handles Import and Export logic.
 */

/**
 * Opens the Import Dialog.
 */
function openImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('IMPORTdialog')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Data');
}

/**
 * Compatibility wrapper for the Sidebar button.
 * Redirects 'promptForImport' calls to the new dialog.
 */
function promptForImport() {
  openImportDialog();
}

/**
 * Handles the server-side import logic called by the dialog.
 * Decodes Base64 input, converts Excel to values (via Drive API), or parses CSV.
 * @param {string} dataUrl - The base64 data URL of the file OR raw text.
 * @param {string} fileType - The MIME type of the file.
 * @param {string} fileName - The name of the file (optional).
 * @return {string} Status message.
 */
function importData(dataUrl, fileType, fileName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.sheets.import);
    
    if (!sheet) {
      throw new Error(`Sheet "${CONFIG.sheets.import}" not found.`);
    }

    let blob;
    const name = fileName || "Imported File";

    // 1. Decode Data to Blob
    if (dataUrl.indexOf("data:") === 0) {
      const base64Data = dataUrl.split(',')[1];
      const decodedBytes = Utilities.base64Decode(base64Data);
      blob = Utilities.newBlob(decodedBytes, fileType, name);
    } else {
      blob = Utilities.newBlob(dataUrl, 'text/csv', name);
    }

    // 2. Determine File Type & Extract Data
    const isExcel = fileType.includes('excel') || 
                    fileType.includes('spreadsheetml') || 
                    name.endsWith('.xlsx') || 
                    name.endsWith('.xls');

    let data = [];

    if (isExcel) {
      // Excel Conversion Logic (Requires Drive API Service)
      try {
        const resource = {
          title: name,
          name: name, // V3 uses 'name', V2 uses 'title'
          mimeType: MimeType.GOOGLE_SHEETS
        };
        
        let tempFile;
        
        // Detect Drive API Version (V2 uses insert, V3 uses create)
        if (Drive.Files.insert) {
          tempFile = Drive.Files.insert(resource, blob);
        } else if (Drive.Files.create) {
          tempFile = Drive.Files.create(resource, blob);
        } else {
          throw new Error("Drive API enabled but neither 'insert' (V2) nor 'create' (V3) methods found.");
        }
        
        const tempFileId = tempFile.id;
        
        // Open the temp sheet and get values
        const tempSs = SpreadsheetApp.openById(tempFileId);
        const tempSheet = tempSs.getSheets()[0];
        data = tempSheet.getDataRange().getValues();
        
        // Clean up: Delete the temp file (V2 uses remove, V3 uses delete)
        if (Drive.Files.remove) {
          Drive.Files.remove(tempFileId);
        } else if (Drive.Files.delete) {
           Drive.Files.delete(tempFileId);
        }

      } catch (driveError) {
        if (driveError.message && driveError.message.includes("Drive is not defined")) {
          return "Error: To import Excel files, you must enable 'Drive API' in the Services menu (left sidebar).";
        }
        throw driveError;
      }

    } else {
      // CSV Parsing Logic
      const csvContent = blob.getDataAsString();
      data = Utilities.parseCsv(csvContent);
    }

    if (!data || data.length === 0) {
      return "Error: No data found in file.";
    }

    // 3. Write Data to Sheet
    sheet.clear();
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    // 4. Dynamically Find Headers
    // Search for the row where the first column matches the ItemID config (e.g., "ID*")
    const targetHeaderID = CONFIG.columnNames.ItemID[0];
    let headerRow = [];
    let foundHeader = false;

    for (let i = 0; i < data.length; i++) {
      // Check first column, trimmed
      if (data[i][0] && data[i][0].toString().trim() === targetHeaderID) {
        headerRow = data[i];
        foundHeader = true;
        break;
      }
    }

    if (!foundHeader) {
      return `Warning: Imported ${data.length} rows, but could not find header row starting with '${targetHeaderID}'. Headers not extracted.`;
    }

    // Filter out empty headers or non-string garbage
    const cleanHeaders = headerRow.filter(h => h && h.toString().trim() !== "");

    // Update the Data sheet with these headers
    console.log(`Extracting ${cleanHeaders.length} headers.`);
    updateDataSheetHeaders(cleanHeaders);

    return `Success: Imported ${data.length} rows from ${name}. Headers extracted.`;

  } catch (e) {
    console.error("Import Error: " + e.message);
    throw e;
  }
}

/**
 * Updates the 'Import_Headers' column in the Data sheet with the provided headers.
 * @param {Array<string>} headers - The array of headers extracted from the imported file.
 */
function updateDataSheetHeaders(headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.sheets.data);

  if (!dataSheet) {
    console.error(`Sheet "${CONFIG.sheets.data}" not found.`);
    return;
  }

  // 1. Find the column index for "Import_Headers"
  // We assume headers are in Row 1 of the Data sheet
  const lastCol = dataSheet.getLastColumn();
  if (lastCol === 0) return; 

  const sheetHeaders = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const targetHeaderName = CONFIG.reportRanges.Import_Headers;
  
  // indexOf is 0-based, so we add 1 for the spreadsheet column index
  const colIndex = sheetHeaders.indexOf(targetHeaderName);

  if (colIndex === -1) {
    console.error(`Column header "${targetHeaderName}" not found in ${CONFIG.sheets.data} sheet.`);
    return;
  }
  const colNumber = colIndex + 1;
  console.log(`Found "${targetHeaderName}" at Column ${colNumber}.`);

  // 2. Clear existing data below the header in that specific column
  const maxRows = dataSheet.getMaxRows();
  if (maxRows > 1) {
    // Clear from Row 2 down to the bottom
    dataSheet.getRange(2, colNumber, maxRows - 1, 1).clearContent();
  }

  // 3. Paste the new headers vertically
  if (headers && headers.length > 0) {
    // filter out empty headers if necessary, or keep them to maintain index alignment
    // Transpose 1D array to 2D array [[val], [val]]
    const outputValues = headers.map(h => [h]);
    
    dataSheet.getRange(2, colNumber, outputValues.length, 1).setValues(outputValues);
    console.log(`Pasted ${headers.length} headers to Data sheet.`);
  }
}

/**
 * Placeholder for export logic.
 */
function handleExport() {
  // Logic from previous steps or new logic will go here
}
