/*
Project Name: FMX Equipment Import non-Gem
Project Version: 5.00
Filename: Import.gs
File Version: 3.06
Chat link: [Insert Link]
*/

/**
 * @file Import.gs
 * @description Handles Import logic with robust Drive API error handling.
 *              Saves the imported xlsx as a template for use during export.
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
 * Manually triggers the column mapping and data transfer to Equipment_Edit.
 * Designed to be called from the Custom Menu or Sidebar Actions.
 */
function manualProcessImport() {
  try {
    const result = processImportedData();
    SpreadsheetApp.getUi().alert(result);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Manual Transfer Failed: " + e.message);
  }
}

/**
 * Handles the server-side import logic called by the dialog.
 * Decodes Base64 input, converts Excel to values (via Drive API), or parses CSV.
 * For xlsx files, saves a copy of the original blob to Drive for use as an
 * export template, preserving FMX-specific metadata (custom properties, drawings, etc.).
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
      let tempFileId = null;
      try {
        if (typeof Drive === 'undefined') {
          throw new Error("Drive API Service not detected.");
        }

        const resource = {
          title: name,
          name: name,
          mimeType: MimeType.GOOGLE_SHEETS
        };

        let tempFile = Drive.Files.insert ? Drive.Files.insert(resource, blob) : Drive.Files.create(resource, blob);
        tempFileId = tempFile.id;
        const tempSs = SpreadsheetApp.openById(tempFileId);
        data = tempSs.getSheets()[0].getDataRange().getValues();

        // 3. Save the original xlsx blob as the export template in Drive.
        //    This preserves FMX-specific metadata (custom properties, drawings, etc.)
        //    that Google's xlsx export would otherwise strip out.
        saveExportTemplate(blob);

      } catch (err) {
        if (
          err.message.includes("Drive API Service not detected") ||
          err instanceof ReferenceError ||
          err.message.includes("Drive is not defined")
        ) {
          throw new Error("Advanced Drive Service is not enabled. Please go to 'Services' (+), find 'Drive API', and add it to the project.");
        }
        throw new Error("XLSX Conversion Error: " + err.message);
      } finally {
        if (tempFileId) {
          try {
            DriveApp.getFileById(tempFileId).setTrashed(true);
          } catch (cleanupError) {
            console.warn("Temporary file cleanup failed: " + tempFileId);
          }
        }
      }

    } else {
      const csvContent = blob.getDataAsString();
      data = Utilities.parseCsv(csvContent);
    }

    if (!data || data.length === 0) {
      throw new Error("No data found in file.");
    }

    // 4. Write Data to Sheet
    sheet.clear();
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    // 5. Extract Headers (Dynamic Search)
    let headerRowIndex = -1;
    const searchLimit = Math.min(10, data.length);
    const requiredHeader = CONFIG.mapping.required[0] || "ID*";

    for (let i = 0; i < searchLimit; i++) {
      if (data[i].includes(requiredHeader)) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      throw new Error(`Could not find required header '${requiredHeader}' in the first ${searchLimit} rows of the file.`);
    }

    const headerRow = data[headerRowIndex];
    const cleanHeaders = headerRow.filter(h => h && h.toString().trim() !== "");

    // Update the available header options in the Data sheet
    updateDataSheetHeaders(cleanHeaders);

    // 6. Transfer to Equipment_Edit based on current selection
    const transferResult = processImportedData();

    return `File "${name}" imported successfully.\n${transferResult}`;

  } catch (e) {
    console.error("Import Error: " + e.message);
    throw e;
  }
}

/**
 * Saves the imported xlsx blob to Drive as the export template.
 * Trashes any previously saved template to avoid accumulating stale files.
 * The file ID is persisted in Script Properties for retrieval during export.
 * @param {GoogleAppsScript.Base.Blob} blob - The original xlsx blob from the import.
 */
function saveExportTemplate(blob) {
  const props = PropertiesService.getScriptProperties();
  const propKey = CONFIG.scriptProperties.templateFileId;
  const existingId = props.getProperty(propKey);

  // Trash the previous saved template if one exists
  if (existingId) {
    try {
      DriveApp.getFileById(existingId).setTrashed(true);
    } catch (e) {
      console.warn("Could not trash previous template file: " + e.message);
    }
  }

  // Save the new template blob and store its file ID
  const savedFile = DriveApp.createFile(blob.setName('FMX_Export_Template.xlsx'));
  props.setProperty(propKey, savedFile.getId());
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

  const lastCol = dataSheet.getLastColumn();
  if (lastCol === 0) return;

  const sheetHeaders = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const targetHeaderName = CONFIG.namedRanges.Import_Headers;

  const colIndex = sheetHeaders.indexOf(targetHeaderName);

  if (colIndex === -1) {
    console.error(`Column header "${targetHeaderName}" not found in ${CONFIG.sheets.data} sheet.`);
    return;
  }
  const colNumber = colIndex + 1;

  const lastRow = dataSheet.getLastRow();
  if (lastRow > 1) {
    dataSheet.getRange(2, colNumber, lastRow - 1, 1).clearContent();
  }

  if (headers && headers.length > 0) {
    const outputValues = headers.map(h => [h]);
    dataSheet.getRange(2, colNumber, outputValues.length, 1).setValues(outputValues);
  }
}

// EOF: Import.gs
