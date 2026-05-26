/*
Project Name: FMX Equipment Import non-Gem
Project Version: 5.00
Filename: Import.gs
File Version: 3.09
Chat link: [Insert Link]
*/

/**
 * @file Import.gs
 * @description Handles Import logic with robust Drive API error handling.
 *              On xlsx import:
 *                1. Converts xlsx to Google Sheet to read data.
 *                2. Validates headers and transfers data to Equipment_Edit.
 *                3. Only after successful validation and transfer:
 *                   - Extracts and stores FMX metadata zip entries as base64
 *                     in Script Properties.
 *                   - Creates a permanent Google Sheet copy for export use.
 *                   - Saves the original xlsx to Drive for audit purposes.
 *                   - Trashes any previously stored copies.
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
 * For xlsx files, saves metadata and a Google Sheet copy for use during export,
 * but only after the import has been fully validated and transferred successfully.
 * @param {string} dataUrl - The base64 data URL of the file OR raw text.
 * @param {string} fileType - The MIME type of the file.
 * @param {string} fileName - The name of the file (optional).
 * @return {string} Status message.
 */
function importData(dataUrl, fileType, fileName) {
  // Lifted to outer scope so they're accessible after the isExcel block
  // for the saveExportTemplate call at the end of a successful import.
  let blob = null;
  let convertedId = null;
  let isExcel = false;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.sheets.import);

    if (!sheet) {
      throw new Error(`Sheet "${CONFIG.sheets.import}" not found.`);
    }

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
    isExcel = fileType.includes('excel') ||
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

        // Convert xlsx to Google Sheet to read the data
        const resource = {
          title: name,
          name: name,
          mimeType: MimeType.GOOGLE_SHEETS
        };

        let tempFile = Drive.Files.insert
          ? Drive.Files.insert(resource, blob)
          : Drive.Files.create(resource, blob);
        tempFileId = tempFile.id;

        const tempSs = SpreadsheetApp.openById(tempFileId);
        data = tempSs.getSheets()[0].getDataRange().getValues();

        // Store the converted ID at outer scope — used by saveExportTemplate
        // at the end of the function after successful validation.
        convertedId = tempFileId;

      } catch (err) {
        if (
          err.message.includes("Drive API Service not detected") ||
          err instanceof ReferenceError ||
          err.message.includes("Drive is not defined")
        ) {
          throw new Error("Advanced Drive Service is not enabled. Please go to 'Services' (+), find 'Drive API', and add it to the project.");
        }
        throw new Error("XLSX Conversion Error: " + err.message);
      }
      // Note: tempFileId is intentionally NOT trashed here.
      // saveExportTemplate (called after successful transfer below) reuses it
      // via convertedId to make a permanent copy, then trashes it itself.

    } else {
      const csvContent = blob.getDataAsString();
      data = Utilities.parseCsv(csvContent);
    }

    if (!data || data.length === 0) {
      throw new Error("No data found in file.");
    }

    // 3. Write Data to RAWImport Sheet
    sheet.clear();
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    // 4. Extract Headers (Dynamic Search)
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

    // 5. Transfer to Equipment_Edit based on current selection
    const transferResult = processImportedData();

    // 6. Save export template — only reached if all validation and transfer
    //    above succeeded. Saves metadata and Google Sheet copy for export use.
    if (isExcel) {
      saveExportTemplate(blob, convertedId);
    }

    return `File "${name}" imported successfully.\n${transferResult}`;

  } catch (e) {
    // If something failed and we have a convertedId that saveExportTemplate
    // never got to use, clean it up now to avoid orphaned Drive files.
    if (convertedId) {
      try {
        DriveApp.getFileById(convertedId).setTrashed(true);
      } catch (cleanupError) {
        console.warn("Cleanup of converted temp file failed: " + cleanupError.message);
      }
    }
    console.error("Import Error: " + e.message);
    throw e;
  }
}

/**
 * Saves everything needed for export:
 *   1. Extracts the four FMX metadata zip entries from the xlsx blob and stores
 *      them as base64 strings in Script Properties.
 *   2. Makes a permanent Google Sheet copy from the already-converted file
 *      (reuses convertedId to avoid a second Drive conversion).
 *   3. Saves a fresh xlsx copy to Drive for audit purposes.
 *   4. Trashes any previously stored xlsx template and Google Sheet template.
 *   5. Trashes the temporary converted file (convertedId) once copied.
 *
 * Only called after a fully successful import and data transfer.
 *
 * @param {GoogleAppsScript.Base.Blob} blob        - The original xlsx blob.
 * @param {string}                     convertedId - File ID of the already-converted
 *                                                   Google Sheet (reused to avoid a
 *                                                   second Drive conversion).
 */
function saveExportTemplate(blob, convertedId) {
  const props = PropertiesService.getScriptProperties();
  const metaKeys = CONFIG.scriptProperties.metadataFiles;

  // ── 1. Extract metadata files from the xlsx zip ───────────────────────────
  // Utilities.unzip requires content type to be application/zip regardless
  // of the original file's MIME type.
  blob.setContentType('application/zip');
  const zipEntries = Utilities.unzip(blob);
  const entryMap = {};
  zipEntries.forEach(function(entry) {
    entryMap[entry.getName()] = entry;
  });

  const missingFiles = [];
  const metaToStore = {};

  Object.keys(metaKeys).forEach(function(zipPath) {
    const propKey = metaKeys[zipPath];
    if (entryMap[zipPath]) {
      metaToStore[propKey] = Utilities.base64Encode(entryMap[zipPath].getBytes());
    } else {
      missingFiles.push(zipPath);
      console.warn('Metadata file not found in xlsx: ' + zipPath);
    }
  });

  if (missingFiles.length > 0) {
    console.warn(
      'The following FMX metadata files were not found in the imported xlsx: ' +
      missingFiles.join(', ') +
      '. The exported file may be rejected by FMX. ' +
      'Try re-downloading the template from FMX and importing again.'
    );
  }

  // Store all found metadata in Script Properties
  Object.keys(metaToStore).forEach(function(propKey) {
    props.setProperty(propKey, metaToStore[propKey]);
  });

  // ── 2. Trash previously stored files ─────────────────────────────────────
  const prevXlsxId  = props.getProperty(CONFIG.scriptProperties.templateFileId);
  const prevSheetId = props.getProperty(CONFIG.scriptProperties.templateSheetId);

  if (prevXlsxId) {
    try { DriveApp.getFileById(prevXlsxId).setTrashed(true); }
    catch (e) { console.warn('Could not trash previous xlsx template: ' + e.message); }
  }
  if (prevSheetId) {
    try { DriveApp.getFileById(prevSheetId).setTrashed(true); }
    catch (e) { console.warn('Could not trash previous sheet template: ' + e.message); }
  }

  // ── 3. Save a fresh xlsx copy to Drive (for audit / future metadata reads) 
  const xlsxFile = DriveApp.createFile(blob.setName('FMX_Export_Template.xlsx'));
  props.setProperty(CONFIG.scriptProperties.templateFileId, xlsxFile.getId());

  // ── 4. Make a permanent Google Sheet copy from the already-converted file ─
  const sheetCopy = DriveApp.getFileById(convertedId).makeCopy('FMX_Export_Template_Sheet');
  props.setProperty(CONFIG.scriptProperties.templateSheetId, sheetCopy.getId());

  // ── 5. Trash the temporary converted file now that we have a permanent copy 
  try {
    DriveApp.getFileById(convertedId).setTrashed(true);
  } catch (e) {
    console.warn('Could not trash temporary converted file: ' + e.message);
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
