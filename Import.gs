/*
Project Name: FMX Equipment Import non-Gem
Project Version: 6.00
Filename: Import.gs
File Version: 6.05
Chat link: [Insert Link]
*/

/**
 * @file Import.gs
 * @description Handles Import logic with robust Drive API error handling.
 *              On xlsx import:
 *                1. Converts xlsx to Google Sheet to read data.
 *                2. Writes Equipment data to RAWImport; validates headers and
 *                   transfers to Equipment_Edit.
 *                3. Writes Meters tab data to RAWImport_Meters; transfers to
 *                   Meters_Edit.
 *                4. Only after all successful validation and transfers:
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
 *
 * For xlsx files:
 *   - Reads the Equipment tab and writes it to RAWImport, then transfers to Equipment_Edit.
 *   - Reads the Meters tab and writes it to RAWImport_Meters, then transfers to Meters_Edit.
 *   - Saves metadata and a Google Sheet copy only after all transfers succeed.
 *
 * @param {string} dataUrl  - The base64 data URL of the file OR raw text.
 * @param {string} fileType - The MIME type of the file.
 * @param {string} fileName - The name of the file (optional).
 * @return {string} Status message.
 */
function importData(dataUrl, fileType, fileName) {
  // Lifted to outer scope so they're accessible in the finally/catch cleanup block
  // and for the saveExportTemplate call at the end of a successful import.
  let blob        = null;
  let convertedId = null;
  let isExcel     = false;

  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
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

    // 2. Determine File Type
    isExcel = fileType.includes('excel') ||
              fileType.includes('spreadsheetml') ||
              name.endsWith('.xlsx') ||
              name.endsWith('.xls');

    let equipmentData = [];
    let metersData    = [];

    if (isExcel) {
      let tempFileId = null;
      try {
        if (typeof Drive === 'undefined') {
          throw new Error("Drive API Service not detected.");
        }

        // Convert xlsx to Google Sheet to read all tabs
        const resource = {
          title: name,
          name: name,
          mimeType: MimeType.GOOGLE_SHEETS
        };

        let tempFile = Drive.Files.insert
          ? Drive.Files.insert(resource, blob)
          : Drive.Files.create(resource, blob);
        tempFileId = tempFile.id;

        // Assign to outer scope immediately so the catch block can clean it up
        // even if the subsequent openById or sheet lookups throw.
        convertedId = tempFileId;

        const tempSs = SpreadsheetApp.openById(tempFileId);

        // ── Read Equipment tab by name ────────────────────────────────────
        const equipmentSheet = tempSs.getSheetByName(CONFIG.sheets.export);
        if (!equipmentSheet) {
          throw new Error(`The required sheet "${CONFIG.sheets.export}" was not found in the imported file.`);
        }
        equipmentData = equipmentSheet.getDataRange().getValues();

        // ── Read Meters tab by name ───────────────────────────────────────
        const metersSheet = tempSs.getSheetByName(CONFIG.sheets.metersExport);
        if (metersSheet) {
          metersData = metersSheet.getDataRange().getValues();
        } else {
          console.warn(`No "${CONFIG.sheets.metersExport}" tab found in the imported xlsx. RAWImport_Meters will not be updated.`);
        }

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
      // Note: tempFileId / convertedId is intentionally NOT trashed here.
      // saveExportTemplate (called after successful transfers below) reuses it
      // via convertedId to make a permanent copy, then trashes it itself.

    } else {
      // CSV / TXT path — Equipment data only; Meters not applicable
      const csvContent = blob.getDataAsString();
      equipmentData = Utilities.parseCsv(csvContent);
    }

    if (!equipmentData || equipmentData.length === 0) {
      throw new Error("No data found in file.");
    }

    // 3. Write Equipment Data to RAWImport
    sheet.clear();
    sheet.getRange(1, 1, equipmentData.length, equipmentData[0].length).setValues(equipmentData);

    // 4. Determine whether the imported file contains real meters data.
    // Checking length alone is not reliable — an empty Google Sheet tab returns
    // [[""]] (length 1), which would cause a false positive. Instead, confirm
    // that the required meters header is actually present somewhere in the data.
    const requiredMeterHeader = CONFIG.mapping.metersRequired[0];
    const hasMetersData = metersData.some(
      row => row.some(cell => cell && cell.toString().trim() === requiredMeterHeader)
    );

    // 5. Write Meters Data to RAWImport_Meters, or clear it if no valid meters data is present.
    // Clearing on empty prevents stale data from a previous import from persisting.
    const metersImportSheet = ss.getSheetByName(CONFIG.sheets.metersImport);
    if (hasMetersData) {
      if (!metersImportSheet) {
        throw new Error(`Sheet ${CONFIG.sheets.metersImport} not found. Please run setup.`);
      }
      metersImportSheet.clear();
      const miMaxRows = metersImportSheet.getMaxRows();
      const miMaxCols = metersImportSheet.getMaxColumns();
      if (miMaxRows < metersData.length) {
        metersImportSheet.insertRowsAfter(miMaxRows, metersData.length - miMaxRows);
      }
      if (miMaxCols < metersData[0].length) {
        metersImportSheet.insertColumnsAfter(miMaxCols, metersData[0].length - miMaxCols);
      }
      metersImportSheet.getRange(1, 1, metersData.length, metersData[0].length).setValues(metersData);
    } else if (metersImportSheet) {
      metersImportSheet.clear();
    }

    // 6. Extract Equipment Headers (Dynamic Search).
    // Uses some()+trim() for robustness against leading/trailing whitespace in cells.
    let headerRowIndex = -1;
    const searchLimit   = Math.min(10, equipmentData.length);
    const requiredHeader = CONFIG.mapping.required[0] || "ID*";

    for (let i = 0; i < searchLimit; i++) {
      if (equipmentData[i].some(cell => cell != null && cell.toString().trim() === requiredHeader)) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      throw new Error(`Could not find required header '${requiredHeader}' in the first ${searchLimit} rows of the file.`);
    }

    const headerRow    = equipmentData[headerRowIndex];
    const cleanHeaders = headerRow.filter(h => h && h.toString().trim() !== "");

    // Update the available header options in the Data sheet
    updateDataSheetHeaders(cleanHeaders);

    // 7. Transfer Equipment data to Equipment_Edit
    const equipmentResult = processImportedData();

    // 8. Transfer Meters data to Meters_Edit, or clear it if no valid meters data is present.
    // Clearing on empty prevents stale data from a previous import from persisting.
    let metersResult = "";
    if (hasMetersData) {
      metersResult = processMetersData();
    } else {
      const metersEditSheet = ss.getSheetByName(CONFIG.sheets.metersEdit);
      if (metersEditSheet) {
        metersEditSheet.clearContents();
      }
    }

    // 9. Save export template — only reached if all validation and transfers
    //    above succeeded. Saves metadata and Google Sheet copy for export use.
    if (isExcel) {
      saveExportTemplate(blob, convertedId);
    }

    const statusParts = [`File "${name}" imported successfully.`, equipmentResult];
    if (metersResult) statusParts.push(metersResult);
    return statusParts.join("\n");

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
  const props   = PropertiesService.getScriptProperties();
  const metaKeys = CONFIG.scriptProperties.metadataFiles;

  // ── 1. Extract metadata files from the xlsx zip ───────────────────────────
  const zipBlob   = blob.copyBlob().setContentType('application/zip');
  const zipEntries = Utilities.unzip(zipBlob);
  const entryMap  = {};
  zipEntries.forEach(function(entry) {
    entryMap[entry.getName()] = entry;
  });

  const missingFiles = [];
  const metaToStore  = {};

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
    missingFiles.forEach(function(zipPath) {
      props.deleteProperty(metaKeys[zipPath]);
    });
  }

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

  // ── 4. Make a permanent Google Sheet copy from the already-converted file
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
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.sheets.data);

  if (!dataSheet) {
    console.error(`Sheet "${CONFIG.sheets.data}" not found.`);
    return;
  }

  const lastCol = dataSheet.getLastColumn();
  if (lastCol === 0) return;

  const sheetHeaders     = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const targetHeaderName = CONFIG.namedRanges.Import_Headers;
  const colIndex         = sheetHeaders.indexOf(targetHeaderName);

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
