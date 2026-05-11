/*
Project Name: FMX Equipment Import non-Gem
Project Version: 4.00
Filename: DataManip.gs
File Version: 3.05
Chat link: [Insert Link]
*/

/**
 * @file DataManip.gs
 * @description Handles data manipulation, header mapping, and transfer from RAWImport to Equipment_Edit.
 */

/**
 * Orchestrates the data transfer from RAWImport to Equipment_Edit based on Selected_Headers.
 * This is the primary function triggered after a successful import.
 */
function processImportedData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const selectedHeaders = getSelectedHeadersList(ss);
    
    if (selectedHeaders.length === 0) {
      throw new Error("No headers found in 'Selected_Headers' range. Please configure settings in the sidebar.");
    }

    const sourceSheet = ss.getSheetByName(CONFIG.sheets.import);
    const targetSheet = ss.getSheetByName(CONFIG.sheets.edit);
    
    if (!sourceSheet || !targetSheet) {
      throw new Error("Source (RAWImport) or Target (Equipment_Edit) sheet is missing.");
    }

    // 1. Get Source Data
    const sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length < 1) throw new Error("Source sheet is empty.");

    // 2. Identify Header Row in Source using the first required header,
    //    which is the database identifier and will always be present.
    const requiredMarker = CONFIG.mapping.required[0];
    let headerRowIndex = -1;
    for (let i = 0; i < Math.min(CONFIG.mapping.headerSearchLimit, sourceData.length); i++) {
      if (sourceData[i].includes(requiredMarker)) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      throw new Error(`Could not find the header row in RAWImport (searched for ${requiredMarker}).`);
    }

    const sourceHeaders = sourceData[headerRowIndex].map(h => h ? h.toString().trim() : "");
    const dataRows = sourceData.slice(headerRowIndex + 1);

    // 3. Map Selected Headers to Source Column Indices
    const colIndices = selectedHeaders.map(header => {
      const idx = sourceHeaders.indexOf(header);
      return idx; // Will be -1 if not found
    });

    // 4. Build Output Matrix
    const output = dataRows.map(row => {
      return colIndices.map(idx => (idx !== -1 ? row[idx] : ""));
    });

    // 5. Prepare Target Sheet
    targetSheet.clearContents();
    
    // Write Headers then Data
    const finalOutput = [selectedHeaders, ...output];
    targetSheet.getRange(1, 1, finalOutput.length, selectedHeaders.length).setValues(finalOutput);

    return `Success: Transferred ${output.length} rows and ${selectedHeaders.length} columns to ${CONFIG.sheets.edit}.`;

  } catch (e) {
    console.error("Data Transfer Error: " + e.message);
    throw e;
  }
}

/**
 * Retrieves the list of selected headers from the Named Range.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The active spreadsheet.
 * @return {string[]} Array of header names.
 */
function getSelectedHeadersList(ss) {
  const range = ss.getRangeByName(CONFIG.namedRanges.Selected_Headers);
  if (!range) return [];
  
  return range.getValues().flat()
    .map(h => h ? h.toString().trim() : "")
    .filter(h => h !== "");
}

/**
 * Sidebar wrapper: saves the selected headers, flushes pending writes to ensure
 * spreadsheet state is synchronized, then re-syncs Equipment_Edit — all in one
 * server round trip. Keeps saveSelectedHeaders() and processImportedData()
 * independently callable for other contexts.
 * @param {string[]} selectedHeaders - The full list of headers (required + user-selected) to save and apply.
 * @return {string} Final status message from processImportedData.
 */
function saveAndProcessHeaders(selectedHeaders) {
  saveSelectedHeaders(selectedHeaders);
  SpreadsheetApp.flush();
  return processImportedData();
}

/**
 * Main controller function to execute the export process.
 * Clears the export tab, copies header rows from RAWImport, and maps
 * data from Equipment_Edit using column header matching.
 * @return {void}
 */
function runExportProcess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const exportSheet = ss.getSheetByName(CONFIG.sheets.export);
  const importSheet = ss.getSheetByName(CONFIG.sheets.import);
  const editSheet = ss.getSheetByName(CONFIG.sheets.edit);

  if (!exportSheet || !importSheet || !editSheet) {
    throw new Error("One or more required sheets are missing. Please verify sheet names.");
  }

  // 1) Clear all data & formatting from Edit_Export
  exportSheet.clear();

  // 2) Copy the first N header rows from RAWImport to Edit_Export
  const importLastCol = importSheet.getLastColumn();
  if (importLastCol > 0) {
    const topRowsRange = importSheet.getRange(1, 1, CONFIG.rows.importHeaderCount, importLastCol);
    topRowsRange.copyTo(exportSheet.getRange(1, 1));
  }

  // 3) Pull values from Equipment_Edit and transfer to Edit_Export, matching headers
  const exportLastCol = exportSheet.getLastColumn();
  if (exportLastCol === 0) return; // No headers to match against

  // Get target headers from the designated header row in Edit_Export
  const targetHeaders = exportSheet.getRange(CONFIG.rows.exportHeaderIndex, 1, 1, exportLastCol).getValues()[0];

  // Get all data from Equipment_Edit for batch processing
  const editData = editSheet.getDataRange().getValues();
  if (editData.length <= CONFIG.rows.editHeaderIndex) return; // No data rows below the header

  const sourceHeaders = editData[CONFIG.rows.editHeaderIndex - 1];
  const sourceRecords = editData.slice(CONFIG.rows.editHeaderIndex);

  // Map columns: Target Column Index -> Source Column Index
  const columnMap = targetHeaders.map(header => {
    if (!header || header.toString().trim() === "") return -1;
    return sourceHeaders.indexOf(header);
  });

  // Build output 2D array by mapping source records to the target column order
  const outputData = sourceRecords.map(record => {
    return columnMap.map(sourceColIndex => {
      return sourceColIndex !== -1 ? record[sourceColIndex] : "";
    });
  });

  // Write the mapped data in a single batch operation below the copied header rows
  if (outputData.length > 0 && outputData[0].length > 0) {
    exportSheet.getRange(
      CONFIG.rows.importHeaderCount + 1, 1,
      outputData.length, outputData[0].length
    ).setValues(outputData);
  }
}
