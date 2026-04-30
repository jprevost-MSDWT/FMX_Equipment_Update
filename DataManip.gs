/*
Project Name: FMX Equipment Import non-Gem
Project Version: 3.00
Filename: DataManip.gs
File Version: 3.01
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

    // 2. Identify Header Row in Source (Dynamic Search)
    const requiredMarker = CONFIG.mapping.required[0];
    let headerRowIndex = -1;
    for (let i = 0; i < Math.min(20, sourceData.length); i++) {
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
 * Placeholder for future formatting logic.
 */
function formatEquipmentEditSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheets.edit);
  if (!sheet) return;
  
  // Example: Auto-resize columns
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}
