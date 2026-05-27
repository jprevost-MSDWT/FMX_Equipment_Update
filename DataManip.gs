/*
Project Name: FMX Equipment Import non-Gem
Project Version: 6.00
Filename: DataManip.gs
File Version: 6.01
Chat link: [Insert Link]
*/

/**
 * @file DataManip.gs
 * @description Handles data manipulation, header mapping, and transfer from
 *              RAWImport → Equipment_Edit and RAWImport_Meters → Meters_Edit.
 */

/**
 * Orchestrates the data transfer from RAWImport to Equipment_Edit based on Selected_Headers.
 * This is the primary function triggered after a successful import.
 * @return {string} Status message describing rows and columns transferred.
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
 * Transfers all columns from RAWImport_Meters to Meters_Edit.
 * No user-configurable header selection — all columns are copied as-is.
 * The header row is located by searching for the Meters required marker
 * defined in CONFIG.mapping.metersRequired.
 * @return {string} Status message describing rows and columns transferred.
 */
function processMetersData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(CONFIG.sheets.metersImport);
    const targetSheet = ss.getSheetByName(CONFIG.sheets.metersEdit);

    if (!sourceSheet || !targetSheet) {
      throw new Error("Source (RAWImport_Meters) or Target (Meters_Edit) sheet is missing.");
    }

    const sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length < 1) throw new Error("RAWImport_Meters sheet is empty.");

    // Locate header row using the Meters required marker
    const requiredMarker = CONFIG.mapping.metersRequired[0];
    let headerRowIndex = -1;
    for (let i = 0; i < Math.min(CONFIG.mapping.headerSearchLimit, sourceData.length); i++) {
      if (sourceData[i].includes(requiredMarker)) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      throw new Error(`Could not find the header row in RAWImport_Meters (searched for "${requiredMarker}").`);
    }

    // All columns are used — no filtering needed
    const headerRow  = sourceData[headerRowIndex].map(h => h ? h.toString().trim() : "");
    const dataRows   = sourceData.slice(headerRowIndex + 1);
    const finalOutput = [headerRow, ...dataRows];

    targetSheet.clearContents();
    targetSheet.getRange(1, 1, finalOutput.length, headerRow.length).setValues(finalOutput);

    return `Success: Transferred ${dataRows.length} rows and ${headerRow.length} columns to ${CONFIG.sheets.metersEdit}.`;

  } catch (e) {
    console.error("Meters Data Transfer Error: " + e.message);
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
 *
 * Equipment Items merge strategy (per row, per column):
 *   - Rows are matched between RAWImport and Equipment_Edit using the ID* column as the key.
 *   - For matched rows, Equipment_Edit values take priority over RAWImport values.
 *   - New items (blank or missing ID* in Equipment_Edit) are exported as-is.
 *
 * Meters export strategy:
 *   - Meters_Edit is written directly to the "Meters" export sheet.
 *   - All columns are included; no RAWImport merge is performed.
 *
 * @return {void}
 */
function runExportProcess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const exportSheet      = ss.getSheetByName(CONFIG.sheets.export);
  const importSheet      = ss.getSheetByName(CONFIG.sheets.import);
  const editSheet        = ss.getSheetByName(CONFIG.sheets.edit);
  const metersEditSheet  = ss.getSheetByName(CONFIG.sheets.metersEdit);
  const metersExportSheet = ss.getSheetByName("Meters");

  if (!exportSheet || !importSheet || !editSheet) {
    throw new Error("One or more required Equipment sheets are missing. Please verify sheet names.");
  }
  if (!metersEditSheet || !metersExportSheet) {
    throw new Error("One or more required Meters sheets are missing (Meters_Edit or Meters). Please verify sheet names.");
  }

  // ── EQUIPMENT ITEMS ───────────────────────────────────────────────────────

  // 1. Validate source before touching the export sheet
  const importLastCol = importSheet.getLastColumn();
  const importLastRow = importSheet.getLastRow();

  if (importLastCol === 0) return; // Nothing to export

  if (importLastRow < CONFIG.rows.importHeaderCount) {
    throw new Error(
      `The source sheet "${CONFIG.sheets.import}" does not have the required ` +
      `${CONFIG.rows.importHeaderCount} header rows.`
    );
  }

  // 2. Clear export sheet and copy header rows from RAWImport
  exportSheet.clear();

  importSheet
    .getRange(1, 1, CONFIG.rows.importHeaderCount, importLastCol)
    .copyTo(exportSheet.getRange(1, 1));

  // 3. Resolve column headers
  const rawData      = importSheet.getDataRange().getValues();
  const rawHeaderRow = rawData[CONFIG.rows.importHeaderCount - 1]
    .map(h => h ? h.toString().trim() : "");
  const rawDataRows  = rawData.slice(CONFIG.rows.importHeaderCount);

  const targetHeaders = rawHeaderRow;
  if (targetHeaders.length === 0) return;

  const editData     = editSheet.getDataRange().getValues();
  if (editData.length <= CONFIG.rows.editHeaderIndex) return;

  const editHeaders  = editData[CONFIG.rows.editHeaderIndex - 1]
    .map(h => h ? h.toString().trim() : "");
  const editDataRows = editData.slice(CONFIG.rows.editHeaderIndex);

  // 4. Build a lookup map: ID* → RAWImport row
  const rawIdColIndex = rawHeaderRow.indexOf(CONFIG.mapping.required[0]);
  const rawById       = {};

  if (rawIdColIndex !== -1) {
    rawDataRows.forEach(row => {
      const id = row[rawIdColIndex] ? row[rawIdColIndex].toString().trim() : "";
      if (id !== "") {
        rawById[id] = row;
      }
    });
  }

  // 5. Resolve column mappings
  const editIdColIndex = editHeaders.indexOf(CONFIG.mapping.required[0]);
  const colMapRaw  = targetHeaders.map(h => rawHeaderRow.indexOf(h));
  const colMapEdit = targetHeaders.map(h => editHeaders.indexOf(h));

  // 6. Build merged output rows
  const outputData = editDataRows.map(editRow => {
    const itemId = (editIdColIndex !== -1 && editRow[editIdColIndex] !== undefined)
      ? editRow[editIdColIndex].toString().trim()
      : "";

    const rawRow = (itemId !== "" && rawById[itemId]) ? rawById[itemId] : null;

    return targetHeaders.map((_, colIdx) => {
      const editVal = colMapEdit[colIdx] !== -1 ? editRow[colMapEdit[colIdx]] : undefined;
      const rawVal  = (rawRow && colMapRaw[colIdx] !== -1) ? rawRow[colMapRaw[colIdx]] : "";

      return colMapEdit[colIdx] !== -1 ? editVal : rawVal;
    });
  });

  // 7. Write merged Equipment data to export sheet
  if (outputData.length > 0 && outputData[0].length > 0) {
    exportSheet.getRange(
      CONFIG.rows.importHeaderCount + 1, 1,
      outputData.length,
      outputData[0].length
    ).setValues(outputData);
  }

  // ── METERS ────────────────────────────────────────────────────────────────

  // Direct copy: Meters_Edit → Meters export sheet (no RAWImport merge needed)
  const metersData = metersEditSheet.getDataRange().getValues();

  metersExportSheet.clearContents();

  if (metersData.length > 0 && metersData[0].length > 0) {
    metersExportSheet.getRange(1, 1, metersData.length, metersData[0].length).setValues(metersData);
  }
}

// EOF: DataManip.gs
