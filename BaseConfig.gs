/*
Project Name: FMX Equipment Import non-Gem
Project Version: 3.00
Filename: BaseConfig.gs
File Version: 3.00
Chat link: [Insert Link]
*/

/**
 * Configuration object for the project.
 * Centralizes all static strings and settings.
 */
const CONFIG = {
  sheets: {
    import: "RAWImport",    // Do NOT change. Also used in HTML
    data: "Data"
  },
  reportRanges: {
    placeholderRange1: 7,  
    Import_Headers: "Import_Headers"
  },
  namedRanges: {
    Import_Headers: "Import_Headers",
    Import_Header_Exclude: "Import_Header_Exclude",
    Import_Headers_Selection: "Import_Headers_Selection",
    Selected_Headers: "Selected_Headers",
    // New Named Ranges for Equipment Data
    Equipment_types_Import: "Equipment_types_Import",
    Equipment_modules_Import: "Equipment_modules_Import",
    Equipment_types_Filtered: "Equipment_types_Filtered"
  },
  columnNames: {
    ItemID: ["ID*"],
    Item_Name: ["Tag*"],
    Item_Type: ["Type*"],
    Item_Building: ["Building*"],
  },
  mapping: {
    // Headers that must ALWAYS be included in Selected_Headers
    required: ["ID*", "Tag*", "Type*", "Building*"],
    // Headers selected by default but can be unselected by user
    default_selected: [
      "01.1:Manufacturer",
      "01.2: Model Number",
      "01.3:Serial Number"
    ]
  }
};

/**
 * Creates a custom menu and sets up the sheet when opened via an installable trigger.
 */
function OnOpen_Triggered(e) {
  createCustomMenu();
  VerifySheets();
  SetupNamedRanges();
  showSidebar();
}

/**
 * Builds and adds the custom menu to the UI.
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Excel Tools")
    .addItem('Open Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Import Excel File', 'promptForImport')
    .addToUi();
}

/**
 * Opens the HTML Sidebar.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle("Excel Tools Control")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Verifies that all sheets defined in CONFIG exist.
 */
function VerifySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheetNames = ss.getSheets().map(sheet => sheet.getName());
  const requiredSheetNames = Object.values(CONFIG.sheets);

  requiredSheetNames.forEach(sheetName => {
    if (allSheetNames.indexOf(sheetName) === -1) {
      ss.insertSheet(sheetName);
    }
  });
}

/**
 * Maps named ranges based on column headers in the Data sheet.
 */
function SetupNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheets.data);
  if (!sheet) return;

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rangesToSetup = CONFIG.namedRanges;

  Object.keys(rangesToSetup).forEach(rangeName => {
    const headerName = rangesToSetup[rangeName];
    const colIndex = headers.indexOf(headerName);
    if (colIndex !== -1) {
      const colLetter = colIndex + 1;
      const numRows = sheet.getMaxRows() - 1;
      if (numRows > 0) {
        const range = sheet.getRange(2, colLetter, numRows, 1);
        ss.setNamedRange(rangeName, range);
      }
    }
  });
}

/**
 * Consolidates all header settings into a single object for the UI.
 * This reduces round-trips from the client to the server.
 * @return {Object} An object containing required, default, and optional header arrays.
 */
function getHeaderConfiguration() {
  try {
    return {
      required: getRequiredHeaders(),
      defaults: getDefaultSelectedHeaders(),
      optionals: getImportHeaderOptions()
    };
  } catch (e) {
    console.error("Error in getHeaderConfiguration: " + e.message);
    throw new Error("Failed to load header configuration.");
  }
}

/**
 * Fetches required headers from CONFIG.
 * @return {string[]}
 */
function getRequiredHeaders() {
  return CONFIG.mapping.required;
}

/**
 * Fetches default selected headers from CONFIG.
 * @return {string[]}
 */
function getDefaultSelectedHeaders() {
  return CONFIG.mapping.default_selected;
}

/**
 * Fetches available header options, excluding those that are required
 * or default-selected (as those are handled separately in UI).
 * @return {string[]} 
 */
function getImportHeaderOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rangeName = CONFIG.namedRanges.Import_Headers_Selection;
  const range = ss.getRangeByName(rangeName);
  
  if (!range) return [];
  
  const required = CONFIG.mapping.required;
  const defaultSel = CONFIG.mapping.default_selected;
  const values = range.getValues().flat();
  
  // Filter out empty values and values already handled by specific logic
  return values.filter(item => {
    return item !== "" && 
           item !== null && 
           !required.includes(item) && 
           !defaultSel.includes(item);
  });
}

/**
 * Saves the selected headers to the Data sheet.
 * @param {string[]} selectedHeaders - User selected headers.
 */
function saveSelectedHeaders(selectedHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheets.data);
  if (!sheet) throw new Error("Data sheet not found.");

  const headerName = CONFIG.namedRanges.Selected_Headers;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colIndex = headers.indexOf(headerName);

  if (colIndex === -1) throw new Error(`Column "${headerName}" not found.`);
  
  const colNumber = colIndex + 1;
  const maxRows = sheet.getMaxRows();

  if (maxRows > 1) {
    sheet.getRange(2, colNumber, maxRows - 1, 1).clearContent();
  }

  if (selectedHeaders && selectedHeaders.length > 0) {
    const output = selectedHeaders.map(h => [h]);
    sheet.getRange(2, colNumber, output.length, 1).setValues(output);
  }
}
