/*
Project Name: FMX Equipment Import non-Gem
Project Version: 1.00
Filename: BaseConfig.gs
File Version: 1.04
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
    placeholderRange1: 7,  // This is a placeholder for the real value later
    Import_Headers: "Import_Headers"
  },
  namedRanges: {
    Import_Headers: "Import_Headers",
    Import_Header_Exclude: "Import_Header_Exclude",
    Import_Headers_Selection: "Import_Headers_Selection"
  },
  columnNames: {
    ItemID: ["ID*"],
    Item_Name: ["Tag*"],
    Item_Type: ["Type*"],
    Item_Building: ["Building*"],
  }
};

/**
 * Creates a custom menu when the spreadsheet opens.
 * Automatically opens the sidebar and verifies environment.
 */
function onOpen() {
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
 * Creates them if they are missing.
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
 * Finds the matching header name in the Data sheet and sets that column
 * (minus the header) as the named range.
 */
function SetupNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheets.data);

  if (!sheet) return;

  const lastCol = sheet.getLastColumn();
  // If no columns, can't set ranges
  if (lastCol === 0) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rangesToSetup = CONFIG.namedRanges;

  // Iterate over the keys (Range Names)
  Object.keys(rangesToSetup).forEach(rangeName => {
    const headerName = rangesToSetup[rangeName];
    const colIndex = headers.indexOf(headerName);

    if (colIndex !== -1) {
      // Column found. Define range from Row 2 to Max Row
      const colLetter = colIndex + 1;
      const numRows = sheet.getMaxRows() - 1; // Exclude header

      if (numRows > 0) {
        const range = sheet.getRange(2, colLetter, numRows, 1);
        ss.setNamedRange(rangeName, range);
      }
    }
  });
}
