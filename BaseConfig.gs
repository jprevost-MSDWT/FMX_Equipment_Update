/*
Project Name: FMX Equipment Import non-Gem
Project Version: 1.00
Filename: BaseConfig.gs
File Version: 1.01
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
 * Automatically opens the sidebar.
 */
function onOpen() {
  createCustomMenu();
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
