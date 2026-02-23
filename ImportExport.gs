/*
Project Name: FMX Equipment Import non-Gem
Project Version: 2.00
Filename: ImportExport.gs
File Version: 2.02
Chat link: [Insert Link]
*/

/**
 * @file ImportExport.gs
 * @description Handles Import logic and Header Extraction.
 */

function openImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('IMPORTdialog')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Data');
}

function promptForImport() {
  openImportDialog();
}

function importData(dataUrl, fileType, fileName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.sheets.import);
    if (!sheet) throw new Error(`Sheet "${CONFIG.sheets.import}" not found.`);

    let blob;
    const name = fileName || "Imported File";

    if (dataUrl.indexOf("data:") === 0) {
      const base64Data = dataUrl.split(',')[1];
      const decodedBytes = Utilities.base64Decode(base64Data);
      blob = Utilities.newBlob(decodedBytes, fileType, name);
    } else {
      blob = Utilities.newBlob(dataUrl, 'text/csv', name);
    }

    const isExcel = fileType.includes('excel') || 
                    fileType.includes('spreadsheetml') || 
                    name.endsWith('.xlsx') || 
                    name.endsWith('.xls');

    let data = [];

    if (isExcel) {
      const resource = { title: name, name: name, mimeType: MimeType.GOOGLE_SHEETS };
      let tempFile = Drive.Files.insert ? Drive.Files.insert(resource, blob) : Drive.Files.create(resource, blob);
      const tempSs = SpreadsheetApp.openById(tempFile.id);
      data = tempSs.getSheets()[0].getDataRange().getValues();
      Drive.Files.remove ? Drive.Files.remove(tempFile.id) : Drive.Files.delete(tempFile.id);
    } else {
      data = Utilities.parseCsv(blob.getDataAsString());
    }

    if (!data || data.length === 0) return "Error: No data found.";

    sheet.clear();
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    const targetHeaderID = CONFIG.columnNames.ItemID[0];
    let headerRow = data.find(row => row[0] && row[0].toString().trim() === targetHeaderID);

    if (!headerRow) return "Warning: Data imported, but ID header not found.";

    const cleanHeaders = headerRow.filter(h => h && h.toString().trim() !== "");
    updateDataSheetHeaders(cleanHeaders);

    return `Success: Imported ${data.length} rows. Headers extracted.`;
  } catch (e) {
    throw e;
  }
}

function updateDataSheetHeaders(headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
  if (!dataSheet) return;

  const sheetHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const colIndex = sheetHeaders.indexOf(CONFIG.reportRanges.Import_Headers);

  if (colIndex === -1) return;
  const colNumber = colIndex + 1;

  if (dataSheet.getMaxRows() > 1) {
    dataSheet.getRange(2, colNumber, dataSheet.getMaxRows() - 1, 1).clearContent();
  }

  if (headers && headers.length > 0) {
    const outputValues = headers.map(h => [h]);
    dataSheet.getRange(2, colNumber, outputValues.length, 1).setValues(outputValues);
  }
}
