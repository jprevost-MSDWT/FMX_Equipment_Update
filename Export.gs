/*
Project Name: FMX Equipment Import non-Gem
Project Version: 5.00
Filename: Export.gs
File Version: 2.12
Chat link: [Insert Link]
*/

/**
 * @file Export.gs
 * @description Handles file export logic. Fetches the saved xlsx import template
 *              directly from Drive, replaces only the worksheet data via zip
 *              manipulation, and returns the result as base64. All FMX-specific
 *              metadata (custom properties, drawings, styles, etc.) is preserved
 *              because the template zip is never passed through Google's conversion.
 *
 * NOTE: The FMX custom properties in docProps/custom.xml contain two version values:
 *   FmxImportTemplate = 3
 *   FmxTemplateVersion = 12
 * These are preserved from the saved template file. If FMX ever rejects the upload
 * after a fresh import, check whether these values have changed in the new template.
 */

/**
 * Displays a modal dialog to initiate the file download.
 * Fetches data via asynchronous client-side call to prevent UI freezing.
 * Uses Blob + createObjectURL to avoid data: URL size limits in Chrome.
 * @return {void}
 */
function showDownloadDialog() {
  const htmlString = `
    <script>
      function triggerDownload(base64Data) {
        try {
          const byteCharacters = atob(base64Data);
          const byteNumbers = new Array(byteCharacters.length);
          for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
          }
          const byteArray = new Uint8Array(byteNumbers);
          const blob = new Blob([byteArray], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          });

          const url = URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url;
          link.download = '${CONFIG.Exporting.EXPORT_FILE_NAME}';
          document.body.appendChild(link);
          link.click();

          URL.revokeObjectURL(url);
          document.getElementById('status').innerText = 'Download initiated. Closing...';
          setTimeout(function() {
            google.script.host.close();
          }, 1500);
        } catch(e) {
          handleError(e);
        }
      }

      function handleError(error) {
        document.getElementById('status').innerText = 'Error: ' + error.message;
        document.getElementById('status').style.color = 'red';
      }

      window.onload = function() {
        google.script.run
          .withSuccessHandler(triggerDownload)
          .withFailureHandler(handleError)
          .getExportData();
      };
    <\/script>
    <div style="font-family: Arial, sans-serif; padding: 20px; text-align: center;">
      <p id="status" style="color: #333; font-size: 16px; margin-bottom: 10px;">Gathering data, please wait...</p>
    </div>
  `;

  const html = HtmlService.createHtmlOutput(htmlString)
      .setWidth(300)
      .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Downloading Report...');
}

/**
 * Fetches the exported XLSX data as a Base64 string.
 * Loads the saved template xlsx from Drive, unzips it, replaces only the
 * worksheet XML entries with fresh data from the active spreadsheet, then
 * rezips and returns as base64. All other template content is untouched.
 * @return {string} Base64 encoded XLSX data.
 */
function getExportData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. Retrieve the saved template file ───────────────────────────────────
  const props = PropertiesService.getScriptProperties();
  const templateFileId = props.getProperty(CONFIG.scriptProperties.templateFileId);

  if (!templateFileId) {
    throw new Error('No import template found. Please run an import first.');
  }

  const templateBlob = DriveApp.getFileById(templateFileId).getBlob();

  // ── 2. Unzip the template ─────────────────────────────────────────────────
  const zipEntries = Utilities.unzip(templateBlob);

  // ── 3. Build a map of filename → blob index for easy lookup ──────────────
  const entryMap = {};
  zipEntries.forEach(function(entry, index) {
    entryMap[entry.getName()] = index;
  });

  // ── 4. For each export tab, get sheet data and replace the worksheet XML ──
  // The template's workbook.xml maps: sheet1 = Equipment Items, sheet2 = Meters.
  // We replace them in order matching CONFIG.Exporting.EXPORT_TABS.
  const tabNames = CONFIG.Exporting.EXPORT_TABS;
  const worksheetPaths = ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'];

  for (let i = 0; i < tabNames.length; i++) {
    const tabName = tabNames[i];
    const worksheetPath = worksheetPaths[i];

    const sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      throw new Error('Sheet "' + tabName + '" not found. Please ensure the tab exists.');
    }

    const entryIndex = entryMap[worksheetPath];
    if (entryIndex === undefined) {
      throw new Error('Worksheet entry "' + worksheetPath + '" not found in template. Please re-run an import.');
    }

    // Read all values (plain values only — no formulas, no formatting needed)
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const data = (lastRow > 0 && lastCol > 0)
      ? sheet.getRange(1, 1, lastRow, lastCol).getValues()
      : [];

    const worksheetXml = buildWorksheetXml(data);
    zipEntries[entryIndex] = Utilities.newBlob(worksheetXml, 'application/xml', worksheetPath);
  }

  // ── 5. Rezip all entries and return as base64 ─────────────────────────────
  const outputBlob = Utilities.zip(zipEntries, CONFIG.Exporting.EXPORT_FILE_NAME);
  return Utilities.base64Encode(outputBlob.getBytes());
}

/**
 * Builds a minimal but valid xlsx worksheet XML string from a 2D array of values.
 * Strings are written as inline strings (t="inlineStr") to avoid needing a
 * shared strings table. Numbers are written as plain numeric values. Dates are
 * converted to xlsx serial numbers. Booleans are written as xlsx booleans (t="b").
 * Empty cells are omitted entirely.
 * @param {Array<Array<*>>} data - 2D array of cell values (rows x columns).
 * @return {string} Complete worksheet XML string.
 */
function buildWorksheetXml(data) {
  const xmlns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
    ' mc:Ignorable="x14ac xr xr2 xr3"' +
    ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"' +
    ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"' +
    ' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"' +
    ' xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"';

  const numRows = data.length;
  const numCols = numRows > 0 ? data[0].length : 0;

  // Dimension ref e.g. "A1:E10"
  const dimRef = (numRows > 0 && numCols > 0)
    ? 'A1:' + colIndexToLetter(numCols) + numRows
    : 'A1';

  let rowsXml = '';

  for (let r = 0; r < numRows; r++) {
    const row = data[r];
    let cellsXml = '';

    for (let c = 0; c < row.length; c++) {
      const val = row[c];

      // Skip entirely empty cells
      if (val === null || val === undefined || val === '') continue;

      const cellRef = colIndexToLetter(c + 1) + (r + 1);
      cellsXml += buildCellXml(cellRef, val);
    }

    // Only write the row element if it has at least one cell
    if (cellsXml !== '') {
      rowsXml += '<row r="' + (r + 1) + '" spans="1:' + numCols + '">' + cellsXml + '</row>';
    }
  }

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<worksheet ' + xmlns + '>' +
    '<dimension ref="' + dimRef + '"/>' +
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>' +
    '<sheetData>' + rowsXml + '</sheetData>' +
    '</worksheet>';
}

/**
 * Builds the XML for a single cell given its reference and value.
 * Type detection order: Date -> Boolean -> Number -> String.
 * @param {string} cellRef - Cell reference e.g. "A1".
 * @param {*} val - The cell value.
 * @return {string} Cell XML fragment.
 */
function buildCellXml(cellRef, val) {
  // Date
  if (val instanceof Date) {
    const serial = dateToSerial(val);
    return '<c r="' + cellRef + '" s="1"><v>' + serial + '</v></c>';
  }

  // Boolean
  if (typeof val === 'boolean') {
    return '<c r="' + cellRef + '" t="b"><v>' + (val ? 1 : 0) + '</v></c>';
  }

  // Number
  if (typeof val === 'number') {
    return '<c r="' + cellRef + '"><v>' + val + '</v></c>';
  }

  // String — use inlineStr to avoid needing a sharedStrings table
  const escaped = val.toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

  return '<c r="' + cellRef + '" t="inlineStr"><is><t>' + escaped + '</t></is></c>';
}

/**
 * Converts a JS Date object to an xlsx date serial number.
 * xlsx uses days since 1900-01-01, with a legacy off-by-one for the
 * non-existent 1900-02-29 (Lotus 1-2-3 compatibility bug).
 * @param {Date} date - The date to convert.
 * @return {number} xlsx serial number.
 */
function dateToSerial(date) {
  const epoch = Date.UTC(1899, 11, 30); // Dec 30 1899 accounts for the Lotus bug
  const ms = date.getTime() - epoch;
  return ms / 86400000; // Convert ms to days
}

/**
 * Converts a 1-based column index to an xlsx column letter (e.g. 1 -> "A", 27 -> "AA").
 * @param {number} colIndex - 1-based column index.
 * @return {string} Column letter string.
 */
function colIndexToLetter(colIndex) {
  let letter = '';
  let n = colIndex;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

// EOF: Export.gs
