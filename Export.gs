/*
Project Name: FMX Equipment Import non-Gem
Project Version: 5.00
Filename: Export.gs
File Version: 2.12
Chat link: [Insert Link]
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
 * Uses a temporary file to combine specific tabs into one export.
 * @return {string} Base64 encoded XLSX data.
 */

function getExportData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. Retrieve the saved template file ID ───────────────────────────────
  const props = PropertiesService.getScriptProperties();
  const templateFileId = props.getProperty(CONFIG.scriptProperties.templateFileId);

  if (!templateFileId) {
    throw new Error('No import template found. Please run an import first.');
  }

  // ── 2. Open the saved template as a Google Sheet ─────────────────────────
  // The template was saved as xlsx during import; opening it via Sheets API
  // converts it temporarily while preserving the original file's structure
  // (including FMX custom properties) for the final export fetch.
  const templateFile = DriveApp.getFileById(templateFileId);
  const tempSs = SpreadsheetApp.open(templateFile);
  const tempSsId = tempSs.getId();

  try {
    // ── 3. Get source sheets from the active spreadsheet ───────────────────
    const tabNames = CONFIG.Exporting.EXPORT_TABS;
    for (const name of tabNames) {
      const sheet = ss.getSheetByName(name);
      if (!sheet) {
        throw new Error('Sheet "' + name + '" not found. Please ensure the tab exists.');
      }

      // Find the matching tab in the template spreadsheet and overwrite its data
      const targetSheet = tempSs.getSheetByName(name);
      if (!targetSheet) {
        throw new Error('Tab "' + name + '" not found in the saved template. Please re-run an import.');
      }

      const sourceData = sheet.getDataRange().getValues();
      targetSheet.clearContents();
      targetSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
    }

    SpreadsheetApp.flush();
    Utilities.sleep(3000);

    // ── 4. Export the updated template as xlsx ─────────────────────────────
    const url = `https://docs.google.com/spreadsheets/d/${tempSsId}/export?format=xlsx`;
    const fetchOptions = {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };

    let response;
    let lastError = null;
    const maxAttempts = 3;
    const retryDelay = 2000;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        response = UrlFetchApp.fetch(url, fetchOptions);
        lastError = null;
        if (response.getResponseCode() === 200) break;
        if (attempt < maxAttempts) {
          console.warn(`Export attempt ${attempt} failed (code: ${response.getResponseCode()}). Retrying...`);
          Utilities.sleep(retryDelay);
        }
      } catch (networkError) {
        lastError = networkError;
        if (attempt < maxAttempts) {
          console.warn(`Export attempt ${attempt} threw a network error: ${networkError.message}. Retrying...`);
          Utilities.sleep(retryDelay);
        }
      }
    }

    if (lastError) {
      throw new Error(`Export failed after ${maxAttempts} attempts due to a network error: ${lastError.message}`);
    }

    if (response.getResponseCode() !== 200) {
      let responseSnippet = '';
      try {
        responseSnippet = response.getContentText().substring(0, 500);
      } catch (e) {
        responseSnippet = '[response body could not be decoded]';
      }
      throw new Error(
        `Failed to fetch the export from Google servers after ${maxAttempts} attempts. ` +
        `Code: ${response.getResponseCode()}. Response: ${responseSnippet}`
      );
    }

    return Utilities.base64Encode(response.getBlob().getBytes());

  } finally {
    // The template file stays in Drive — do NOT trash it.
    // Only close out any in-memory references.
  }
}

// EOF: Export.gs
