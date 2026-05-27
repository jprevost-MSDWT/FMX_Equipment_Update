/*
Project Name: FMX Equipment Import non-Gem
Project Version: 5.00
Filename: Export.gs
File Version: 2.16
Chat link: [Insert Link]
*/

/**
 * @file Export.gs
 * @description Handles file export. Writes data into the stored Google Sheet
 *              template, exports it as xlsx, then patches in the FMX metadata
 *              files (custom.xml, core.xml, app.xml, _rels/.rels) that Google
 *              strips during export. Metadata is stored as base64 in Script
 *              Properties at import time and injected via a pure-JS zip writer.
 *
 * NOTE: The FMX custom properties in docProps/custom.xml contain two version values:
 *   FmxImportTemplate = 3
 *   FmxTemplateVersion = 12
 * These are captured from the original import file. If FMX ever rejects the upload
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
 * Writes current tab data into the stored Google Sheet template, exports it
 * as xlsx, then patches in the FMX metadata files stored at import time.
 * @return {string} Base64 encoded XLSX data.
 */
function getExportData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  // ── 1. Retrieve stored IDs and metadata ───────────────────────────────────
  const templateSheetId = props.getProperty(CONFIG.scriptProperties.templateSheetId);
  if (!templateSheetId) {
    throw new Error('No template sheet found. Please run an import first.');
  }

  const metadataKeys = CONFIG.scriptProperties.metadataFiles;
  const metadataFiles = {};
  for (const key of Object.keys(metadataKeys)) {
    const val = props.getProperty(metadataKeys[key]);
    if (!val) {
      throw new Error('Missing metadata "' + key + '". Please re-run an import.');
    }
    metadataFiles[key] = val; // stored as base64
  }

  // ── 2. Write current tab data into the stored Google Sheet ────────────────
  const templateSs = SpreadsheetApp.openById(templateSheetId);
  const tabNames = CONFIG.Exporting.EXPORT_TABS;

  for (const tabName of tabNames) {
    const sourceSheet = ss.getSheetByName(tabName);
    if (!sourceSheet) {
      throw new Error('Sheet "' + tabName + '" not found. Please ensure the tab exists.');
    }

    const targetSheet = templateSs.getSheetByName(tabName);
    if (!targetSheet) {
      throw new Error('Tab "' + tabName + '" not found in template sheet. Please re-run an import.');
    }

    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();

    if (lastRow > 0 && lastCol > 0) {
      const data = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
      targetSheet.clearContents();
      targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    } else {
      targetSheet.clearContents();
    }
  }

  SpreadsheetApp.flush();
  Utilities.sleep(3000);

  // ── 3. Export the Google Sheet as xlsx ────────────────────────────────────
  const token = ScriptApp.getOAuthToken();
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + templateSheetId + '/export?format=xlsx';

  let response;
  let lastError = null;
  const maxAttempts = 3;
  const retryDelay = 2000;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      response = UrlFetchApp.fetch(exportUrl, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      lastError = null;
      if (response.getResponseCode() === 200) break;
      if (attempt < maxAttempts) {
        console.warn('Export attempt ' + attempt + ' failed (code: ' + response.getResponseCode() + '). Retrying...');
        Utilities.sleep(retryDelay);
      }
    } catch (networkError) {
      lastError = networkError;
      if (attempt < maxAttempts) {
        console.warn('Export attempt ' + attempt + ' threw a network error: ' + networkError.message + '. Retrying...');
        Utilities.sleep(retryDelay);
      }
    }
  }

  if (lastError) {
    throw new Error('Export failed after ' + maxAttempts + ' attempts: ' + lastError.message);
  }
  if (response.getResponseCode() !== 200) {
    throw new Error('Export failed. Code: ' + response.getResponseCode() + '. Response: ' + response.getContentText().substring(0, 500));
  }

  // ── 4. Patch FMX metadata into the exported xlsx ──────────────────────────
  // Utilities.unzip requires content type to be application/zip regardless
  // of what Google's export endpoint returns in the Content-Type header.
  const exportedBlob = response.getBlob();
  exportedBlob.setContentType('application/zip');
  const patchedBytes = patchXlsxMetadata(exportedBlob, metadataFiles);

  return Utilities.base64Encode(patchedBytes);
}

/**
 * Patches FMX metadata files into an xlsx blob using native GAS zip methods.
 * Unzips the exported xlsx, removes any existing entries for the metadata paths,
 * appends the correct metadata entries from Script Properties, and rezips.
 * All blobs passed to Utilities.zip() originate from Utilities.unzip() or are
 * created with 'application/octet-stream' content type, which GAS accepts.
 * @param {GoogleAppsScript.Base.Blob} zipBlob - The exported xlsx blob (content type must be application/zip).
 * @param {Object} metadataFiles - Map of zip path -> base64 encoded file content.
 * @return {number[]} Patched xlsx as a GAS byte array.
 */
function patchXlsxMetadata(zipBlob, metadataFiles) {
  const entries = Utilities.unzip(zipBlob);
  const pathsToReplace = new Set(Object.keys(metadataFiles));
  const newBlobs = [];

  // Keep all existing entries except the ones being replaced
  for (const entry of entries) {
    if (!pathsToReplace.has(entry.getName())) {
      newBlobs.push(entry);
    }
  }

  // Append replacement metadata entries
  for (const path of Object.keys(metadataFiles)) {
    const bytes = Utilities.base64Decode(metadataFiles[path]);
    newBlobs.push(Utilities.newBlob(bytes, 'application/octet-stream', path));
  }

  return Utilities.zip(newBlobs, 'patched.xlsx').getBytes();
}

// EOF: Export.gs
