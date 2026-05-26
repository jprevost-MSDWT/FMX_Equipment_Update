/*
Project Name: FMX Equipment Import non-Gem
Project Version: 5.00
Filename: Export.gs
File Version: 2.14
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
  const exportedBytes = response.getBlob().getBytes();
  const patchedBytes = patchXlsxMetadata(exportedBytes, metadataFiles);

  return Utilities.base64Encode(patchedBytes);
}

/**
 * Patches FMX metadata files into an xlsx byte array.
 * Parses the zip structure of the exported xlsx, removes any existing entries
 * for the metadata paths, appends the correct metadata entries from the stored
 * template, and rebuilds a valid zip binary.
 * @param {number[]} zipBytes - The exported xlsx as a GAS byte array.
 * @param {Object} metadataFiles - Map of zip path -> base64 encoded file content.
 * @return {number[]} Patched xlsx as a GAS byte array.
 */
function patchXlsxMetadata(zipBytes, metadataFiles) {
  // Decode the zip and extract all entries except the ones we're replacing
  const entries = parseZipEntries(zipBytes);
  const pathsToReplace = new Set(Object.keys(metadataFiles));

  // Filter out entries that will be replaced
  const keptEntries = entries.filter(function(e) {
    return !pathsToReplace.has(e.path);
  });

  // Build new entries from stored metadata (base64 → bytes)
  const newEntries = Object.keys(metadataFiles).map(function(path) {
    const bytes = Utilities.base64Decode(metadataFiles[path]);
    return { path: path, bytes: bytes };
  });

  // Rebuild the zip with kept + new entries
  return buildZip(keptEntries.concat(newEntries));
}

/**
 * Parses a zip file byte array and returns an array of entry objects.
 * Only reads the local file entries (not the central directory).
 * @param {number[]} bytes - Raw zip bytes.
 * @return {Array<{path: string, bytes: number[]}>} Array of extracted entries.
 */
function parseZipEntries(bytes) {
  const entries = [];
  let i = 0;

  while (i < bytes.length - 4) {
    // Local file header signature: PK\x03\x04
    if (bytes[i] !== 0x50 || bytes[i+1] !== 0x4B || bytes[i+2] !== 0x03 || bytes[i+3] !== 0x04) {
      break;
    }

    const compression   = (bytes[i+8]  & 0xFF) | ((bytes[i+9]  & 0xFF) << 8);
    const compressedSize   = (bytes[i+18] & 0xFF) | ((bytes[i+19] & 0xFF) << 8) |
                             ((bytes[i+20] & 0xFF) << 16) | ((bytes[i+21] & 0xFF) << 24);
    const uncompressedSize = (bytes[i+22] & 0xFF) | ((bytes[i+23] & 0xFF) << 8) |
                             ((bytes[i+24] & 0xFF) << 16) | ((bytes[i+25] & 0xFF) << 24);
    const fileNameLen   = (bytes[i+26] & 0xFF) | ((bytes[i+27] & 0xFF) << 8);
    const extraLen      = (bytes[i+28] & 0xFF) | ((bytes[i+29] & 0xFF) << 8);

    const headerEnd = i + 30 + fileNameLen + extraLen;

    // Decode filename
    let path = '';
    for (let j = i + 30; j < i + 30 + fileNameLen; j++) {
      path += String.fromCharCode(bytes[j] & 0xFF);
    }

    // Extract compressed data
    const compressedData = bytes.slice(headerEnd, headerEnd + compressedSize);

    // Decompress if needed
    let entryBytes;
    if (compression === 0) {
      // Stored (no compression)
      entryBytes = compressedData;
    } else if (compression === 8) {
      // Deflate — use GAS Utilities to decompress
      // Wrap in a minimal zlib stream for Utilities.ungzip compatibility:
      // GAS doesn't have raw inflate, so we re-wrap as a gzip blob via zip round-trip
      entryBytes = inflateRaw(compressedData, uncompressedSize);
    } else {
      throw new Error('Unsupported zip compression type ' + compression + ' for entry: ' + path);
    }

    entries.push({ path: path, bytes: entryBytes });
    i = headerEnd + compressedSize;
  }

  return entries;
}

/**
 * Decompresses raw deflate-compressed bytes using a GAS-compatible approach.
 * Wraps the raw deflate data in a minimal zip structure so Utilities.unzip
 * can decompress it, then extracts the result.
 * @param {number[]} compressedBytes - Raw deflate compressed data.
 * @param {number} uncompressedSize - Expected size after decompression.
 * @return {number[]} Decompressed bytes.
 */
function inflateRaw(compressedBytes, uncompressedSize) {
  // Build a minimal zip file containing just this one entry with a dummy name,
  // then use Utilities.unzip to decompress it.
  const dummyName = 'x';
  const nameBytes = [dummyName.charCodeAt(0)];
  const nameLen = nameBytes.length;

  // CRC32 placeholder (0) — Utilities.unzip doesn't validate CRC
  const crc = [0, 0, 0, 0];
  const compSize = intToLeBytes(compressedBytes.length);
  const uncompSize = intToLeBytes(uncompressedSize);

  // Local file header
  const localHeader = [
    0x50, 0x4B, 0x03, 0x04, // signature
    0x14, 0x00,             // version needed
    0x00, 0x00,             // flags
    0x08, 0x00,             // compression: deflate
    0x00, 0x00, 0x00, 0x00, // mod time/date
  ].concat(crc)
   .concat(compSize)
   .concat(uncompSize)
   .concat([nameLen, 0x00, 0x00, 0x00]) // filename len, extra len
   .concat(nameBytes);

  const localOffset = intToLeBytes(0);

  // Central directory entry
  const centralEntry = [
    0x50, 0x4B, 0x01, 0x02, // signature
    0x14, 0x00,             // version made by
    0x14, 0x00,             // version needed
    0x00, 0x00,             // flags
    0x08, 0x00,             // compression
    0x00, 0x00, 0x00, 0x00, // mod time/date
  ].concat(crc)
   .concat(compSize)
   .concat(uncompSize)
   .concat([nameLen, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00])
   .concat(localOffset)
   .concat(nameBytes);

  const centralSize = intToLeBytes(centralEntry.length);
  const centralOffset = intToLeBytes(localHeader.length + compressedBytes.length);

  // End of central directory
  const eocd = [
    0x50, 0x4B, 0x05, 0x06, // signature
    0x00, 0x00,             // disk number
    0x00, 0x00,             // disk with CD
    0x01, 0x00,             // entries on disk
    0x01, 0x00,             // total entries
  ].concat(centralSize)
   .concat(centralOffset)
   .concat([0x00, 0x00]); // comment length

  const miniZip = localHeader
    .concat(Array.from(compressedBytes))
    .concat(centralEntry)
    .concat(eocd);

  const miniBlob = Utilities.newBlob(miniZip, 'application/zip', 'mini.zip');
  const unzipped = Utilities.unzip(miniBlob);
  return unzipped[0].getBytes();
}

/**
 * Converts a 32-bit integer to a 4-byte little-endian array.
 * @param {number} n - Integer to convert.
 * @return {number[]} 4-byte little-endian array.
 */
function intToLeBytes(n) {
  return [
    (n & 0xFF),
    (n >> 8)  & 0xFF,
    (n >> 16) & 0xFF,
    (n >> 24) & 0xFF
  ];
}

/**
 * Builds a valid zip file byte array from an array of entry objects.
 * Uses Utilities.zip() by constructing properly typed blobs from unzip output,
 * working around GAS content-type restrictions by building the zip binary manually.
 * @param {Array<{path: string, bytes: number[]}>} entries - Entries to include.
 * @return {number[]} Complete zip file as a byte array.
 */
function buildZip(entries) {
  const localHeaders = [];
  const centralDirectory = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = [];
    for (let i = 0; i < entry.path.length; i++) {
      nameBytes.push(entry.path.charCodeAt(i) & 0xFF);
    }
    const nameLen = nameBytes.length;

    // Compress the entry bytes using a mini zip round-trip
    const compressedBytes = deflateBytes(entry.bytes);
    const crc = computeCrc32(entry.bytes);
    const crcBytes = intToLeBytes(crc);
    const compSize = intToLeBytes(compressedBytes.length);
    const uncompSize = intToLeBytes(entry.bytes.length);
    const offsetBytes = intToLeBytes(offset);

    // Local file header
    const localHeader = [
      0x50, 0x4B, 0x03, 0x04,
      0x14, 0x00,
      0x00, 0x00,
      0x08, 0x00,             // deflate
      0x00, 0x00, 0x00, 0x00,
    ].concat(crcBytes)
     .concat(compSize)
     .concat(uncompSize)
     .concat([nameLen & 0xFF, (nameLen >> 8) & 0xFF])
     .concat([0x00, 0x00])
     .concat(nameBytes);

    localHeaders.push(localHeader.concat(Array.from(compressedBytes)));
    offset += localHeader.length + compressedBytes.length;

    // Central directory entry
    const cdEntry = [
      0x50, 0x4B, 0x01, 0x02,
      0x14, 0x00,
      0x14, 0x00,
      0x00, 0x00,
      0x08, 0x00,
      0x00, 0x00, 0x00, 0x00,
    ].concat(crcBytes)
     .concat(compSize)
     .concat(uncompSize)
     .concat([nameLen & 0xFF, (nameLen >> 8) & 0xFF])
     .concat([0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00])
     .concat(offsetBytes)
     .concat(nameBytes);

    centralDirectory.push(cdEntry);
  }

  // Flatten local headers
  const localSection = [].concat(...localHeaders);
  const cdSection = [].concat(...centralDirectory);
  const cdSize = intToLeBytes(cdSection.length);
  const cdOffset = intToLeBytes(localSection.length);

  const eocd = [
    0x50, 0x4B, 0x05, 0x06,
    0x00, 0x00,
    0x00, 0x00,
    entries.length & 0xFF, (entries.length >> 8) & 0xFF,
    entries.length & 0xFF, (entries.length >> 8) & 0xFF,
  ].concat(cdSize)
   .concat(cdOffset)
   .concat([0x00, 0x00]);

  return localSection.concat(cdSection).concat(eocd);
}

/**
 * Compresses bytes using deflate by creating a mini zip and extracting
 * the compressed payload from it.
 * @param {number[]} bytes - Raw bytes to compress.
 * @return {number[]} Deflate-compressed bytes.
 */
function deflateBytes(bytes) {
  const blob = Utilities.newBlob(bytes, 'application/octet-stream', 'entry');
  const zipped = Utilities.zip([blob], 'temp.zip');
  const zipBytes = zipped.getBytes();

  // Local file header is 30 + filename length bytes
  // filename is 'entry' = 5 chars, so data starts at offset 35
  // Read compressed size from header bytes 18-21
  const compressedSize = (zipBytes[18] & 0xFF) | ((zipBytes[19] & 0xFF) << 8) |
                         ((zipBytes[20] & 0xFF) << 16) | ((zipBytes[21] & 0xFF) << 24);
  const fileNameLen = (zipBytes[26] & 0xFF) | ((zipBytes[27] & 0xFF) << 8);
  const extraLen    = (zipBytes[28] & 0xFF) | ((zipBytes[29] & 0xFF) << 8);
  const dataStart   = 30 + fileNameLen + extraLen;

  return zipBytes.slice(dataStart, dataStart + compressedSize);
}

/**
 * Computes a CRC32 checksum for a byte array.
 * Required for valid zip local file headers.
 * @param {number[]} bytes - Bytes to checksum.
 * @return {number} CRC32 value.
 */
function computeCrc32(bytes) {
  const table = makeCrc32Table();
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) {
    crc = (crc >>> 8) ^ table[(crc ^ bytes[i]) & 0xFF];
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

/**
 * Builds the CRC32 lookup table.
 * @return {number[]} 256-entry CRC32 table.
 */
function makeCrc32Table() {
  const table = [];
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[i] = c;
  }
  return table;
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
