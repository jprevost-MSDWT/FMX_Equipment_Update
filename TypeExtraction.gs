/*
Project Name: FMX Equipment Import non-Gem
Project Version: 4.00
Filename: TypeExtraction_Update.gs
File Version: 2.12
Chat link: [Insert Link]
*/

/**
 * CONFIGURATION: Update these lists to "teach" the Lexer about your FMX setup.
 * =========================================================================
 */

const FMX_KNOWN_MODULES = [
  "Contractor Badge-Key Tracking request",
  "Fleet Maintenance request",
  "Inventory request",
  "Maintenance request",
  "Planned maintenance",
  "Project request",
  "Technology request",
  "Transportation request"
];

const FMX_MODULE_KEYWORDS = [
  "fleet", 
  "maintenance", 
  "planned", 
  "inventory", 
  "contractor badge", 
  "technology", 
  "transportation", 
  "project", 
  "request"
];

const FMX_ROOT_CATEGORIES = [
  "*Testing Type",
  "Building Key",
  "Contractor Badge",
  "Contractor Key",
  "Custodial",
  "Electronics",
  "Elevators & Lifts",
  "Life Safety",
  "MEP",
  "Test Equipt",
  "Unused",
  "Vehicle"
];

/**
 * Main logic starts here.
 * =========================================================================
 */

function extractDataFromPDF() {
  const ui = SpreadsheetApp.getUi();
  const fileUrl = 'https://warrenk12.gofmx.com/admin-settings/equipment-types?format=pdf&useOnlySelectedColumns=False';
  
  const htmlScript = `
    <script>
      window.open('${fileUrl}', '_blank');
      google.script.run
        .withSuccessHandler(google.script.host.close)
        .showImportDialog();
    </script>
    <div style="font-family: Arial; padding: 20px; text-align: center;">
      Triggering secure download...
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlScript).setWidth(300).setHeight(150);
  ui.showModalDialog(htmlOutput, 'Downloading Report...');
}

function showImportDialog() {
  const template = HtmlService.createTemplateFromFile('TypeExtractionDiag');
  const html = template.evaluate().setWidth(400).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import File from Computer');
}

function processUploadedPDF(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tempDocId = null; 
  
  try {
    let fileBlob;
    if (formObject && formObject.fileInput) {
      fileBlob = Utilities.newBlob(formObject.fileInput.getBytes(), MimeType.PDF, "TempReport.pdf");
    } else if (typeof formObject === 'string' && formObject.includes(',')) {
      const base64Data = formObject.split(',')[1];
      fileBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.PDF, "TempReport.pdf");
    } else {
      throw new Error("Invalid file upload.");
    }
    
    const timestamp = new Date().getTime();
    
    // --- RESTORED: Drive API v2 / v3 Compatibility Block ---
    try {
      const resource = { title: "Temp OCR Doc - " + timestamp, mimeType: MimeType.GOOGLE_DOCS };
      const tempFile = Drive.Files.insert(resource, fileBlob, {ocr: true});
      tempDocId = tempFile.id;
    } catch(e) {
      try {
        const resource = { name: "Temp OCR Doc - " + timestamp, mimeType: MimeType.GOOGLE_DOCS };
        const tempFile = Drive.Files.create(resource, fileBlob);
        tempDocId = tempFile.id;
      } catch (e2) {
        throw new Error("Drive API Error. Please verify the Drive API is enabled in Services.");
      }
    }
    
    if (!tempDocId) throw new Error("Upload failed: Could not retrieve a valid Document ID.");
    
    const doc = DocumentApp.openById(tempDocId);
    let rawText = doc.getBody().getText();

    // 1. GLOBAL SCRUBBER
    rawText = rawText.replace(/Equipment Type Admin Settings/gi, '\n');
    rawText = rawText.replace(/--- PAGE \d+ ---/gi, '\n');
    rawText = rawText.replace(/Showing \d+[–-]\d+ of \d+ records/gi, '\n');
    rawText = rawText.replace(/Page \d+ of \d+/gi, '\n');
    rawText = rawText.replace(/[\uE000-\uF8FF\u25A0-\u25FF\u200B-\u200D\uFEFF☑]/g, '\n');
    rawText = rawText.replace(/["“”]/g, '');

    // --- RESTORED: Un-glue Squashed Headers ---
    rawText = rawText.replace(/NamesModules/gi, '\n');
    rawText = rawText.replace(/Name Modules/gi, '\n');
    rawText = rawText.replace(/Name\s*(?=\*Testing Type)/gi, '\n');

    const safeModulesStr = FMX_KNOWN_MODULES.map(m => m.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')).join('|');
    const modStartRegex = new RegExp(`([a-zA-Z0-9\\)])(${safeModulesStr})`, 'g');
    rawText = rawText.replace(modStartRegex, '$1\n$2');

    let lines = rawText.split(/[\r\n\t]+/).map(l => l.trim()).filter(l => l.length > 0);
    lines = lines.filter(l => !['name', 'names', 'modules', 'namesmodules', 'name modules', 'w', ','].includes(l.toLowerCase()));

    const extractedData = [];
    const seenRecords = new Set();
    let currentName = "";
    let currentModule = "";

    function flushRecord() {
        if (currentName) {
            let n = currentName.replace(/^,+|,+$/g, '').replace(/\s+/g, ' ').trim();
            let m = currentModule.replace(/^,+|,+$/g, '').replace(/\s+/g, ' ').trim();

            // --- RECOVERY LOGIC: Check if modules are stuck in the Name ---
            for (let known of FMX_KNOWN_MODULES) {
                // If name ends exactly with a known module
                if (n.toLowerCase().endsWith(known.toLowerCase())) {
                    let cutIndex = n.toLowerCase().lastIndexOf(known.toLowerCase());
                    let moveText = n.substring(cutIndex).trim();
                    n = n.substring(0, cutIndex).trim();
                    m = moveText + (m ? ", " + m : "");
                }
            }

            if (n) {
                let key = n + "|||" + m;
                if (!seenRecords.has(key)) {
                    seenRecords.add(key);
                    extractedData.push([n, m]);
                }
            }
        }
        currentName = "";
        currentModule = "";
    }

    const gluedPattern = new RegExp(`^(.*?)\\s+((?:${safeModulesStr})(?:,\\s*|$).*)`, 'i');

    for (let i = 0; i < lines.length; i++) {
        let line = lines[i];
        let checkLine = line.replace(/^,/, '').trim();
        
        let splitMatch = line.match(gluedPattern);
        if (splitMatch) {
            let nPart = splitMatch[1].trim();
            let mPart = splitMatch[2].trim();
            
            if (nPart.endsWith(',') || FMX_KNOWN_MODULES.some(m => m.toLowerCase() === nPart.toLowerCase())) {
               currentModule += (currentModule ? ", " : "") + line;
               continue;
            }

            if (currentName && currentModule) flushRecord();
            currentName = (currentName ? currentName + " " : "") + nPart;
            currentModule = (currentModule ? currentModule + ", " : "") + mPart;
            continue;
        }

        let isMod = false;
        let isRoot = FMX_ROOT_CATEGORIES.some(root => checkLine.startsWith(root));
        
        if (!isRoot) {
            if (/^none$|^\(none\)$|^-$/i.test(checkLine)) {
                isMod = true;
            } else {
                let low = checkLine.toLowerCase();
                isMod = FMX_MODULE_KEYWORDS.some(kw => low.includes(kw)) && !checkLine.includes('>');
            }
        }

        if (isMod) {
            currentModule = (currentModule ? currentModule + ", " : "") + checkLine;
        } else {
            let isCont = /^[a-z\(]/.test(checkLine) || checkLine.startsWith('>');
            
            if (currentName && isCont) {
                currentName += " " + checkLine;
            } else {
                if (currentName) flushRecord();
                currentName = checkLine;
            }
        }
    }
    flushRecord();

    if (extractedData.length > 0) {
      const targetRange = ss.getRangeByName("Type_Import");
      targetRange.clearContent();
      const sheet = targetRange.getCell(1,1).getSheet();
      sheet.getRange(targetRange.getRow(), targetRange.getColumn(), extractedData.length, 2).setValues(extractedData);
      ss.toast(`Extracted ${extractedData.length} records!`, "Success");
    }
  } catch (e) {
    throw e;
  } finally {
    // --- UPDATED: Robust temporary file cleanup ---
    if (tempDocId) {
      try {
        DriveApp.getFileById(tempDocId).setTrashed(true);
      } catch (cleanupError) {
        console.warn("Temporary file cleanup failed: " + tempDocId);
      }
    }
  }
}

// EOF: TypeExtraction_Update.gs
