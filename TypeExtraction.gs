/*
Project Name: FMX Equipment Import non-Gem
Project Version: 4.00
Filename: TypeExtraction.gs
File Version: 2.14
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
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tempDocId = null;
  let extractedData = [];
  
  try {
    const response = UrlFetchApp.fetch(fileUrl, {
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error("Failed to fetch PDF. HTTP Response Code: " + response.getResponseCode());
    }

    const blob = response.getBlob();
    
    // =========================================================================
    // CRITICAL API V3 WARNING - DO NOT REVERT
    // Google Apps Script now defaults to Drive API v3. 
    // Do NOT revert back to 'title' (use 'name') or 'Drive.Files.insert()' (use 'create').
    // Reverting will cause a "Drive.Files.insert is not a function" error.
    // =========================================================================
    
    // IMPORTANT: Make sure the "Drive API" is added in the Services panel (+ icon) of your Apps Script editor.
    const resource = {
      name: 'FMX_Temp_Extraction', // v3 REQUIRES 'name' (do not change back to 'title')
      mimeType: MimeType.GOOGLE_DOCS
    };
    
    // v3 REQUIRES 'create' (do not change back to 'insert')
    const tempDocFile = Drive.Files.create(resource, blob);
    tempDocId = tempDocFile.id;
    
    const tempDoc = DocumentApp.openById(tempDocId);
    const text = tempDoc.getBody().getText();
    const lines = text.split('\n');
    
    let currentName = "";
    let currentModule = "";
    
    function flushRecord() {
        if (currentName) {
            extractedData.push([currentName.trim(), currentModule.trim()]);
        }
        currentName = "";
        currentModule = "";
    }
    
    for (let i = 0; i < lines.length; i++) {
        let checkLine = lines[i].trim();
        if (!checkLine) continue;

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
      const sheet = targetRange.getCell(1, 1).getSheet();
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
         console.warn("Could not delete temp file: " + tempDocId, cleanupError);
      }
    }
  }
}
