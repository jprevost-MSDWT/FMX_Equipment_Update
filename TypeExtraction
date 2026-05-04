/**
 * Downloads an FMX PDF via the user's browser, opens an upload dialog, 
 * extracts Names and Modules using a Streaming Text Lexer, and writes to "Type_Import".
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
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlScript)
    .setWidth(300)
    .setHeight(150);

  ui.showModalDialog(htmlOutput, 'Downloading Report...');
}

function showImportDialog() {
  // Updated to match the new HTML file name
  const template = HtmlService.createTemplateFromFile('TypeExtractionDiag');
  const html = template.evaluate()
    .setWidth(400)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import File from Computer');
}

function processUploadedPDF(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    let fileBlob;
    
    // Accept the file blob from the form
    if (formObject && formObject.fileInput) {
      fileBlob = Utilities.newBlob(formObject.fileInput.getBytes(), MimeType.PDF, "TempReport.pdf");
    } else if (typeof formObject === 'string' && formObject.includes(',')) {
      const base64Data = formObject.split(',')[1];
      fileBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.PDF, "TempReport.pdf");
    } else {
      throw new Error("Invalid file upload. Please ensure you attached a file.");
    }
    
    let tempDocId;
    const timestamp = new Date().getTime();
    
    // Run OCR via Drive API
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

    // =========================================================================
    // 1. GLOBAL SCRUBBER
    // Vaporize headers, pagination, icons, and quotes before they break the parser
    // =========================================================================
    rawText = rawText.replace(/Equipment Type Admin Settings/gi, '\n');
    rawText = rawText.replace(/--- PAGE \d+ ---/gi, '\n');
    rawText = rawText.replace(/Showing \d+[–-]\d+ of \d+ records/gi, '\n');
    rawText = rawText.replace(/Page \d+ of \d+/gi, '\n');
    rawText = rawText.replace(/[\uE000-\uF8FF\u25A0-\u25FF\u200B-\u200D\uFEFF☑]/g, '\n');
    rawText = rawText.replace(/["“”]/g, '');

    // Split text into a pure stream of raw chunks
    let lines = rawText.split(/[\r\n\t]+/).map(l => l.trim()).filter(l => l.length > 0);
    
    // Remove isolated table header words that survived the scrub
    lines = lines.filter(line => {
        let l = line.toLowerCase();
        return l !== 'name' && l !== 'names' && l !== 'modules' && l !== 'namesmodules' && l !== 'name modules' && l !== 'w' && l !== ',';
    });

    const extractedData = [];
    const seenRecords = new Set();
    let currentName = "";
    let currentModule = "";

    // Helper to safely commit a parsed row to the final array
    function flushRecord() {
        if (currentName) {
            let n = currentName.replace(/^,+|,+$/g, '').replace(/\s+/g, ' ').trim();
            let m = currentModule.replace(/^,+|,+$/g, '').replace(/\s+/g, ' ').trim();

            // Fix commas dropped by OCR (e.g. "Maintenance requestPlanned maintenance")
            m = m.replace(/request(?=[A-Z])/gi, 'request, ');
            m = m.replace(/maintenance(?=[A-Z])/gi, 'maintenance, ');
            m = m.replace(/,(\s*,)+/g, ',').trim();

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

    // =========================================================================
    // 2. STREAMING LEXER WITH STRICT DICTIONARY
    // Uses your exact list of modules to flawlessly identify cuts
    // =========================================================================
    
    // The strict dictionary of every module FMX generates 
    const knownModules = [
        "Contractor Badge-Key Tracking request",
        "Fleet Maintenance request",
        "Inventory request",
        "Maintenance request",
        "Planned maintenance",
        "Project request",
        "Technology request", // Kept from your previous output
        "Transportation request" // Kept from your previous output
    ];
    
    // Dynamically build a Regex pattern from the dictionary to catch glued lines
    const moduleRegexStr = knownModules.map(m => m.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')).join('|');
    const gluedPattern = new RegExp(`^(.*?)\\s*((?:${moduleRegexStr})(?:,\\s*|$).*)`, 'i');

    for (let i = 0; i < lines.length; i++) {
        let line = lines[i];
        
        // A. Check for Glued Name + Module
        let splitMatch = line.match(gluedPattern);
        if (splitMatch) {
            let n = splitMatch[1].trim();
            let m = splitMatch[2].trim();

            if (n) {
                n = n.replace(/^,/, '').trim();
                // FMX Hierarchy Rule: Any line with a ">" is ALWAYS a brand new item.
                if (n.includes('>') && currentName && !currentModule) {
                    flushRecord();
                    currentName = n;
                } else if (currentName && !currentModule) {
                    currentName += " " + n; // Wrap continuation
                } else {
                    flushRecord();
                    currentName = n;
                }
            }
            
            if (currentModule) {
                currentModule += ", " + m;
            } else {
                currentModule = m;
            }
            continue;
        }

        // B. Check for Pure Module Chunks
        let checkLine = line.replace(/^,/, '').trim();
        let isMod = false;
        
        // Check if the line is just a placeholder dash/none
        if (/^none$|^\(none\)$|^-$/i.test(checkLine)) {
            if (!checkLine.includes('>')) isMod = true;
        } else {
            // Check against our strict dictionary keywords
            let lowerCheck = checkLine.toLowerCase();
            const modKeywords = ["fleet", "maintenance", "planned", "inventory", "contractor badge", "technology", "transportation", "project", "request"];
            for (let kw of modKeywords) {
                if (lowerCheck.includes(kw) && !lowerCheck.includes('>')) {
                    isMod = true;
                    break;
                }
            }
        }

        if (isMod) {
            if (currentModule) {
                // Determine if we need to append with a comma or just a space to fix broken wraps
                let segments = currentModule.split(',');
                let lastSeg = segments[segments.length - 1].trim().toLowerCase();
                let isComplete = lastSeg.endsWith('request') || lastSeg.endsWith('planned maintenance') || lastSeg === 'none' || lastSeg === '(none)' || lastSeg === '-';
                
                if (isComplete && !checkLine.toLowerCase().startsWith('request')) {
                    currentModule += ", " + checkLine;
                } else {
                    currentModule += " " + checkLine;
                }
            } else {
                currentModule = checkLine;
            }
        } else {
            // C. Must be a Name Chunk!
            if (currentModule) {
                // If we were reading modules and now hit a name, the previous item is complete.
                flushRecord();
            }
            
            let cleanName = line.replace(/^,/, '').trim();
            if (cleanName.includes('>')) {
                flushRecord();
                currentName = cleanName;
            } else {
                if (currentName) {
                    currentName += " " + cleanName;
                } else {
                    currentName = cleanName;
                }
            }
        }
    }
    flushRecord(); // Catch the final item

    DriveApp.getFileById(tempDocId).setTrashed(true);
    
    // =========================================================================
    // 3. WRITE TO SPREADSHEET
    // =========================================================================
    if (extractedData.length > 0) {
      const targetRange = ss.getRangeByName("Type_Import");
      
      if (targetRange) {
        targetRange.clearContent();
        const startCell = targetRange.getCell(1, 1);
        const sheet = startCell.getSheet();
        const outputRange = sheet.getRange(
          startCell.getRow(), 
          startCell.getColumn(), 
          extractedData.length, 
          extractedData[0].length
        );
        outputRange.setValues(extractedData);
        ss.toast(`Extracted ${extractedData.length} records successfully!`, "Success", 5);
        return true; 
      } else {
        throw new Error("Could not find the named range 'Type_Import'.");
      }
    } else {
      throw new Error("No data was extracted. The file might not match the expected format.");
    }
    
  } catch (error) {
    throw new Error(error.message || error.toString()); 
  }
}
