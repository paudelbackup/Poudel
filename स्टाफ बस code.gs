const FOLDER_ID = "1s4i2xXEbzsawTmsVPWHRSpgVkiNVB3KO"; 
const SPREADSHEET_ID = "1Q8YgtLChQ2dLTem-Tz2_NelDrCiJwDeJYzRMEBSCrSI";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- सेटिङ व्यवस्थापन (Dropdowns) ---
function getSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Add Setting") || ss.insertSheet("Add Setting");
  const data = sheet.getDataRange().getValues();
  const settings = { busNumber: [], driverName: [], instName: [] };
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) settings.busNumber.push(data[i][0]);
    if (data[i][1]) settings.driverName.push(data[i][1]);
    if (data[i][2]) settings.instName.push(data[i][2]);
  }
  return settings;
}

// --- नम्बर कन्भर्टर (नेपाली <=> अंग्रेजी) ---
function toEngNum(n) {
  if (!n) return "0";
  const nepDigits = {'०':'0','१':'1','२':'2','३':'3','४':'4','५':'5','६':'6','७':'7','८':'8','९':'9'};
  return n.toString().replace(/[०-९]/g, d => nepDigits[d]);
}

function toNepNum(n) {
  if (!n) return "";
  const nepDigits = ['०','१','२','३','४','५','६','७','८','९'];
  return n.toString().replace(/\d/g, d => nepDigits[d]);
}

// --- मुख्य डेटा इन्ट्री (Process) ---
function process(data, photoObj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = data.nepMonthName; // जस्तै: २०८३ जेठ
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    // इमेज र कोलम अर्डर अनुसारका हेडर्स (A to U)
    const headers = [
      "मिति (BS)", "मिति (AD)", "बार", "बस नं", "संस्था/रुट", "ट्रिप", 
      "सिफ्ट", "ड्राइभर", "लिटर", "रेट", "डिजल रकम", "स्टार्ट KM", 
      "आजको KM", "चलेको KM", "कहाँबाट", "कहाँसम्म", "भाडा", "खर्च", 
      "बचत", "कैफियत", "फोटो"
    ];
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight("bold")
           .setBackground("#10b981") // Green Header
           .setFontColor("white")
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle");
      sheet.setFrozenRows(1);
    }

    const lastRow = sheet.getLastRow();

    // फोटो अपलोड लजिक
    let photoLink = "फोटो छैन";
    if (photoObj && photoObj.base64) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const blob = Utilities.newBlob(Utilities.base64Decode(photoObj.base64), photoObj.mimeType, photoObj.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoLink = '=HYPERLINK("' + file.getUrl() + '", "फोटो हेर्नुहोस्")';
    }

    // डेटा तयारी (Calculations)
    const curLit = parseFloat(toEngNum(data.dLiter)) || 0;
    const curRat = parseFloat(toEngNum(data.dRate)) || 0;
    const curAmt = curLit * curRat;
    
    const startKM = parseFloat(toEngNum(data.lastKM)) || 0;
    const todayKM = parseFloat(toEngNum(data.currentKM)) || 0;
    const drivenKM = (todayKM > startKM) ? (todayKM - startKM) : 0;

    const resAmt = parseFloat(toEngNum(data.resAmt)) || 0;
    const resExp = parseFloat(toEngNum(data.resExp)) || 0;
    const resSav = resAmt - resExp;

    const instOrRoute = (data.entryType === "Institution" ? data.instName : (data.fromLoc + " - " + data.toLoc));

    // सिटको कोलम (A to U) अनुसार डेटा बाध्ने
    const rowData = [
      toNepNum(data.nepDateRaw), // A
      data.engDate,               // B
      data.nepDay,                // C
      data.busNumber,             // D
      instOrRoute,                // E
      data.trip || "-",           // F
      data.shift || "-",          // G
      data.driverName,            // H
      curLit,                     // I
      curRat,                     // J
      curAmt,                     // K
      startKM,                    // L
      todayKM,                    // M
      drivenKM,                   // N
      data.fromLoc || "-",        // O
      data.toLoc || "-",          // P
      resAmt,                     // Q
      resExp,                     // R
      resSav,                     // S
      data.remarks || "-",        // T
      photoLink                   // U
    ];

    // रो एड गर्ने
    sheet.appendRow(rowData);
    const nLastRow = sheet.getLastRow();
    
    // --- फर्म्याटिङ: सेन्टर, ग्याप र उचाइ कन्ट्रोल ---
    const rowRange = sheet.getRange(nLastRow, 1, 1, headers.length);
    rowRange.setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setWrap(false); // डाटा लामो भए पनि उचाइ नबढ्ने, दायाँबायाँ मात्र जाने
    
    sheet.setRowHeight(nLastRow, 35); // प्रत्येक रोको फिक्स्ड उचाइ

    // कोलमको चौडाइ मिलाउने (अटो फिट + एक्स्ट्रा ग्याप)
    sheet.autoResizeColumns(1, headers.length);
    for (let col = 1; col <= headers.length; col++) {
      let currentWidth = sheet.getColumnWidth(col);
      sheet.setColumnWidth(col, currentWidth + 40); // छेउछाउमा स्पष्ट ग्याप
    }

    return "SUCCESS";
  } catch (e) { 
    return "Error: " + e.toString(); 
  } finally { 
    lock.releaseLock(); 
  }
}
