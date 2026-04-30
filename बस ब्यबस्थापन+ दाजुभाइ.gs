/**
 * बस व्यवस्थापन प्रणाली - Code.gs
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('बस व्यवस्थापन प्रणाली')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function processEntry(obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = obj.sheetYearMonth; 
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = ["मिति (BS)", "मिति (AD)", "प्रकार", "नाम", "ब्याज दर %", "थप सावाँ/बिल", "किस्ता/भुक्तानी", "तिरेको ब्याज", "बाँकी मौज्दात", "कैफियत", "फोटो"];
    sheet.appendRow(headers);
  }

  const photoStatus = obj.imageBlob ? saveFile(obj) : "फोटो छैन";

  sheet.appendRow([
    obj.nepDate, obj.engDate, obj.cat, obj.name, obj.rate, 
    obj.addSawa_or_Bill, obj.paidKista_or_Pay, obj.paidByaj, 
    obj.totalAmt, obj.remarks, photoStatus
  ]);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    range.sort([{column: 2, ascending: true}]); 
  }

  formatMySheet(sheet);
  return "SUCCESS";
}

function getSummary(name, cat, selectedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let sawa = 0, accruedInterest = 0, count = 0, lastDateAD = "-", lastDateBS = "-", lastRate = 0, lastK = 0;
  const selDate = new Date(selectedDate);
  
  let allEntries = [];

  sheets.forEach(sheet => {
    if(sheet.getName() === "Settings" || sheet.getName() === "सेटिङ") return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // दाजुभाइको हकमा रोहित वा कुमारले गरेको सबै कारोबारको समरी हेर्न
      if (data[i][2] === cat && data[i][3].includes(name)) {
        allEntries.push(data[i]);
      }
    }
  });

  allEntries.sort((a, b) => new Date(a[1]) - new Date(b[1]));

  allEntries.forEach(row => {
    let rowDate = new Date(row[1]);
    if (rowDate > selDate) return;

    let currentRate = parseFloat(row[4]) || 0;
    lastRate = currentRate;
    
    if (lastDateAD !== "-" && (cat === "ऋण" || cat === "व्यक्तिगत")) {
      let days = Math.floor((rowDate - new Date(lastDateAD)) / (1000 * 60 * 60 * 24));
      if (days > 0) accruedInterest += (sawa * (currentRate / 100) * days) / 365;
    }
    
    let tS = parseFloat(row[5]) || 0, pK = parseFloat(row[6]) || 0, pB = parseFloat(row[7]) || 0;

    if (cat === "व्यक्तिगत") {
      let totalPaidThisTime = pK + pB;
      if(totalPaidThisTime > 0) lastK = totalPaidThisTime; 
      sawa += tS; accruedInterest -= pB; sawa -= pK; 
    } 
    else if (cat === "ऋण") {
      if(pK > 0) lastK = pK; else if(tS > 0) lastK = tS;
      sawa += tS; accruedInterest -= pB; sawa -= (pK - pB);
    } 
    else { 
      lastK = pK > 0 ? pK : (tS > 0 ? tS : 0);
      sawa += (tS - pK); 
    }
    
    lastDateAD = row[1]; lastDateBS = row[0]; count++;
  });

  if (lastDateAD !== "-" && (cat === "ऋण" || cat === "व्यक्तिगत")) {
    let finalDays = Math.floor((selDate - new Date(lastDateAD)) / (1000 * 60 * 60 * 24));
    if (finalDays > 0) accruedInterest += (sawa * (lastRate / 100) * finalDays) / 365;
  }

  return { 
    sawa: Math.round(sawa), accruedInterest: Math.round(accruedInterest), 
    total: Math.round(sawa + accruedInterest), count: count, 
    lastDate: lastDateBS, rate: lastRate, lastK: Math.round(lastK) 
  };
}

function formatMySheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0) return;
  sheet.getRange(1, 1, lastRow, lastCol).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontFamily("Mukta").setFontSize(10).setWrap(false);
  sheet.autoResizeColumns(1, lastCol);
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  }
}

function saveFile(obj) {
  try {
    let folder, folders = DriveApp.getFoldersByName("Bus_Management_Photos");
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Bus_Management_Photos");
    const blob = Utilities.newBlob(Utilities.base64Decode(obj.imageBlob.split(',')[1]), "image/jpeg", obj.name + "_" + obj.nepDate + ".jpg");
    const file = folder.createFile(blob);
    return '=HYPERLINK("' + file.getUrl() + '", "फोटो हेर्नुस्")';
  } catch(e) { return "Error"; }
}

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sSheet = ss.getSheetByName("सेटिङ") || ss.insertSheet("सेटिङ");
  const data = sSheet.getDataRange().getValues();
  let settings = {};
  for(let i=1; i<data.length; i++) { settings[data[i][0]] = data[i][1]; }
  return settings;
}

function updateSettings(k, v) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sSheet = ss.getSheetByName("सेटिङ") || ss.insertSheet("सेटिङ");
  const data = sSheet.getDataRange().getValues();
  let found = false;
  for(let i=1; i<data.length; i++) { if(data[i][0] == k) { sSheet.getRange(i+1, 2).setValue(v); found = true; break; } }
  if(!found) { sSheet.appendRow([k, v]); }
  return "SUCCESS";
}
