const SPREADSHEET_ID = "13zSjJyoojvkchunfE_wLBckWQlP0Q6hFjHlxFlhbcGA";
const FOLDER_ID = "1FSlCdPa2H8GVR74jSMf5q_MrwmOa4mA2";

function getNepalDate() {
  let now = new Date();
  let nepalTime = new Date(now.getTime() + (now.getTimezoneOffset() + 345) * 60000);
  return nepalTime;
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('बस व्यवस्थापन प्रणाली')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("सेटिङ") || ss.insertSheet("सेटिङ");
  const data = sheet.getDataRange().getValues();
  let s = { inst: "", loan: "", repair: "", staff: "", pers: "", dateOff: 0 };
  data.forEach(r => { if(r[0]) s[r[0]] = r[1]; });
  return s;
}

function getSummary(name, cat, currentEngDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  let allEntries = [];
  
  sheets.forEach(sheet => {
    if (["सेटिङ", "सारांश"].includes(sheet.getName())) return;
    let data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][3]?.toString().trim() === name.toString().trim() && 
          data[i][2]?.toString().trim() === cat.toString().trim()) {
        allEntries.push({
          nepDate: data[i][0],
          engDate: new Date(data[i][1]),
          addSawa: parseFloat(data[i][4]) || 0,
          paidKista: parseFloat(data[i][5]) || 0,
          paidByaj: parseFloat(data[i][6]) || 0,
          rate: parseFloat(data[i][7]) || 12
        });
      }
    }
  });

  allEntries.sort((a, b) => a.engDate - b.engDate);
  let runningSawa = 0, totalInterestAccrued = 0, lastDate = null, currentRate = 12, lastK = 0, lastTransactionAmt = 0;

  allEntries.forEach(entry => {
    if (lastDate) {
      let days = Math.floor((entry.engDate - lastDate) / (1000 * 60 * 60 * 24));
      if (days > 0) totalInterestAccrued += (runningSawa * (currentRate / 100) * days) / 365;
    }
    runningSawa += (entry.addSawa - entry.paidKista);
    totalInterestAccrued -= entry.paidByaj;
    currentRate = entry.rate;
    lastDate = entry.engDate;
    if(entry.paidKista > 0) lastK = entry.paidKista;
    lastTransactionAmt = entry.addSawa !== 0 ? entry.addSawa : entry.paidKista;
  });

  let selectedDate = new Date(currentEngDate);
  if (lastDate && selectedDate > lastDate) {
    let days = Math.floor((selectedDate - lastDate) / (1000 * 60 * 60 * 24));
    totalInterestAccrued += (runningSawa * (currentRate / 100) * days) / 365;
  }

  return {
    lastDate: allEntries.length > 0 ? allEntries[allEntries.length-1].nepDate : "-",
    sawa: runningSawa,
    accruedInterest: Math.round(totalInterestAccrued),
    rate: currentRate,
    count: allEntries.length,
    lastPaidKista: lastK,
    lastTransaction: lastTransactionAmt,
    total: Math.round(runningSawa + totalInterestAccrued)
  };
}

function processEntry(obj) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dataSheet = getMonthlySheet(obj.nepDate);
  let photoLink = "फोटो छैन";
  
  if (obj.imageBlob && obj.imageBlob.includes("base64")) {
    try {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const contentType = obj.imageBlob.split(":")[1].split(";")[0];
      const bytes = Utilities.base64Decode(obj.imageBlob.split(",")[1]);
      const blob = Utilities.newBlob(bytes, contentType, obj.name + "_" + obj.nepDate + ".jpg");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoLink = '=HYPERLINK("' + file.getUrl() + '", "फोटो हेर्नुहोस्")';
    } catch (e) { photoLink = "त्रुटि"; }
  }

  dataSheet.appendRow([obj.nepDate, obj.engDate, obj.cat, obj.name, obj.addSawa_or_Bill, obj.paidKista_or_Pay, obj.paidByaj, obj.rate, obj.totalAmt, obj.remarks, photoLink]);
  const lastRow = dataSheet.getLastRow();
  if (lastRow > 1) {
    dataSheet.getRange(2, 1, lastRow - 1, dataSheet.getLastColumn()).sort({column: 2, ascending: true});
  }
  formatSheet(dataSheet);
  return "SUCCESS";
}

function formatSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1) return;
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  sheet.autoResizeColumns(1, lastCol);
  for (let i = 1; i <= lastCol; i++) { sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 30); }
  range.setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(false);
  sheet.getRange(1, 1, 1, lastCol).setBackground("#ffb400").setFontWeight("bold").setFontColor("#000000");
}

function getMonthlySheet(nepDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const parts = nepDate.split(" ");
  const sheetName = parts[0] + " " + parts[1];
  let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["मिति", "अङ्ग्रेजी मिति", "प्रकार", "नाम/संस्था", "थप सावाँ/बिल", "किस्ता/तिरेको", "ब्याज तिरेको", "दर%", "कुल बाँकी", "कैफियत", "फोटो"]);
  }
  return sheet;
}

function updateSettings(t, v) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const s = ss.getSheetByName("सेटिङ") || ss.insertSheet("सेटिङ");
  const d = s.getDataRange().getValues();
  for(let i=0; i<d.length; i++) {
    if(d[i][0] === t) { s.getRange(i+1, 2).setValue(v); formatSheet(s); return; }
  }
  s.appendRow([t, v]);
  formatSheet(s);
}
