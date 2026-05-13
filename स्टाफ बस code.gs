const SPREADSHEET_ID = '1Q8YgtLChQ2dLTem-Tz2_NelDrCiJwDeJYzRMEBSCrSI'; 
const FOLDER_ID = '1s4i2xXEbzsawTmsVPWHRSpgVkiNVB3KO';

function doGet() {
  // एप खोल्ने बित्तिकै Settings सिट चेक गर्ने र बनाउने
  ensureSettingsSheet(); 
  
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Bus Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// १. Settings सिट अनिवार्य बनाउने र फर्म्याट गर्ने
function ensureSettingsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Settings");
  
  if (!sheet) {
    sheet = ss.insertSheet("Settings");
    const headers = ["बस नम्बर", "ड्राइभर", "संस्था", "ट्रिप"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, 4)
         .setBackground("#3498db")
         .setFontColor("white")
         .setFontWeight("bold")
         .setHorizontalAlignment("center");
    formatMySheet(sheet);
  }
}

// २. सेटिङबाट डाटा तान्ने (Dropdown को लागि)
function getSettings() {
  ensureSettingsSheet();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Settings");
  const data = sheet.getDataRange().getValues();
  
  const settings = { busNumber: [], driverName: [], instName: [], tripList: [] };
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) settings.busNumber.push(data[i][0]);
    if (data[i][1]) settings.driverName.push(data[i][1]);
    if (data[i][2]) settings.instName.push(data[i][2]);
    if (data[i][3]) settings.tripList.push(data[i][3]);
  }
  return settings;
}

// ३. नयाँ सेटिङ थप्दा ब्याकअप राख्ने
function updateGlobalSettings(key, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Settings");
  if (!sheet) { ensureSettingsSheet(); sheet = ss.getSheetByName("Settings"); }

  const data = sheet.getDataRange().getValues();
  let colIndex = ['busNumber','driverName','instName','tripList'].indexOf(key);

  // डुप्लिकेट चेक
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex] == value) return "EXISTS";
  }

  // खाली सेल खोज्ने वा नयाँ रो थप्ने
  let rowToUse = data.length + 1;
  for (let i = 1; i < data.length; i++) {
    if (!data[i][colIndex]) {
      rowToUse = i + 1;
      break;
    }
  }

  sheet.getRange(rowToUse, colIndex + 1).setValue(value);
  formatMySheet(sheet);
  return "SUCCESS";
}

// ४. मुख्य डाटा इन्ट्री प्रोसेस
function process(data, image) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = data.nepMonthName.replace(/\s+/g, ""); 
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const header = ["मिति (BS)", "बार", "मिति (AD)", "प्रकार", "संस्था/रुट", "बस नं", "ड्राइभर", "लिटर", "रेट", "डिजेल रकम", "आजको KM", "चलेको KM", "रिजर्भ रकम", "बैना/खर्च", "बचत", "कुल डिजल लिटर", "कुल डिजल रकम", "कुल रिजर्भ बचत", "विवरण", "फोटो", "KEY"];
      sheet.appendRow(header);
      sheet.getRange(1, 1, 1, header.length).setBackground("#2ecc71").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    let photoUrl = "";
    let photoLabel = "फोटो छैन";

    if (image && image.base64) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const fileName = "IMG_" + data.nepDateRaw.replace(/\//g, '-') + "_" + data.busNumber;
      const blob = Utilities.newBlob(Utilities.base64Decode(image.base64), image.mimeType, fileName);
      photoUrl = folder.createFile(blob).getUrl();
      photoLabel = "फोटो हेर्नुहोस्";
    }

    const rowData = [
      data.nepDateRaw, data.nepDay, data.engDate, 
      (data.entryType === 'Institution' ? 'संस्था' : 'रिजर्भ'),
      (data.entryType === 'Institution' ? data.instName : data.fromLoc + " - " + data.toLoc),
      data.busNumber, data.driverName, data.dLiter, data.dRate, 
      (data.dLiter * data.dRate), data.currentKM, data.runKM || 0,
      data.resAmt, data.resExp, (data.resAmt - data.resExp),
      data.dLiter, (data.dLiter * data.dRate), (data.entryType === 'Reserve' ? (data.resAmt - data.resExp) : "-"),
      data.remarks, photoLabel, data.key || ""
    ];

    sheet.appendRow(rowData);
    const lastRow = sheet.getLastRow();
    
    if (photoUrl !== "") {
      sheet.getRange(lastRow, 20).setFormula(`=HYPERLINK("${photoUrl}","${photoLabel}")`);
    }

    // फर्म्याटिङ लागू गर्ने
    formatMySheet(sheet);

    return "SUCCESS";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

// ५. मास्टर फर्म्याटिङ फङ्सन (नछोपिने र ग्याप राख्ने गरी)
function formatMySheet(sheet) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return;

  const range = sheet.getRange(1, 1, lastRow, lastCol);
  
  // सेटिङहरू: र्‍याप अन, एलाइनमेन्ट सेन्टर
  range.setWrap(true); 
  range.setVerticalAlignment("middle");
  range.setHorizontalAlignment("center");
  range.setFontFamily("Mukta");

  // पहिले अटो रिसाइज गर्ने
  sheet.autoResizeColumns(1, lastCol);
  
  // त्यसपछि प्रत्येक कोलममा ठूलो ग्याप (Padding) र मिनिमम साइज दिने
  for (let i = 1; i <= lastCol; i++) {
    let currentWidth = sheet.getColumnWidth(i);
    // कम्तिमा १२०px चौडाइ र थप ५०px को ग्याप (तपाईँले भने जस्तै खुला बनाउन)
    let newWidth = Math.max(currentWidth + 50, 120);
    sheet.setColumnWidth(i, newWidth);
  }

  // विशेष कोलमहरूका लागि अझ ठूलो चौडाइ (विवरण जस्तै)
  if (lastCol >= 19) {
    sheet.setColumnWidth(5, 200);  // संस्था/रुट
    sheet.setColumnWidth(19, 250); // विवरण (सबैभन्दा खुला)
    sheet.setColumnWidth(20, 140); // फोटो
  }
  
  // हेडर रोलाई अझ प्रस्ट बनाउने
  sheet.getRange(1, 1, 1, lastCol).setFontSize(11).setWrap(true);
}
