const SPREADSHEET_ID = '1Q8YgtLChQ2dLTem-Tz2_NelDrCiJwDeJYzRMEBSCrSI'; 
const FOLDER_ID = '1s4i2xXEbzsawTmsVPWHRSpgVkiNVB3KO';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Bus Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSettingsFromServer() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Settings");
    const settings = { busNumber: [], driverName: [], instName: [], tripList: [] };
    if (!sheet) return settings;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) settings.busNumber.push(data[i][0]);
      if (data[i][1]) settings.driverName.push(data[i][1]);
      if (data[i][2]) settings.instName.push(data[i][2]);
      if (data[i][3]) settings.tripList.push(data[i][3]);
    }
    return settings;
  } catch(e) { return { busNumber: [], driverName: [], instName: [], tripList: [] }; }
}

function getLastKM(busNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  let lastKM = 0;
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === "Settings") continue;
    let data = sheets[i].getDataRange().getValues();
    for (let j = data.length - 1; j >= 1; j--) {
      if (data[j][7] == busNo) {
        lastKM = parseFloat(data[j][12]) || 0;
        return lastKM;
      }
    }
  }
  return lastKM;
}

function process(data, image) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = data.nepMonthName.replace(/\s+/g, ""); 
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const header = ["मिति (BS)", "बार", "मिति (AD)", "प्रकार", "ट्रिप", "सिफ्ट", "संस्था/रुट", "बस नं", "ड्राइभर", "लिटर", "रेट", "डिजेल रकम", "आजको KM", "चलेको KM", "रिजर्भ रकम", "बैना/खर्च", "बचत", "कुल डिजल लिटर", "कुल डिजल रकम", "कुल रिजर्भ बचत", "विवरण", "फोटो"];
      sheet.appendRow(header);
      sheet.getRange(1, 1, 1, header.length).setBackground("#2ecc71").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1, header.length, 110);
    }

    let currentSav = parseFloat(data.resAmt - data.resExp) || 0;
    let photoUrl = "";
    if (image && image.base64) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const blob = Utilities.newBlob(Utilities.base64Decode(image.base64), image.mimeType, "IMG_" + Date.now());
      photoUrl = folder.createFile(blob).getUrl();
    }

    const rowData = [
      data.nepDateRaw, data.nepDay, data.engDate, 
      (data.entryType === 'Institution' ? 'संस्था' : 'रिजर्भ'),
      data.trip, data.shift,
      (data.entryType === 'Institution' ? data.instName : data.fromLoc + " - " + data.toLoc),
      data.busNumber, data.driverName, data.dLiter, data.dRate, 
      data.dAmount, data.currentKM, (data.currentKM - data.prevKM),
      data.resAmt, data.resExp, currentSav,
      "", "", "", // Running totals calculations can be added here
      data.remarks, photoUrl ? "फोटो हेर्नुहोस्" : "फोटो छैन"
    ];

    sheet.appendRow(rowData);
    let lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, rowData.length).setHorizontalAlignment("center").setVerticalAlignment("middle");

    if (data.entryType === 'Reserve') {
       sheet.getRange(lastRow, 1, 1, 22).setFontColor("#ef4444").setFontWeight("bold");
    }

    if (photoUrl) {
      sheet.getRange(lastRow, 22).setFormula(`=HYPERLINK("${photoUrl}","फोटो हेर्नुहोस्")`);
    }

    return "SUCCESS";
  } catch (e) { return "Error: " + e.toString(); }
}
