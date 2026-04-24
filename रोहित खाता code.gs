const SPREADSHEET_ID = "13zSjJyoojvkchunfE_wLBckWQlP0Q6hFjHlxFlhbcGA"; 
const FOLDER_ID = "1FSlCdPa2H8GVR74jSMf5q_MrwmOa4mA2"; 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('बस व्यवस्थापन')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Setting") || ss.insertSheet("Setting");
  const data = sheet.getDataRange().getValues();
  let s = { inst: "", loan: "", repair: "", staff: "", pers: "", dateOff: 0 };
  data.forEach(r => { if(r[0]) s[r[0]] = r[1]; });
  return s;
}

function updateSettings(type, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Setting");
  const data = sheet.getDataRange().getValues();
  let found = false;
  for(let i=0; i<data.length; i++){
    if(data[i][0] === type) { 
      sheet.getRange(i+1, 2).setValue(value); 
      found = true; break; 
    }
  }
  if(!found) sheet.appendRow([type, value]);
  return "SUCCESS";
}

function getSummary(name, cat) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Summary");
  if(!sheet) return { oldDate: "-", oldSawa: 0, total: 0 };
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][3] == name && data[i][2] == cat) {
      return { oldDate: data[i][0], oldSawa: parseFloat(data[i][4]) || 0, total: parseFloat(data[i][6]) || 0 };
    }
  }
  return { oldDate: "-", oldSawa: 0, total: 0 };
}

function processEntry(obj) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dataSheet = ss.getSheetByName("Data") || ss.insertSheet("Data");
  const summarySheet = ss.getSheetByName("Summary") || ss.insertSheet("Summary");
  
  let imageUrl = "";
  if (obj.imageBlob) {
    try {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const contentType = obj.imageBlob.substring(obj.imageBlob.indexOf(":") + 1, obj.imageBlob.indexOf(";"));
      const bytes = Utilities.base64Decode(obj.imageBlob.split(",")[1]);
      const blob = Utilities.newBlob(bytes, contentType, "Bus_" + Date.now() + ".jpg");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrl = file.getUrl();
    } catch (e) { imageUrl = "Error: " + e.toString(); }
  }

  const rowData = [obj.nepDate, new Date(), obj.category, obj.targetName, obj.sawaAmt, obj.byajAmt, obj.totalAmt, obj.remarks, imageUrl];
  dataSheet.appendRow(rowData);

  const sData = summarySheet.getDataRange().getValues();
  let found = false;
  for(let i=1; i<sData.length; i++){
    if(sData[i][3] === obj.targetName && sData[i][2] === obj.category) {
      summarySheet.getRange(i+1, 1).setValue(obj.nepDate);
      summarySheet.getRange(i+1, 5).setValue(obj.sawaAmt);
      summarySheet.getRange(i+1, 6).setValue(obj.byajAmt);
      summarySheet.getRange(i+1, 7).setValue(obj.totalAmt);
      found = true; break;
    }
  }
  if(!found) summarySheet.appendRow(rowData);

  return "डाटा सुरक्षित भयो!";
}
