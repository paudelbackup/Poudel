const SPREADSHEET_ID = "13zSjJyoojvkchunfE_wLBckWQlP0Q6hFjHlxFlhbcGA";
const FOLDER_ID = "1FSlCdPa2H8GVR74jSMf5q_MrwmOa4mA2";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('बस व्यवस्थापन प्रणाली')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
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

function getMonthlySheet(nepDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const parts = nepDate.split(" "); 
  const sheetName = parts[0] + " " + parts[1];
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // ब्याज दरलाई कैफियत (८ औं नम्बर) भन्दा अगाडि (७ औं नम्बरमा) राखिएको छ
    const headers = [["मिति", "अङ्ग्रेजी मिति", "विवरण प्रकार", "नाम/संस्था", "थप सावाँ/बिल", "किस्ता/तिरेको", "ब्याज दर%", "कूल बाँकी", "कैफियत", "फोटो"]];
    sheet.getRange(1, 1, 1, 10).setValues(headers)
         .setBackground("#ffb400").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSummary(name, cat) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const summarySheet = ss.getSheetByName("सारांश") || ss.insertSheet("सारांश");
  const data = summarySheet.getDataRange().getValues();
  
  // नाम र विवरण प्रकार दुवै मिलेको अन्तिम रेकर्ड खोज्ने
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][3].toString().trim() === name.toString().trim() && data[i][2].toString().trim() === cat.toString().trim()) {
      return { 
        lastDate: data[i][0], 
        sawa: parseFloat(data[i][4]) || 0, 
        total: parseFloat(data[i][7]) || 0, // कुल बाँकी अब ७ औं इन्डेक्समा छ
        rate: parseFloat(data[i][6]) || 12,
        count: i 
      };
    }
  }
  return { lastDate: "-", sawa: 0, total: 0, rate: 12, count: 0 };
}

function processEntry(obj) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const summarySheet = ss.getSheetByName("सारांश") || ss.insertSheet("सारांश");
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

  // विवरण प्रकार खाली नजाओस् भनेर सुनिश्चित गरिएको छ
  const rowData = [
    obj.nepDate,
    obj.engDate,
    obj.cat, 
    obj.name,
    obj.cat === "ऋण" ? obj.addSawa : obj.billAmt,
    obj.cat === "ऋण" ? (obj.paidSawa + obj.kista) : obj.payAmt,
    obj.rate,       // ब्याज दर (कैफियत भन्दा अगाडि)
    obj.totalAmt,   // कूल बाँकी
    obj.remarks,    // कैफियत
    photoLink       // फोटो
  ];

  dataSheet.appendRow(rowData);
  const lastRow = dataSheet.getLastRow();
  dataSheet.getRange(lastRow, 1, 1, 10).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(false);
  
  for (let col = 1; col <= 10; col++) {
    dataSheet.autoResizeColumn(col);
    dataSheet.setColumnWidth(col, dataSheet.getColumnWidth(col) + 15);
  }

  // सारांश अपडेट
  const sData = summarySheet.getDataRange().getValues();
  let found = false;
  for(let i=1; i<sData.length; i++){
    if(sData[i][3].toString().trim() === obj.name.toString().trim() && sData[i][2].toString().trim() === obj.cat.toString().trim()) {
      summarySheet.getRange(i+1, 1, 1, 10).setValues([rowData]);
      found = true; 
      break;
    }
  }
  if(!found) summarySheet.appendRow(rowData);

  return "सफल भयो";
}

function updateSettings(type, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("सेटिङ") || ss.insertSheet("सेटिङ");
  const data = sheet.getDataRange().getValues();
  let found = false;
  for(let i=0; i<data.length; i++){
    if(data[i][0] === type) { sheet.getRange(i+1, 2).setValue(value); found = true; break; }
  }
  if(!found) sheet.appendRow([type, value]);
}
