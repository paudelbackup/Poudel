// १. तपाईंको विवरणहरू यहाँ राख्नुहोस्
const SPREADSHEET_ID = "13zSjJyoojvkchunfE_wLBckWQlP0Q6hFjHlxFlhbcGA";
const FOLDER_ID = "1FSlCdPa2H8GVR74jSMf5q_MrwmOa4mA2";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('बस व्यवस्थापन प्रणाली')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// सेटिङहरू तान्ने (ड्रपडाउन र गतेको लागि)
function getSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Setting") || ss.insertSheet("Setting");
  const data = sheet.getDataRange().getValues();
  let s = { inst: "", loan: "", repair: "", staff: "", pers: "", dateOff: 0 };
  data.forEach(r => { if(r[0]) s[r[0]] = r[1]; });
  return s;
}

// महिना अनुसार नयाँ शिट बनाउने र नेपाली हेडर राख्ने फङ्सन
function getMonthlySheet(nepDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const parts = nepDate.split(" "); // "२०८३ बैशाख १२ शनिबार" बाट "२०८३" र "बैशाख" निकाल्छ
  const sheetName = parts[0] + " " + parts[1]; // शिटको नाम: "२०८३ बैशाख"
  
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // नेपाली हेडरहरू
    const headers = [["मिति", "अङ्ग्रेजी मिति", "विवरण प्रकार", "नाम/संस्था", "थप सावाँ/बिल", "किस्ता/तिरेको", "कूल बाँकी", "कैफियत", "फोटो", "ब्याज दर%"]];
    
    // हेडर डिजाइन: रङ्गिन, बोल्ड र सेन्टर
    sheet.getRange(1, 1, 1, 10).setValues(headers)
         .setBackground("#ffb400")
         .setFontWeight("bold")
         .setHorizontalAlignment("center")
         .setVerticalAlignment("middle");
    
    sheet.setFrozenRows(1); // पहिलो लाइन फ्रिज गर्ने
    sheet.setColumnWidths(1, 10, 130); // कोलमको चौडाई मिलाउने
  }
  return sheet;
}

// नाम अनुसारको पुरानो हिसाब (Summary) तान्ने - Super Fast
function getSummary(name, cat) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const summarySheet = ss.getSheetByName("Summary") || ss.insertSheet("Summary");
  const data = summarySheet.getDataRange().getValues();
  
  // अन्तिमबाट खोज्दै जाने (छिटो हुन्छ)
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][3] == name && data[i][2] == cat) {
      return { 
        lastDate: data[i][0], 
        sawa: parseFloat(data[i][4]) || 0, 
        byaj: 0, // यहाँ थप ब्याज लोजिक राख्न सकिन्छ
        total: parseFloat(data[i][6]) || 0, 
        rate: parseFloat(data[i][9]) || 12,
        count: i 
      };
    }
  }
  return { lastDate: "-", sawa: 0, byaj: 0, total: 0, rate: 12, count: 0 };
}

// डाटा सेभ गर्ने मुख्य फङ्सन
function processEntry(obj) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const summarySheet = ss.getSheetByName("Summary") || ss.insertSheet("Summary");
  const dataSheet = getMonthlySheet(obj.nepDate); // महिना अनुसारको शिट लिने
  
  let photoLink = "फोटो छैन";
  
  // १. फोटो अपलोड लोजिक
  if (obj.imageBlob && obj.imageBlob.includes("base64")) {
    try {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const contentType = obj.imageBlob.split(":")[1].split(";")[0];
      const bytes = Utilities.base64Decode(obj.imageBlob.split(",")[1]);
      const blob = Utilities.newBlob(bytes, contentType, obj.name + "_" + obj.nepDate + ".jpg");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoLink = '=HYPERLINK("' + file.getUrl() + '", "फोटो हेर्नुहोस्")';
    } catch (e) {
      photoLink = "अपलोड त्रुटि";
    }
  }

  // २. शिटमा बस्ने डाटाको लाइन (Row)
  const rowData = [
    obj.nepDate,                                  // A: मिति
    obj.engDate,                                  // B: AD मिति
    obj.cat,                                      // C: विवरण प्रकार
    obj.name,                                     // D: नाम/संस्था
    obj.cat === "ऋण" ? obj.addSawa : obj.billAmt, // E: थप रकम
    obj.cat === "ऋण" ? (obj.paidSawa + obj.kista) : obj.payAmt, // F: तिरेको
    obj.totalAmt,                                 // G: कूल बाँकी
    obj.remarks,                                  // H: कैफियत
    photoLink,                                    // I: फोटो लिङ्क
    obj.rate                                      // J: ब्याज दर
  ];

  // ३. Monthly Sheet मा डाटा थप्ने र फम्र्याटिङ गर्ने
  dataSheet.appendRow(rowData);
  const lastRow = dataSheet.getLastRow();
  const range = dataSheet.getRange(lastRow, 1, 1, 10);
  
  // डाटा सेन्टर गर्ने र र्‍याप (Wrap) गर्ने ताकि टेक्स्ट नछोपियोस्
  range.setHorizontalAlignment("center")
       .setVerticalAlignment("middle")
       .setWrap(true);
  dataSheet.setRowHeight(lastRow, 45); // पर्याप्त ग्यापको लागि

  // ४. Summary Sheet अपडेट गर्ने (ताकि अर्को पटक झट्टै हिसाब आओस्)
  const sData = summarySheet.getDataRange().getValues();
  let found = false;
  for(let i=1; i<sData.length; i++){
    if(sData[i][3] === obj.name && sData[i][2] === obj.cat) {
      summarySheet.getRange(i+1, 1, 1, 10).setValues([rowData]);
      found = true; 
      break;
    }
  }
  if(!found) summarySheet.appendRow(rowData);

  return "डाटा सफलतापूर्वक सुरक्षित गरियो!";
}

// सेटिङ अपडेट गर्ने (नाम थप्ने वा गते मिलाउने)
function updateSettings(type, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Setting") || ss.insertSheet("Setting");
  const data = sheet.getDataRange().getValues();
  let found = false;
  for(let i=0; i<data.length; i++){
    if(data[i][0] === type) { 
      sheet.getRange(i+1, 2).setValue(value); 
      found = true; 
      break; 
    }
  }
  if(!found) sheet.appendRow([type, value]);
}
