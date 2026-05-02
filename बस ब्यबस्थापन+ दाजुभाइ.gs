/**
 * बस व्यवस्थापन प्रणाली - सुधारिएको Code.gs
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('बस व्यवस्थापन प्रणाली')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// १. डाटा सुरक्षित गर्ने फंक्सन
function processEntry(obj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = obj.sheetYearMonth; 
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    if (sheet.getLastRow() === 0) {
      const headers = ["मिति (BS)", "मिति (AD)", "प्रकार", "नाम", "ब्याज दर %", "थप सावाँ/बिल", "किस्ता/भुक्तानी", "तिरेको ब्याज", "कुल बाँकी", "कैफियत", "फोटो"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
    }

    const photoStatus = (obj.imageBlob && obj.imageBlob.includes(',')) ? saveFile(obj) : "फोटो छैन";

    sheet.appendRow([
      obj.nepDate, 
      obj.engDate, 
      obj.cat, 
      obj.name, 
      parseFloat(obj.rate) || 0, 
      parseFloat(obj.addSawa_or_Bill) || 0, 
      parseFloat(obj.paidKista_or_Pay) || 0, 
      parseFloat(obj.paidByaj) || 0, 
      parseFloat(obj.totalAmt) || 0, 
      obj.remarks, 
      photoStatus
    ]);

    formatMySheet(sheet);
    return "SUCCESS";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

// २. समरी र ब्याज गणना (सुधारिएको लोजिक)
function getSummary(name, cat, selectedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let sawa = 0, accruedInterest = 0, count = 0, lastDateAD = null, lastDateBS = "-", lastRate = 0, lastK = 0;
  const selDate = new Date(selectedDate);
  
  let allEntries = [];

  sheets.forEach(sheet => {
    if (["Settings", "सेटिङ"].includes(sheet.getName())) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // नाम र प्रकार मिल्नुपर्छ (Case-insensitive check)
      if (data[i][2] === cat && data[i][3].toString().trim() === name.trim()) {
        allEntries.push(data[i]);
      }
    }
  });

  // मिति अनुसार क्रमबद्ध गर्ने
  allEntries.sort((a, b) => new Date(a[1]) - new Date(b[1]));

  allEntries.forEach(row => {
    let rowDate = new Date(row[1]);
    if (rowDate > selDate) return;

    let currentRate = parseFloat(row[4]) || 0;
    
    // ब्याज गणना: अघिल्लो कारोबार देखि अहिले सम्मको
    if (lastDateAD && (cat === "ऋण" || cat === "व्यक्तिगत")) {
      let days = Math.floor((rowDate - lastDateAD) / (1000 * 60 * 60 * 24));
      if (days > 0) accruedInterest += (sawa * (lastRate / 100) * days) / 365;
    }
    
    let tS = parseFloat(row[5]) || 0; // थप सावाँ
    let pK = parseFloat(row[6]) || 0; // किस्ता/फिर्ता
    let pB = parseFloat(row[7]) || 0; // ब्याज तिरेको

    if (cat === "व्यक्तिगत" || cat === "ऋण") {
      lastK = pK > 0 ? pK : (tS > 0 ? tS : 0);
      sawa += (tS - pK);
      accruedInterest -= pB; // तिरेको ब्याज घटाउने
    } else { 
      lastK = pK > 0 ? pK : (tS > 0 ? tS : 0);
      sawa += (tS - pK); 
    }
    
    lastDateAD = rowDate; 
    lastDateBS = row[0]; 
    lastRate = currentRate;
    count++;
  });

  // अन्तिम कारोबार देखि आज सम्मको बाँकी ब्याज थप्ने
  if (lastDateAD && (cat === "ऋण" || cat === "व्यक्तिगत")) {
    let finalDays = Math.floor((selDate - lastDateAD) / (1000 * 60 * 60 * 24));
    if (finalDays > 0) accruedInterest += (sawa * (lastRate / 100) * finalDays) / 365;
  }

  return { 
    sawa: Math.round(sawa), 
    accruedInterest: Math.max(0, Math.round(accruedInterest)), // ब्याज माइनसमा नजाओस्
    total: Math.round(sawa + Math.max(0, accruedInterest)), 
    count: count, 
    lastDate: lastDateBS, 
    rate: lastRate, 
    lastK: Math.round(lastK) 
  };
}

// ३. फोटो सेभ गर्ने (सुधारिएको Drive access)
function saveFile(obj) {
  try {
    let folder, folders = DriveApp.getFoldersByName("Bus_Management_Photos");
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Bus_Management_Photos");
    
    const contentType = obj.imageBlob.split(',')[0].split(':')[1].split(';')[0];
    const bytes = Utilities.base64Decode(obj.imageBlob.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, obj.name + "_" + obj.nepDate + ".jpg");
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // फोटो हेर्न मिल्ने बनाउन
    return '=HYPERLINK("' + file.getUrl() + '", "फोटो हेर्नुस्")';
  } catch(e) { return "Error: " + e.toString(); }
}

// ४. सेटिङ र अन्य युटिलिटी
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sSheet = ss.getSheetByName("सेटिङ") || ss.insertSheet("सेटिङ");
  if(sSheet.getLastRow() < 1) sSheet.appendRow(["Setting Key", "Setting Value"]);
  const data = sSheet.getDataRange().getValues();
  let settings = {};
  for(let i=1; i<data.length; i++) { if(data[i][0]) settings[data[i][0]] = data[i][1]; }
  return settings;
}

function updateSettings(k, v) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sSheet = ss.getSheetByName("सेटिङ") || ss.insertSheet("सेटिङ");
  const data = sSheet.getDataRange().getValues();
  let found = false;
  for(let i=1; i<data.length; i++) { 
    if(data[i][0] == k) { 
      sSheet.getRange(i+1, 2).setValue(v); 
      found = true; 
      break; 
    } 
  }
  if(!found) sSheet.appendRow([k, v]);
  return "SUCCESS";
}

function formatMySheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return;
  
  sheet.getRange(1, 1, lastRow, lastCol)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Mukta")
    .setFontSize(10);
    
  sheet.autoResizeColumns(1, lastCol);
  
  // कोलम चौडाइमा २५ पिक्सेल प्याडिङ थप्ने
  for(let i=1; i<=lastCol; i++) {
    sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 25);
  }
}
