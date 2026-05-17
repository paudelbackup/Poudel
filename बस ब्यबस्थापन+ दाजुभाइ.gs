/**
 * बस व्यवस्थापन प्रणाली - सुपर फास्ट र अटोमेटिक (Full Version)
 * अपडेट: मिति अनुसार अटो-सर्टिङ सुविधा सहित
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('बस व्यवस्थापन प्रणाली')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// १. डाटा सुरक्षित गर्ने र मिति अनुसार अटो-सर्ट गर्ने फंक्सन
function processEntry(obj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = obj.sheetYearMonth.trim(); 
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    // नयाँ महिनाको सिट हो भने हेडर राख्ने
    if (sheet.getLastRow() === 0) {
      const headers = ["मिति (BS)", "मिति (AD)", "प्रकार", "नाम", "ब्याज दर %", "थप सावाँ/बिल", "किस्ता/भुक्तानी", "तिरेको ब्याज", "कुल बाँकी", "कैफियत", "फोटो"];
      sheet.appendRow(headers);
      
      // हेडर रेन्ज लिने
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      
      // १. हेडरलाई बोल्ड, ब्याकग्राउन्ड सेट गर्ने र सेन्टर एलाइन गर्ने
      headerRange.setFontWeight("bold").setBackground("#f3f3f3").setHorizontalAlignment("center");
      
      // २. हेडरमा चारैतिर पातलो कालो बोर्डर राख्ने (परिवर्तन गरिएको)
      headerRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    }

    const photoStatus = (obj.imageBlob && obj.imageBlob.includes(',')) ? saveFile(obj) : "फोटो छैन";

    // डाटा इन्ट्री गर्ने
    sheet.appendRow([
      obj.nepDate, 
      obj.engDate, 
      obj.cat.trim(), 
      obj.name.trim(), 
      parseFloat(obj.rate) || 0, 
      parseFloat(obj.addSawa_or_Bill) || 0, 
      parseFloat(obj.paidKista_or_Pay) || 0, 
      parseFloat(obj.paidByaj) || 0, 
      parseFloat(obj.totalAmt) || 0, 
      obj.remarks, 
      photoStatus
    ]);

    // --- मिति अनुसार डाटा मिलाउने लजिक (Sorting) ---
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) { 
      const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
      // कोलम २ (English Date - AD) को आधारमा सानो देखि ठूलो क्रममा मिलाउने
      range.sort({column: 2, ascending: true});
    }
    // ----------------------------------------------

    formatMySheet(sheet);
    return "SUCCESS";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

// २. समरी र ब्याज गणना (सुपर फास्ट फिल्टरिङ लोजिक - जस्ताको तस्तै)
function getSummary(name, cat, selectedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let sawa = 0, accruedInterest = 0, count = 0;
  let lastDateAD = null, lastDateBS = "-", lastRate = 0, lastK = 0;
  
  const selDate = new Date(selectedDate);
  const targetName = name.trim();
  const targetCat = cat.trim();
  
  let allEntries = [];
  const nepMonths = ["वैशाख", "जेठ", "असार", "साउन", "भदौ", "असोज", "कात्तिक", "मंसिर", "पुष", "माघ", "फागुन", "चैत", "बैशाख"];

  sheets.forEach(sheet => {
    const sName = sheet.getName();
    let isMonthSheet = nepMonths.some(m => sName.includes(m));
    if (!isMonthSheet) return;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][3] && 
          data[i][2].toString().trim() === targetCat && 
          data[i][3].toString().trim() === targetName) {
        allEntries.push(data[i]);
      }
    }
  });

  allEntries.sort((a, b) => new Date(a[1]) - new Date(b[1]));

  allEntries.forEach(row => {
    let rowDate = new Date(row[1]);
    if (rowDate > selDate) return; 

    let currentRate = parseFloat(row[4]) || 0;
    if (lastDateAD && (targetCat === "ऋण" || targetCat === "व्यक्तिगत")) {
      let days = Math.floor((rowDate.getTime() - lastDateAD.getTime()) / (1000 * 60 * 60 * 24));
      if (days > 0) accruedInterest += (sawa * (lastRate / 100) * days) / 365;
    }
    
    let addAmt = parseFloat(row[5]) || 0; 
    let subAmt = parseFloat(row[6]) || 0; 
    let interestPaid = parseFloat(row[7]) || 0; 

    if (targetCat === "ऋण" || targetCat === "व्यक्तिगत") {
      sawa += (addAmt - subAmt);
      accruedInterest -= interestPaid; 
    } else { 
      sawa += (addAmt - subAmt); 
    }
    
    lastDateAD = rowDate; 
    lastDateBS = row[0]; 
    lastRate = currentRate;
    lastK = subAmt > 0 ? subAmt : (addAmt > 0 ? addAmt : 0);
    count++;
  });

  if (lastDateAD && (targetCat === "ऋण" || targetCat === "व्यक्तिगत")) {
    let finalDays = Math.floor((selDate.getTime() - lastDateAD.getTime()) / (1000 * 60 * 60 * 24));
    if (finalDays > 0) accruedInterest += (sawa * (lastRate / 100) * finalDays) / 365;
  }

  return { 
    sawa: Math.round(sawa), 
    accruedInterest: Math.max(0, Math.round(accruedInterest)), 
    total: Math.round(sawa + (targetCat === "ऋण" || targetCat === "व्यक्तिगत" ? Math.max(0, accruedInterest) : 0)), 
    count: count, 
    lastDate: lastDateBS, 
    rate: lastRate, 
    lastK: Math.round(lastK) 
  };
}

// ३. फोटो र सेटिङ (Standard)
function saveFile(obj) {
  try {
    let folder, folders = DriveApp.getFoldersByName("Bus_Management_Photos");
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Bus_Management_Photos");
    const bytes = Utilities.base64Decode(obj.imageBlob.split(',')[1]);
    const blob = Utilities.newBlob(bytes, "image/jpeg", obj.name + "_" + obj.nepDate + ".jpg");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return '=HYPERLINK("' + file.getUrl() + '", "फोटो हेर्नुस्")';
  } catch(e) { return "Error: " + e.toString(); }
}

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sSheet = ss.getSheetByName("सेटिङ") || ss.insertSheet("सेटिङ");
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
    if(data[i][0] == k) { sSheet.getRange(i+1, 2).setValue(v); found = true; break; } 
  }
  if(!found) sSheet.appendRow([k, v]);
  return "SUCCESS";
}

function formatMySheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  sheet.getRange(1, 1, lastRow, lastCol).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontFamily("Mukta");
  sheet.autoResizeColumns(1, lastCol);
  
  // ३. पहिलो कोलम (मिति) लाई अझ फराकिलो (Wide) बनाउने ताकि डाटाहरू तेर्सो धर्कोमै अटाइरहून्
  sheet.setColumnWidth(1, 160); 
}
