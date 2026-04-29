/**
 * बस व्यवस्थापन प्रणाली - पूर्ण कोड (Code.gs)
 * यो कोडले डेटा इन्ट्री, महिना अनुसार नयाँ सिट निर्माण, 
 * र सिट फर्म्याटिङ (Row Height & Center alignment) व्यवस्थित गर्दछ।
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('बस व्यवस्थापन प्रणाली')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// १. डेटा इन्ट्री र सिट व्यवस्थापन गर्ने मुख्य फङ्सन
function processEntry(obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = obj.sheetYearMonth; 
  let sheet = ss.getSheetByName(sheetName);
  
  // यदि यो महिनाको नयाँ सिट छैन भने बनाउने
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = ["मिति (BS)", "मिति (AD)", "प्रकार", "नाम", "ब्याज दर %", "थप सावाँ/बिल", "किस्ता/भुक्तानी", "तिरेको ब्याज", "बाँकी मौज्दात", "कैफियत", "फोटो"];
    sheet.appendRow(headers);
  }

  // फोटो लिङ्क तयार गर्ने
  const photoStatus = obj.imageBlob ? saveFile(obj) : "फोटो छैन";

  // सिटमा नयाँ डेटा थप्ने
  sheet.appendRow([
    obj.nepDate, obj.engDate, obj.cat, obj.name, obj.rate, 
    obj.addSawa_or_Bill, obj.paidKista_or_Pay, obj.paidByaj, 
    obj.totalAmt, obj.remarks, photoStatus
  ]);

  // मिति (BS) अनुसार क्रमबद्ध गर्ने
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort({column: 1, ascending: true});
  }

  // सिटको डिजाइन र उचाइ मिलाउने
  formatMySheet(sheet);
  
  return "SUCCESS";
}

// २. सिटलाई चिटिक्क बनाउने र Row Height २१ मा फिक्स गर्ने फङ्सन
function formatMySheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0) return;

  const range = sheet.getRange(1, 1, lastRow, lastCol);

  // सबै सेललाई बीचमा (Center) राख्ने र Row नतन्किने बनाउने
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); 
  range.setHorizontalAlignment("center");
  range.setVerticalAlignment("middle");
  range.setFontFamily("Mukta");
  range.setFontSize(10);

  // कोलमको चौडाइ डेटा अनुसार अटोमेटिक मिलाउने
  sheet.autoResizeColumns(1, lastCol);
  
  // थप स्पष्टताको लागि कोलममा २५ पिक्सेल प्याडिङ थप्ने
  for (let i = 1; i <= lastCol; i++) {
    let currentWidth = sheet.getColumnWidth(i);
    sheet.setColumnWidth(i, currentWidth + 25); 
  }

  // हेडर डिजाइन (रंगीचंगी)
  const colors = ["#FFD1DC", "#C1E1C1", "#AEC6CF", "#FDFD96", "#FFB347", "#B39EB5", "#FF6961", "#77DD77", "#84B6F4", "#F49AC2", "#CB99C9"];

  for (let j = 1; j <= lastCol; j++) {
    let headerCell = sheet.getRange(1, j);
    headerCell.setBackground(colors[(j - 1) % colors.length]);
    headerCell.setFontWeight("bold");
    headerCell.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  }
  
  // डेटा एरियामा बोर्डर लगाउने
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  }

  // मुख्य समाधान: सबै Row को उचाइ २१ पिक्सेलमा फिक्स गर्ने (Standard Size)
  sheet.setRowHeights(1, lastRow, 21);
}

// ३. ड्यासबोर्डको लागि पुराना विवरण र ब्याज हिसाब गर्ने फङ्सन
function getSummary(name, cat, selectedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let sawa = 0, accruedInterest = 0, count = 0, lastDateAD = "-", lastDateBS = "-", lastRate = 0, lastK = 0;
  const selDate = new Date(selectedDate);

  sheets.forEach(sheet => {
    // सेटिङ सिटलाई हिसाबमा नमिसाउने
    if(sheet.getName() === "Settings" || sheet.getName() === "सेटिङ") return;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === cat && data[i][3] === name) {
        let rowDate = new Date(data[i][1]);
        if (rowDate > selDate) continue;
        
        let currentRate = parseFloat(data[i][4]) || 0;
        lastRate = currentRate;
        
        // ब्याज हिसाब (ऋण वा व्यक्तिगतको लागि मात्र)
        if (lastDateAD !== "-" && (cat === "ऋण" || cat === "व्यक्तिगत")) {
          let days = Math.floor((rowDate - new Date(lastDateAD)) / (1000 * 60 * 60 * 24));
          if (days > 0) accruedInterest += (sawa * (currentRate / 100) * days) / 365;
        }
        
        let tS = parseFloat(data[i][5]) || 0, pK = parseFloat(data[i][6]) || 0, pB = parseFloat(data[i][7]) || 0;
        
        // "अन्तिम" मा देखाउन पछिल्लो कारोबार रकम (lastK) स्टोर गर्ने
        if (cat === "ऋण" || cat === "व्यक्तिगत") {
          lastK = pK > 0 ? pK : (tS > 0 ? tS : 0);
          sawa += tS; accruedInterest -= pB; sawa -= (pK - pB);
        } else { 
          lastK = pK > 0 ? pK : (tS > 0 ? tS : 0);
          sawa += (tS - pK); 
        }
        
        lastDateAD = data[i][1]; 
        lastDateBS = data[i][0]; 
        count++;
      }
    }
  });

  // अन्तिम मिति देखि आज सम्मको बाँकी ब्याज जोड्ने
  if (lastDateAD !== "-" && (cat === "ऋण" || cat === "व्यक्तिगत")) {
    let finalDays = Math.floor((selDate - new Date(lastDateAD)) / (1000 * 60 * 60 * 24));
    if (finalDays > 0) accruedInterest += (sawa * (lastRate / 100) * finalDays) / 365;
  }

  return { 
    sawa: Math.round(sawa), 
    accruedInterest: Math.round(accruedInterest), 
    total: Math.round(sawa + accruedInterest), 
    count: count, 
    lastDate: lastDateBS, 
    rate: lastRate, 
    lastK: lastK 
  };
}

// ४. गुगल ड्राइभमा फोटो सेभ गर्ने फङ्सन
function saveFile(obj) {
  try {
    let folder, folders = DriveApp.getFoldersByName("Bus_Management_Photos");
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Bus_Management_Photos");
    const blob = Utilities.newBlob(Utilities.base64Decode(obj.imageBlob.split(',')[1]), "image/jpeg", obj.name + "_" + obj.nepDate + ".jpg");
    const file = folder.createFile(blob);
    return '=HYPERLINK("' + file.getUrl() + '", "फोटो हेर्नुस्")';
  } catch(e) { return "Error"; }
}

// ५. सेटिङ सिटबाट डेटा तान्ने फङ्सन
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sSheet = ss.getSheetByName("सेटिङ") || ss.insertSheet("सेटिङ");
  const data = sSheet.getDataRange().getValues();
  let settings = {};
  for(let i=1; i<data.length; i++) { settings[data[i][0]] = data[i][1]; }
  return settings;
}

// ६. सेटिङ अपडेट गर्ने र सिट फर्म्याट गर्ने फङ्सन
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
  
  if(!found) {
    sSheet.appendRow([k, v]);
  }
  
  sSheet.getRange(1, 1, 1, 2).setValues([["Key", "Value"]]);
  formatMySheet(sSheet); // सेटिङ सिटलाई पनि चिटिक्क बनाउने
  return "SUCCESS";
}
