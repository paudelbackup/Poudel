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
    sheet.appendRow(["मिति (BS)", "मिति (AD)", "प्रकार", "नाम", "ब्याज दर %", "थप सावाँ/बिल", "किस्ता/भुक्तानी", "तिरेको ब्याज", "बाँकी मौज्दात", "कैफियत", "फोटो"]);
    sheet.getRange(1, 1, 1, 11).setFontWeight("bold").setBackground("#f3f3f3").setHorizontalAlignment("center");
  }

  const photoStatus = obj.imageBlob ? saveFile(obj) : "फोटो छैन";

  sheet.appendRow([
    obj.nepDate, obj.engDate, obj.cat, obj.name, obj.rate, 
    obj.addSawa_or_Bill, obj.paidKista_or_Pay, obj.paidByaj, 
    obj.totalAmt, obj.remarks, photoStatus
  ]);

  sheet.autoResizeColumns(1, 11);
  return "SUCCESS";
}

function getSummary(name, cat, selectedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let sawa = 0, accruedInterest = 0, count = 0, lastDateAD = "-", lastDateBS = "-", lastRate = 0, lastK = 0;
  const selDate = new Date(selectedDate);

  sheets.forEach(sheet => {
    if(sheet.getName() === "Settings") return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === cat && data[i][3] === name) {
        let rowDate = new Date(data[i][1]);
        if (rowDate > selDate) continue;
        let currentRate = parseFloat(data[i][4]) || 0;
        lastRate = currentRate;
        
        if (lastDateAD !== "-" && (cat === "ऋण" || cat === "व्यक्तिगत")) {
          let days = Math.floor((rowDate - new Date(lastDateAD)) / (1000 * 60 * 60 * 24));
          if (days > 0) accruedInterest += (sawa * (currentRate / 100) * days) / 365;
        }
        
        let tS = parseFloat(data[i][5]) || 0, pK = parseFloat(data[i][6]) || 0, pB = parseFloat(data[i][7]) || 0;
        lastK = pK;
        if (cat === "व्यक्तिगत" || cat === "ऋण") {
          sawa += tS; accruedInterest -= pB; sawa -= (pK - pB);
        } else { sawa += (tS - pK); }
        lastDateAD = data[i][1]; lastDateBS = data[i][0]; count++;
      }
    }
  });

  if (lastDateAD !== "-" && (cat === "ऋण" || cat === "व्यक्तिगत")) {
    let finalDays = Math.floor((selDate - new Date(lastDateAD)) / (1000 * 60 * 60 * 24));
    if (finalDays > 0) accruedInterest += (sawa * (lastRate / 100) * finalDays) / 365;
  }

  return { sawa: Math.round(sawa), accruedInterest: Math.round(accruedInterest), total: Math.round(sawa + accruedInterest), count: count, lastDate: lastDateBS, rate: lastRate, lastK: lastK };
}

function saveFile(obj) {
  try {
    let folder, folders = DriveApp.getFoldersByName("Bus_Management_Photos");
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Bus_Management_Photos");
    const blob = Utilities.newBlob(Utilities.base64Decode(obj.imageBlob.split(',')[1]), "image/jpeg", obj.name + "_" + obj.nepDate + ".jpg");
    return folder.createFile(blob).getUrl();
  } catch(e) { return "Error"; }
}

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sSheet = ss.getSheetByName("Settings") || ss.insertSheet("Settings");
  const data = sSheet.getDataRange().getValues();
  let settings = {};
  for(let i=1; i<data.length; i++) { settings[data[i][0]] = data[i][1]; }
  return settings;
}

function updateSettings(k, v) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sSheet = ss.getSheetByName("Settings") || ss.insertSheet("Settings");
  const data = sSheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == k) { sSheet.getRange(i+1, 2).setValue(v); return "DONE"; }
  }
  sSheet.appendRow([k, v]);
  return "SUCCESS";
}
