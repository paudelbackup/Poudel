function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('बस व्यवस्थापन प्रणाली')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// १. डाटा सेभ र सिट फर्म्याटिङ (Bold, Center, Padding)
function processEntry(obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Transactions");
  if (!sheet) {
    sheet = ss.insertSheet("Transactions");
    sheet.appendRow(["मिति (BS)", "मिति (AD)", "प्रकार", "नाम", "ब्याज दर %", "थप सावाँ/बिल", "किस्ता/भुक्तानी", "तिरेको ब्याज", "बाँकी मौज्दात", "कैफियत", "फोटो लिङ्क"]);
  }
  sheet.getRange(1, 1, 1, 11).setFontWeight("bold").setBackground("#f3f3f3").setHorizontalAlignment("center");

  sheet.appendRow([obj.nepDate, obj.engDate, obj.cat, obj.name, obj.rate, obj.addSawa_or_Bill, obj.paidKista_or_Pay, obj.paidByaj, obj.totalAmt, obj.remarks, obj.imageBlob ? saveFile(obj) : ""]);

  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(lastRow, 1, 1, 11);
  range.setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(false);
  sheet.autoResizeColumns(1, 11);
  for(let i=1; i<=11; i++){ sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 30); }
  return "SUCCESS";
}

// २. समरी र ब्याज क्याल्कुलेसन
function getSummary(name, cat, selectedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");
  if (!sheet) return { sawa: 0, accruedInterest: 0, total: 0, count: 0, lastDate: "-", rate: 0, lastK: 0 };

  const data = sheet.getDataRange().getValues();
  let sawa = 0, accruedInterest = 0, count = 0, lastDateAD = "-", lastDateBS = "-", lastRate = 0, lastK = 0;
  const selDate = new Date(selectedDate);

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

function getSettings() { return PropertiesService.getScriptProperties().getProperties(); }
function updateSettings(k, v) { PropertiesService.getScriptProperties().setProperty(k, v); return "SUCCESS"; }
