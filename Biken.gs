const FOLDER_ID = "1s4i2xXEbzsawTmsVPWHRSpgVkiNVB3KO"; 
const SPREADSHEET_ID = "1Q8YgtLChQ2dLTem-Tz2_NelDrCiJwDeJYzRMEBSCrSI";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Add Setting") || ss.insertSheet("Add Setting");
  const data = sheet.getDataRange().getValues();
  const settings = { busNumber: [], driverName: [], instName: [] };
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) settings.busNumber.push(data[i][0]);
    if (data[i][1]) settings.driverName.push(data[i][1]);
    if (data[i][2]) settings.instName.push(data[i][2]);
  }
  return settings;
}

function saveSettingToSheet(key, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Add Setting") || ss.insertSheet("Add Setting");
  const col = (key === 'busNumber') ? 1 : (key === 'driverName' ? 2 : 3);
  sheet.getRange(sheet.getLastRow() + 1, col).setValue(value);
  return "SAVED";
}

function removeSettingFromSheet(key, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Add Setting");
  if (!sheet) return;
  const col = (key === 'busNumber') ? 1 : (key === 'driverName' ? 2 : 3);
  const data = sheet.getRange(1, col, sheet.getLastRow()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString() === value.toString()) {
      sheet.getRange(i + 1, col).deleteCells(SpreadsheetApp.Dimension.ROWS);
      break;
    }
  }
}

function toEngNum(n) {
  if (!n) return "0";
  const nepDigits = {'०':'0','१':'1','२':'2','३':'3','४':'4','५':'5','६':'6','७':'7','८':'8','९':'9'};
  return n.toString().replace(/[०-९]/g, d => nepDigits[d]);
}

function toNepNum(n) {
  if (!n) return "";
  const nepDigits = ['०','१','२','३','४','५','६','७','८','९'];
  return n.toString().replace(/\d/g, d => nepDigits[d]);
}

function getLastKM(busNumber, currentMonthName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const months = ["वैशाख", "जेठ", "असार", "साउन", "भदौ", "असोज", "कात्तिक", "मंसिर", "पुष", "माघ", "फागुन", "चैत"];
  let currentIndex = months.indexOf(currentMonthName);
  let searchBus = toEngNum(busNumber).toString().trim();
  for (let i = currentIndex; i >= 0; i--) {
    let sheetName = "२०८३ " + months[i];
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      let lastRow = sheet.getLastRow();
      if (lastRow < 2) continue;
      let values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
      for (let j = values.length - 1; j >= 0; j--) {
        if (toEngNum(values[j][5]).toString().trim() === searchBus) {
          let lastKM = parseFloat(toEngNum(values[j][10]));
          if (!isNaN(lastKM) && lastKM > 0) return lastKM;
        }
      }
    }
  }
  return 0;
}

function process(data, photoObj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = "२०८३ " + data.nepMonthName;
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    const headers = ["मिति (BS)", "बार", "मिति (AD)", "प्रकार", "संस्था/रुट", "बस नं", "ड्राइभर", "लिटर", "रेट", "डिजल रकम", "आजको KM", "चलेको KM", "रिजर्भ रकम", "बैना/खर्च", "बचत", "कुल डिजल लिटर", "कुल डिजल रकम", "कुल बचत/ब्यालेन्स", "विवरण", "फोटो", "KEY"];
    
    // १. रङ्गीचङ्गी हेडर र स्टाइल
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      const colors = ["#E06666", "#F6B26B", "#FFD966", "#93C47D", "#76A5AF", "#6FA8DC", "#8E7CC3", "#C27BA0", "#A4C2F4", "#B4A7D6", "#D5A6BD", "#F4CCCC", "#FCE5CD", "#FFF2CC", "#D9EAD3", "#D0E0E3", "#CFE2F3", "#D9D2E9", "#EAD1DC", "#DD7E6B", "#CCCCCC"];
      for (let h = 0; h < headers.length; h++) {
        sheet.getRange(1, h + 1).setBackground(colors[h % colors.length])
             .setFontWeight("bold")
             .setFontColor("black")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle")
             .setBorder(true, true, true, true, true, true);
      }
      sheet.setRowHeight(1, 40);
      sheet.setFrozenRows(1);
    }

    const instOrRoute = (data.entryType === "Institution" ? data.instName : data.routeFrom + " - " + data.routeTo);
    const busNumStr = data.busNumber.toString().trim();
    
    let photoLink = "फोटो छैन";
    if (photoObj && photoObj.base64) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const blob = Utilities.newBlob(Utilities.base64Decode(photoObj.base64), photoObj.mimeType, photoObj.fileName);
      photoLink = '=HYPERLINK("' + folder.createFile(blob).getUrl() + '", "फोटो हेर्नुहोस्")';
    }

    const lastKMVal = getLastKM(data.busNumber, data.nepMonthName);
    const todayKMInput = parseFloat(toEngNum(data.currentKM)) || 0;
    let drivenKM = (lastKMVal > 0 && todayKMInput > lastKMVal) ? (todayKMInput - lastKMVal) : 0;

    let runningDieselLiter = parseFloat(toEngNum(data.dLiter)) || 0;
    let runningDieselAmount = parseFloat(toEngNum(data.dAmount)) || 0;
    let runningBalance = (data.entryType === "Reserve") ? (parseFloat(toEngNum(data.balance)) || 0) : 0;
    
    const existingData = sheet.getDataRange().getValues();
    for(let r = 1; r < existingData.length; r++) {
       if (data.entryType === "Institution") {
         if (existingData[r][3] === "संस्था" && existingData[r][4] === data.instName && existingData[r][5].toString().trim() === busNumStr) {
           runningDieselLiter += parseFloat(toEngNum(existingData[r][7])) || 0;
           runningDieselAmount += parseFloat(toEngNum(existingData[r][9])) || 0;
         }
       } else if (data.entryType === "Reserve" && existingData[r][3] === "रिजर्भ") {
         runningDieselLiter += parseFloat(toEngNum(existingData[r][7])) || 0;
         runningDieselAmount += parseFloat(toEngNum(existingData[r][9])) || 0;
         runningBalance += parseFloat(toEngNum(existingData[r][14])) || 0;
       }
    }

    const rowData = [
      toNepNum(data.nepDateRaw), data.nepDay, data.engDate, 
      (data.entryType === "Institution" ? "संस्था" : "रिजर्भ"),
      instOrRoute, busNumStr, data.driverName,
      data.dLiter || 0, data.dRate || 0, data.dAmount || 0,
      todayKMInput, drivenKM,
      data.totalReserveAmount || 0, data.staffAllowance || 0, data.balance || 0,
      runningDieselLiter.toFixed(2), Math.round(runningDieselAmount), Math.round(runningBalance),
      data.remarks || "", photoLink, (data.entryType === "Institution" ? (data.nepDateRaw + "|" + instOrRoute + "|" + busNumStr) : "RES_" + Utilities.getUuid())
    ];

    sheet.appendRow(rowData);
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow, 1, 1, headers.length);
    
    // २. बक्सको साइज नमिल्ने समस्याको समाधान (Vertical फिक्स गर्ने)
    range.setHorizontalAlignment("center")
         .setVerticalAlignment("middle")
         .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); // यसले रो को उचाई बढ्न दिँदैन

    sheet.setRowHeight(lastRow, 30); // हरेक रो को उचाई ३० मा फिक्स

    if (data.shiftColor) {
      sheet.getRange(lastRow, 5).setFontColor(data.shiftColor).setFontWeight("bold");
    }
    if (data.entryType === "Reserve") {
      range.setFontColor("#ff0000").setFontWeight("bold");
    }

    // ३. कोलम रिसाइज लजिक (KEY लाई ठूलो र अरूलाई अटो गर्ने)
    sheet.autoResizeColumns(1, 20); // १ देखि २० सम्म अटो गर्ने
    
    // KEY कोलम (२१) लाई पर्याप्त ठाउँ दिने ताकि नछोपियोस्
    sheet.setColumnWidth(21, 180); 

    // कोलमहरूमा थप खुल्ला ठाउँ दिने
    for(let c=1; c <= 20; c++) {
      let w = sheet.getColumnWidth(c);
      sheet.setColumnWidth(c, w + 25); 
    }

    return "SUCCESS";
  } catch (e) { return "Error: " + e.toString(); }
  finally { lock.releaseLock(); }
}
