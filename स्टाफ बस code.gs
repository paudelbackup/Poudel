const FOLDER_ID = "1pgnhX7iHuxAMiWviDe5m0Q2B0VoxQ8oe"; 
const SPREADSHEET_ID = "1yd_Z3VdHTYELlSg-bwd6YOB2Tts_GZV-9hAWsZhPCDI";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- सेटिङ व्यवस्थापन ---
function getSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    sheet = ss.insertSheet("Settings");
    sheet.appendRow(["busNumber", "driverName", "instName"]);
  }
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
  const sheet = ss.getSheetByName("Settings") || ss.insertSheet("Settings");
  const col = (key === 'busNumber') ? 1 : (key === 'driverName' ? 2 : 3);
  sheet.getRange(sheet.getLastRow() + 1, col).setValue(value);
  return "SAVED";
}

function removeSettingFromSheet(key, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Settings");
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

// --- गणना लजिक ---
function toEngNum(n) {
  if (n === undefined || n === null || n === "") return "0";
  const nepDigits = {'०':'0','१':'1','२':'2','३':'3','४':'4','५':'5','६':'6','७':'7','८':'8','९':'9'};
  return n.toString().replace(/[०-९]/g, d => nepDigits[d]);
}

function toNepNum(n) {
  if (n === undefined || n === null) return "";
  const nepDigits = ['०','१','२','३','४','५','६','७','८','९'];
  return n.toString().replace(/\d/g, d => nepDigits[d]);
}

function getLastKM(busNumber, currentMonthName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const months = ["बैशाख", "जेठ", "असार", "साउन", "भदौ", "असोज", "कात्तिक", "मंसिर", "पुष", "माघ", "फागुन", "चैत"];
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

// --- डाटा प्रशोधन र अटो-साइज लजिक ---
function process(data, photoObj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = "२०८३ " + data.nepMonthName;
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    const headers = ["मिति (BS)", "बार", "मिति (AD)", "प्रकार", "संस्था/रुट", "बस नं", "ड्राइभर", "लिटर", "रेट", "डिजल रकम", "आजको KM", "चलेको KM", "रिजर्भ रकम", "बैना/खर्च", "बचत", "कुल डिजल लिटर", "कुल डिजल रकम", "कुल रिजर्भ बचत", "विवरण", "फोटो", "KEY"];
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#22c55e").setFontColor("white").setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // --- नयाँ UNIQUE KEY लजिक (मिति + संस्था + बस नम्बर) ---
    const instOrRoute = (data.entryType === "Institution" ? data.instName : data.routeFrom + " - " + data.routeTo);
    const currentKey = (data.nepDateRaw + "|" + instOrRoute + "|" + data.busNumber).toString().trim();
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const keyData = sheet.getRange(2, 21, lastRow - 1, 1).getValues().flat();
      if (keyData.includes(currentKey)) {
        return "ERROR_DUPLICATE: यो बसको लागि यो संस्थामा आजको इन्ट्री भइसकेको छ!";
      }
    }

    let photoLink = "फोटो छैन";
    if (photoObj && photoObj.base64) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const blob = Utilities.newBlob(Utilities.base64Decode(photoObj.base64), photoObj.mimeType, photoObj.fileName);
      photoLink = '=HYPERLINK("' + folder.createFile(blob).getUrl() + '", "फोटो हेर्नुहोस्")';
    }

    const lastKMVal = getLastKM(data.busNumber, data.nepMonthName);
    const todayKMInput = parseFloat(toEngNum(data.currentKM)) || 0;
    let drivenKM = (lastKMVal > 0 && todayKMInput > lastKMVal) ? (todayKMInput - lastKMVal) : 0;

    let runningDieselLiter = (parseFloat(toEngNum(data.dLiter)) || 0);
    let runningDieselAmount = (parseFloat(toEngNum(data.dAmount)) || 0);
    let runningReserveBalance = (data.entryType === "Reserve") ? (parseFloat(toEngNum(data.balance)) || 0) : 0;
    
    const existingData = sheet.getDataRange().getValues();
    for(let r = 1; r < existingData.length; r++) {
      if (data.entryType === "Institution" && existingData[r][3] === "संस्था" && existingData[r][4] === data.instName) {
          runningDieselLiter += parseFloat(toEngNum(existingData[r][7])) || 0;
          runningDieselAmount += parseFloat(toEngNum(existingData[r][9])) || 0;
      } else if (data.entryType === "Reserve" && existingData[r][3] === "रिजर्भ") {
          runningDieselLiter += parseFloat(toEngNum(existingData[r][7])) || 0;
          runningDieselAmount += parseFloat(toEngNum(existingData[r][9])) || 0;
          runningReserveBalance += parseFloat(toEngNum(existingData[r][14])) || 0;
      }
    }

    const rowData = [
      toNepNum(data.nepDateRaw), data.nepDay, data.engDate, 
      (data.entryType === "Institution" ? "संस्था" : "रिजर्भ"),
      instOrRoute,
      data.busNumber, data.driverName,
      data.dLiter || 0, data.dRate || 0, data.dAmount || 0,
      todayKMInput, drivenKM,
      data.totalReserveAmount || 0, data.staffAllowance || 0, data.balance || 0,
      runningDieselLiter.toFixed(2), Math.round(runningDieselAmount),
      (data.entryType === "Reserve" ? Math.round(runningReserveBalance) : ""),
      data.remarks || "", photoLink, currentKey
    ];

    sheet.appendRow(rowData);
    
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 21).sort({column: 1, ascending: true});
    }

    const finalRow = sheet.getLastRow();
    const range = sheet.getRange(2, 1, finalRow - 1, 21);
    const values = range.getValues();
    
    for (let i = 0; i < values.length; i++) {
      let rowNum = i + 2;
      let rRange = sheet.getRange(rowNum, 1, 1, 21);
      rRange.setVerticalAlignment("middle").setHorizontalAlignment("center");
      if (values[i][3] === "रिजर्भ") {
        rRange.setFontColor("#ff0000").setFontWeight("bold");
      } else {
        rRange.setFontColor("#000000").setFontWeight("normal");
      }
    }

    sheet.autoResizeColumns(1, 21);
    for (let col = 1; col <= 21; col++) {
      let currentWidth = sheet.getColumnWidth(col);
      sheet.setColumnWidth(col, currentWidth + 40); 
    }

    return "SUCCESS";
  } catch (e) { return "Error: " + e.toString(); }
  finally { lock.releaseLock(); }
}
