const FOLDER_ID = "1s4i2xXEbzsawTmsVPWHRSpgVkiNVB3KO"; 
const SPREADSHEET_ID = "1Q8YgtLChQ2dLTem-Tz2_NelDrCiJwDeJYzRMEBSCrSI";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- सेटिङ व्यवस्थापन ---
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

// --- गणना र नम्बर परिवर्तन ---
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
      let data = sheet.getDataRange().getValues();
      for (let j = data.length - 1; j >= 1; j--) {
        if (toEngNum(data[j][5]).toString().trim() === searchBus) {
          let lastKM = parseFloat(toEngNum(data[j][10]));
          if (!isNaN(lastKM) && lastKM > 0) return lastKM;
        }
      }
    }
  }
  return 0;
}

// --- मुख्य डेटा प्रशोधन ---
function process(data, photoObj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = "२०८३ " + data.nepMonthName;
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    const headers = ["मिति (BS)", "बार", "मिति (AD)", "प्रकार", "संस्था/रुट", "बस नं", "ड्राइभर", "लिटर", "रेट", "डिजल रकम", "आजको KM", "चलेको KM", "रिजर्भ रकम", "बैना/खर्च", "बचत", "कुल डिजेल लिटर", "कुल डिजेल रकम", "कुल रिजर्भ बचत", "विवरण", "फोटो", "KEY"];
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#22c55e").setFontColor("white").setHorizontalAlignment("center").setVerticalAlignment("middle");
      sheet.setFrozenRows(1);
    }

    const lastRow = sheet.getLastRow();

    // मानहरू तान्ने
    const curLit = parseFloat(toEngNum(data.dLiter)) || 0;
    const curAmt = parseFloat(toEngNum(data.dAmount)) || 0;
    
    // सुधार २: रिजर्भ सेक्सन नाफा हिसाब = भाडा - खर्च - डिजेल रकम
    const resAmt = parseFloat(toEngNum(data.totalReserveAmount)) || 0;
    const resExp = parseFloat(toEngNum(data.staffAllowance)) || 0;
    let curBal = 0;
    if (data.entryType === "Reserve") {
      curBal = resAmt - resExp - curAmt; // भाडा - खर्च - डिजेल रकम
    } else {
      curBal = parseFloat(toEngNum(data.balance)) || 0;
    }

    const busNumStr = data.busNumber.toString().trim();
    const instOrRoute = (data.entryType === "Institution" ? data.instName : data.routeFrom + " - " + data.routeTo);

    // सुधार ३: बस र संस्था मिल्दा मात्र विधागत ब्याकअप (रनिङ टोटल) हिसाब गर्ने
    let prevTotalLiter = 0;
    let prevTotalDAmount = 0;
    let prevTotalResBal = 0;

    if (lastRow > 1) {
      const fullData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      for (let i = 0; i < fullData.length; i++) {
        if (fullData[i][5].toString().trim() === busNumStr && fullData[i][4].toString().trim() === instOrRoute) {
          prevTotalLiter += parseFloat(fullData[i][7]) || 0;
          prevTotalDAmount += parseFloat(fullData[i][9]) || 0;
          prevTotalResBal += parseFloat(fullData[i][14]) || 0;
        }
      }
    }

    const newTotalLiter = prevTotalLiter + curLit;
    const newTotalDAmount = prevTotalDAmount + curAmt;
    const newTotalResBal = prevTotalResBal + curBal;
    
    let photoLink = "फोटो छैन";
    if (photoObj && photoObj.base64) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const blob = Utilities.newBlob(Utilities.base64Decode(photoObj.base64), photoObj.mimeType, photoObj.fileName);
      photoLink = '=HYPERLINK("' + folder.createFile(blob).getUrl() + '", "फोटो हेर्नुहोस्")';
    }

    const lastKMVal = parseFloat(toEngNum(data.lastKM)) || 0;
    const todayKMInput = parseFloat(toEngNum(data.currentKM)) || 0;
    let drivenKM = (todayKMInput > lastKMVal) ? (todayKMInput - lastKMVal) : 0;

    const rowData = [
      toNepNum(data.nepDateRaw), data.nepDay, data.engDate, 
      (data.entryType === "Institution" ? "संस्था" : "रिजर्भ"),
      instOrRoute, busNumStr, data.driverName,
      curLit, data.dRate || 0, curAmt,
      todayKMInput, drivenKM,
      resAmt, resExp, curBal,
      newTotalLiter, newTotalDAmount, newTotalResBal, 
      data.remarks || "", photoLink, (data.entryType === "Institution" ? (data.nepDateRaw + "|" + instOrRoute + "|" + busNumStr) : "RES_" + Utilities.getUuid())
    ];

    sheet.appendRow(rowData);
    const nLastRow = sheet.getLastRow();
    const range = sheet.getRange(nLastRow, 1, 1, headers.length);
    range.setHorizontalAlignment("center").setVerticalAlignment("middle");

    if (data.shiftColor) {
      sheet.getRange(nLastRow, 5).setFontColor(data.shiftColor).setFontWeight("bold");
    }

    if (data.entryType === "Reserve") {
      range.setFontColor("#ff0000").setFontWeight("bold");
    }

    if (nLastRow > 1) {
      sheet.getRange(2, 1, nLastRow - 1, headers.length).sort({column: 1, ascending: true});
    }

    sheet.autoResizeColumns(1, headers.length);
    for (let col = 1; col <= headers.length; col++) {
      sheet.setColumnWidth(col, sheet.getColumnWidth(col) + 35); 
    }

    // सुधार ४: बस, ड्राइभर, संस्था, ट्रिप र अन्तिम गतेको ब्याकअप 'Last_Settings' मा राख्ने
    saveLastActiveSettings(data);

    return "SUCCESS";
  } catch (e) { return "Error: " + e.toString(); }
  finally { lock.releaseLock(); }
}

// सुधार ४: अन्तिम गते र सेटिङहरू ब्याकअप राख्ने फंक्शन
function saveLastActiveSettings(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sSheet = ss.getSheetByName("Last_Settings") || ss.insertSheet("Last_Settings");
    sSheet.clearContents();
    sSheet.appendRow(["Parameter", "Value"]);
    sSheet.appendRow(["Last_Bus", data.busNumber || ""]);
    sSheet.appendRow(["Last_Driver", data.driverName || ""]);
    sSheet.appendRow(["Last_Institution", data.instName || ""]);
    sSheet.appendRow(["Last_Trip", data.trip || ""]);
    sSheet.appendRow(["Manual_Offset", data.manualOffsetSaved || "0"]);
  } catch(e){}
}
