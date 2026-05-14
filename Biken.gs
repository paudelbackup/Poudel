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
  const settings = { busNumber: [], driverName: [], instName: [], tripList: [], lastOffset: 0 };
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) settings.busNumber.push(data[i][0].toString());
    if (data[i][1]) settings.driverName.push(data[i][1].toString());
    if (data[i][2]) settings.instName.push(data[i][2].toString());
    if (data[i][3]) settings.tripList.push(data[i][3].toString());
  }
  
  let sSheet = ss.getSheetByName("Last_Settings");
  if (sSheet) {
    let sData = sSheet.getDataRange().getValues();
    sData.forEach(row => {
      if (row[0] === "Last_Bus") settings.lastBus = row[1];
      if (row[0] === "Last_Driver") settings.lastDriver = row[1];
      if (row[0] === "Last_Institution") settings.lastInstitution = row[1];
      if (row[0] === "Last_Trip") settings.lastTrip = row[1];
      if (row[0] === "Manual_Offset") settings.lastOffset = parseInt(row[1]) || 0;
    });
  }
  return settings;
}

function saveSettingToSheet(key, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Add Setting") || ss.insertSheet("Add Setting");
  const col = (key === 'busNumber') ? 1 : (key === 'driverName' ? 2 : (key === 'instName' ? 3 : 4));
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    const data = sheet.getRange(1, col, lastRow).getValues();
    for (let i = 0; i < data.length; i++) { if (data[i][0].toString().trim() === value.toString().trim()) return "EXISTS"; }
  }
  sheet.getRange(sheet.getLastRow() + 1, col).setValue(value);
  return "SAVED";
}

function removeSettingFromSheet(key, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Add Setting");
  if (!sheet) return "NOT_FOUND";
  const col = (key === 'busNumber') ? 1 : (key === 'driverName' ? 2 : (key === 'instName' ? 3 : 4));
  const data = sheet.getRange(1, col, sheet.getLastRow()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim() === value.toString().trim()) {
      sheet.getRange(i + 1, col).deleteCells(SpreadsheetApp.Dimension.ROWS);
      return "DELETED";
    }
  }
  return "NOT_FOUND";
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
      let data = sheet.getDataRange().getValues();
      for (let j = data.length - 1; j >= 1; j--) {
        // यहाँ तोEngNum ले नेपाली अङ्क भए पनि अङ्ग्रेजीमा बदलेर चेक गर्छ
        if (toEngNum(data[j][5]).toString().trim() === searchBus) {
          let lastKM = parseFloat(toEngNum(data[j][12])); 
          if (!isNaN(lastKM) && lastKM > 0) return lastKM;
        }
      }
    }
  }
  return 0;
}

// गते परिवर्तन गर्दा सिधै सेटिङमा सेभ गर्ने फङ्ग्सन
function saveDateOffset(offset) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sSheet = ss.getSheetByName("Last_Settings") || ss.insertSheet("Last_Settings");
    let data = sSheet.getDataRange().getValues();
    let found = false;
    for(let i=0; i<data.length; i++) {
      if(data[i][0] === "Manual_Offset") {
        sSheet.getRange(i+1, 2).setValue(offset);
        found = true; break;
      }
    }
    if(!found) sSheet.appendRow(["Manual_Offset", offset]);
    return "SAVED";
  } catch(e) { return e.toString(); }
}

function process(data, photoObj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = "२०८३ " + data.nepMonthName;
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    const headers = ["मिति (BS)", "बार", "मिति (AD)", "प्रकार", "संस्था/रुट", "बस नं", "ड्राइभर", "सिफ्ट", "ट्रिप", "लिटर", "रेट", "डिजल रकम", "आजको KM", "चलेको KM", "रिजर्भ रकम", "बैना/खर्च", "बचत", "कुल डिजेल लिटर", "कुल डिजेल रकम", "कुल रिजर्भ बचत", "विवरण", "फोटो", "KEY"];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#22c55e").setFontColor("white").setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    const curLit = parseFloat(toEngNum(data.dLiter)) || 0;
    const curAmt = parseFloat(toEngNum(data.dAmount)) || 0;
    const resAmt = parseFloat(toEngNum(data.totalReserveAmount)) || 0;
    const resExp = parseFloat(toEngNum(data.staffAllowance)) || 0;
    let curBal = (data.entryType === "Reserve") ? (resAmt - resExp - curAmt) : (parseFloat(toEngNum(data.balance)) || 0);

    const busNumStr = data.busNumber.toString().trim();
    const instOrRoute = (data.entryType === "Institution" ? data.instName : data.routeFrom + " - " + data.routeTo);
    const lastKMVal = parseFloat(toEngNum(data.lastKM)) || 0;
    const todayKMInput = parseFloat(toEngNum(data.currentKM)) || 0;
    let drivenKM = (todayKMInput > lastKMVal) ? (todayKMInput - lastKMVal) : 0;
    const nepaliDateInNepDigits = toNepNum(data.nepDateRaw);

    let photoLink = "फोटो छैन";
    if (photoObj && photoObj.base64) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const blob = Utilities.newBlob(Utilities.base64Decode(photoObj.base64), photoObj.mimeType, photoObj.fileName);
      photoLink = '=HYPERLINK("' + folder.createFile(blob).getUrl() + '", "फोटो हेर्नुहोस्")';
    }

    const rowData = [
      nepaliDateInNepDigits, data.nepDay, data.engDate, 
      (data.entryType === "Institution" ? "संस्था" : "रिजर्भ"),
      instOrRoute, busNumStr, data.driverName,
      data.shift || "", data.trip || "", 
      curLit, data.dRate || 0, curAmt,
      todayKMInput, drivenKM,
      resAmt, resExp, curBal,
      "", "", "", // कुल हिसाब पछि सिटमै फर्मुलाबाट मिलाउन सकिन्छ
      data.remarks || "", photoLink, (data.entryType === "Institution" ? (nepaliDateInNepDigits + "|" + instOrRoute + "|" + busNumStr) : "RES_" + Utilities.getUuid())
    ];

    sheet.appendRow(rowData);
    sheet.autoResizeColumns(1, headers.length);
    saveLastActiveSettings(data);
    return "SUCCESS";
  } catch (e) { return "Error: " + e.toString(); }
  finally { lock.releaseLock(); }
}

function saveLastActiveSettings(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sSheet = ss.getSheetByName("Last_Settings") || ss.insertSheet("Last_Settings");
    // यहाँ हामी पहिले नै भएका कुरा अपडेट गर्छौँ, clear गर्दैनौं ताकि Manual_Offset नहटोस्
    let existingData = sSheet.getDataRange().getValues();
    let updates = { "Last_Bus": data.busNumber, "Last_Driver": data.driverName, "Last_Institution": data.instName, "Last_Trip": data.trip, "Manual_Offset": data.manualOffsetSaved };
    
    for (let key in updates) {
      let found = false;
      for (let i=0; i<existingData.length; i++) {
        if (existingData[i][0] === key) {
          sSheet.getRange(i+1, 2).setValue(updates[key]);
          found = true; break;
        }
      }
      if (!found) sSheet.appendRow([key, updates[key]]);
    }
  } catch(e){}
}
