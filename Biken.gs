/**
 * बस व्यवस्थापन प्रणाली - पूर्ण अपडेटेड Code.gs
 * सुधार गरिएको: बार (Day), अटो-सिट, फर्म्याटिङ र फोटो प्रिभ्यू सपोर्ट
 */

const FOLDER_ID = "1pgnhX7iHuxAMiWviDe5m0Q2B0VoxQ8oe"; 
const SPREADSHEET_ID = "1diMgxaMz8OS8Fm17W8QXz_FyTrXb61b-_cNcFxMr0xw";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- सेटिङ व्यवस्थापन (Dropdowns) ---
function getSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Add Setting");
  if (!sheet) {
    sheet = ss.insertSheet("Add Setting");
    sheet.appendRow(["busNumber", "driverName", "instName"]);
    sheet.getRange(1,1,1,3).setFontWeight("bold").setBackground("#f3f3f3");
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
  const sheet = ss.getSheetByName("Add Setting") || ss.insertSheet("Add Setting");
  const col = (key === 'busNumber') ? 1 : (key === 'driverName' ? 2 : 3);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, col).setValue(value);
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

// --- उपयोगिता (Utility) लजिक ---
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

// पछिल्लो किलोमिटर खोज्ने फङ्सन (सबै सिट चेक गर्छ)
function getLastKM(busNumber, currentMonthName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const months = ["बैशाख", "जेठ", "असार", "साउन", "भदौ", "असोज", "कात्तिक", "मंसिर", "पुष", "माघ", "फागुन", "चैत"];
  let currentIndex = months.indexOf(currentMonthName);
  let searchBus = toEngNum(busNumber).toString().trim();
  
  // हालको र अघिल्लो महिनाहरूमा बसको अन्तिम KM खोज्ने
  for (let i = currentIndex; i >= 0; i--) {
    let sheetName = "२०८३ " + months[i];
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      let lastRow = sheet.getLastRow();
      if (lastRow < 2) continue;
      let values = sheet.getRange(2, 1, lastRow - 1, 12).getValues(); // KM कोलम सम्मको डेटा
      for (let j = values.length - 1; j >= 0; j--) {
        if (toEngNum(values[j][5]).toString().trim() === searchBus) {
          let lastKM = parseFloat(toEngNum(values[j][10])); // Column K = आजको KM
          if (!isNaN(lastKM) && lastKM > 0) return lastKM;
        }
      }
    }
  }
  return 0;
}

// --- मुख्य डेटा प्रशोधन (Process) ---
function process(data, photoObj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = "२०८३ " + data.nepMonthName;
    let sheet = ss.getSheetByName(sheetName);
    
    // १. सिट छैन भने बनाउने र हेडर्स राख्ने
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = ["मिति (BS)", "बार", "मिति (AD)", "प्रकार", "संस्था/रुट", "बस नं", "ड्राइभर", "लिटर", "रेट", "डिजल रकम", "आजको KM", "चलेको KM", "रिजर्भ रकम", "बैना/खर्च", "बचत", "कुल डिजल लिटर", "कुल डिजल रकम", "कुल रिजर्भ बचत", "विवरण", "फोटो", "KEY"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight("bold")
           .setBackground("#4ade80") // Neon Green
           .setFontColor("black")
           .setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    const instOrRoute = (data.entryType === "Institution" ? data.instName : data.routeFrom + " - " + data.routeTo);
    const currentKey = (data.nepDateRaw + "|" + instOrRoute + "|" + data.busNumber).toString().trim();
    
    // २. डुप्लिकेट इन्ट्री चेक (संस्थाको लागि)
    if (data.entryType === "Institution") {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const keyData = sheet.getRange(2, 21, lastRow - 1, 1).getValues().flat();
        if (keyData.includes(currentKey)) {
          return "ERROR_DUPLICATE: यो बसको लागि आज इन्ट्री भइसकेको छ!";
        }
      }
    }

    // ३. फोटो अपलोड
    let photoLink = "फोटो छैन";
    if (photoObj && photoObj.base64) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const blob = Utilities.newBlob(Utilities.base64Decode(photoObj.base64), photoObj.mimeType, photoObj.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoLink = '=HYPERLINK("' + file.getUrl() + '", "फोटो हेर्नुहोस्")';
    }

    // ४. किलोमिटर हिसाब (नेपाली अंकलाई अंग्रेजीमा बदलेर)
    const lastKMVal = parseFloat(toEngNum(data.lastKM)) || 0;
    const todayKMInput = parseFloat(toEngNum(data.currentKM)) || 0;
    let drivenKM = (todayKMInput > lastKMVal) ? (todayKMInput - lastKMVal) : 0;

    // ५. रनिङ टोटल (Totaling Logic)
    let dLit = parseFloat(toEngNum(data.dLiter)) || 0;
    let dAmt = parseFloat(toEngNum(data.dAmount)) || 0;
    let resBal = parseFloat(toEngNum(data.balance)) || 0;

    // ६. डेटा पङ्क्ति तयार गर्ने
    const rowData = [
      toNepNum(data.nepDateRaw), 
      data.nepDay, // बार
      data.engDate, 
      (data.entryType === "Institution" ? "संस्था" : "रिजर्भ"),
      instOrRoute, 
      data.busNumber, 
      data.driverName,
      dLit, 
      parseFloat(toEngNum(data.dRate)) || 0, 
      dAmt,
      todayKMInput, 
      drivenKM,
      parseFloat(toEngNum(data.totalReserveAmount)) || 0, 
      parseFloat(toEngNum(data.staffAllowance)) || 0, 
      resBal,
      "", // कुल डिजल लिटर (पछि फर्मुला हाल्न सकिन्छ)
      "", // कुल डिजल रकम
      "", // कुल रिजर्भ बचत
      data.remarks || "", 
      photoLink, 
      (data.entryType === "Institution" ? currentKey : "RES_" + Utilities.getUuid())
    ];

    sheet.appendRow(rowData);
    
    // ७. फर्म्याटिङ र स्टाइल
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow, 1, 1, 21);
    range.setHorizontalAlignment("center").setVerticalAlignment("middle");
    
    // सिफ्टको कलर संस्था/रुट कोलममा लगाउने
    if (data.shiftColor) {
      sheet.getRange(lastRow, 5).setFontColor(data.shiftColor).setFontWeight("bold");
    }

    // रिजर्भ हो भने पूरै लाइन रातो/गुलाबी बनाउने
    if (data.entryType === "Reserve") {
      range.setBackground("#fff1f2").setFontColor("#e11d48").setFontWeight("bold");
    }

    sheet.autoResizeColumns(1, 21);
    return "SUCCESS";

  } catch (e) { 
    return "Error: " + e.toString(); 
  } finally { 
    lock.releaseLock(); 
  }
}
