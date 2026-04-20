const FOLDER_ID = "1pgnhX7iHuxAMiWviDe5m0Q2B0VoxQ8oe"; 

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// अङ्कलाई नेपालीमा बदल्ने फङ्सन
function toNepNum(n) {
  const nepDigits = ['०','१','२','३','४','५','६','७','८','९'];
  return n.toString().replace(/\d/g, d => nepDigits[d]);
}

function getLastKM(busNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let maxKM = 0;
  for (let i = 0; i < sheets.length; i++) {
    const data = sheets[i].getDataRange().getValues();
    for (let j = data.length - 1; j >= 1; j--) {
      if (data[j][5] && data[j][5].toString().trim() === busNumber.toString().trim()) { 
        let val = parseFloat(data[j][10]);
        if(!isNaN(val)) maxKM = Math.max(maxKM, val);
      }
    }
  }
  return maxKM;
}

function process(data, photoObj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // सिटको नाम: २०८३ वैशाख
    const sheetName = "२०८३ " + data.nepMonthName;

    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    if (sheet.getLastRow() === 0) {
      const headers = ["मिति (BS)", "बार", "मिति (AD)", "प्रकार", "संस्था/रुट", "बस नं", "ड्राइभर", "लिटर", "रेट", "डिजल रकम", "आजको KM", "चलेको KM", "रिजर्भ रकम", "बैना/खर्च", "बचत", "विवरण", "बसको कुल डिजल", "बसको कुल बचत", "फोटो"];
      sheet.appendRow(headers).getRange(1, 1, 1, 19).setFontWeight("bold").setBackground("#22c55e").setFontColor("white").setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    let photoLink = "फोटो छैन";
    if (photoObj && photoObj.base64) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const blob = Utilities.newBlob(Utilities.base64Decode(photoObj.base64), photoObj.mimeType, photoObj.fileName);
      photoLink = '=HYPERLINK("' + folder.createFile(blob).getUrl() + '", "फोटो हेर्नुहोस्")';
    }

    const nepType = data.entryType === "Institution" ? "संस्था" : "रिजर्भ";
    const instName = data.entryType === "Institution" ? data.instName : (data.routeFrom + " - " + data.routeTo);

    // कुल डिजल र बचत गणना
    let totalBusDiesel = (parseFloat(data.dAmount) || 0);
    let totalBusBalance = (parseFloat(data.balance) || 0);
    const allSheets = ss.getSheets();
    allSheets.forEach(s => {
      const sData = s.getDataRange().getValues();
      for(let r = 1; r < sData.length; r++) {
        if(sData[r][5] && sData[r][5].toString().trim() === data.busNumber.toString().trim()) {
          totalBusDiesel += (parseFloat(sData[r][9]) || 0);
          totalBusBalance += (parseFloat(sData[r][14]) || 0);
        }
      }
    });

    const rowData = [
      toNepNum(data.nepDateRaw), 
      data.nepDay, 
      data.engDate, 
      nepType, 
      instName, 
      data.busNumber, 
      data.driverName,
      data.dLiter, 
      data.dRate, 
      data.dAmount, 
      data.currentKM, 
      (parseFloat(data.currentKM) - parseFloat(data.lastKM || 0)),
      data.totalReserveAmount || 0, 
      data.staffAllowance || 0, 
      data.balance || 0,
      data.remarks || "", 
      Math.round(totalBusDiesel), 
      Math.round(totalBusBalance), 
      photoLink
    ];

    sheet.appendRow(rowData);
    
    // मिति अनुसार मिलाउने (Sorting)
    const lastRow = sheet.getLastRow();
    if(lastRow > 2) {
      sheet.getRange(2, 1, lastRow - 1, 19).sort({column: 1, ascending: true});
    }

    // १. सबै कोलमलाई अटो-रिसाइज गर्ने
    sheet.autoResizeColumns(1, 19);

    // २. कोलमको चौडाइमा थप २० पिक्सेल खाली ठाउँ (Padding) थप्ने ताकि अक्षर स्पष्ट देखियोस्
    for (let col = 1; col <= 19; col++) {
      let currentWidth = sheet.getColumnWidth(col);
      sheet.setColumnWidth(col, currentWidth + 25); // यहाँ २५ थपिएको छ
    }

    sheet.getRange(lastRow, 1, 1, 19).setVerticalAlignment("middle").setHorizontalAlignment("center");

    return "SUCCESS";
  } catch (e) { return "Error: " + e.toString(); }
}
