const FOLDER_ID = "1pgnhX7iHuxAMiWviDe5m0Q2B0VoxQ8oe"; 

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function toNepNum(n) {
  const nepDigits = ['०','१','२','३','४','५','६','७','८','९'];
  return n.toString().replace(/\d/g, d => nepDigits[d]);
}

// बसको अन्तिम KM निकाल्ने फङ्सन (सस्था वा रिजर्भ जुनसुकै भए पनि बस नं हेर्छ)
function getLastKM(busNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let latestKM = 0;
  for (let i = 0; i < sheets.length; i++) {
    const data = sheets[i].getDataRange().getValues();
    // तलबाट माथि खोज्दै जाने (पछिल्लो रेकर्ड पहिला भेटिन्छ)
    for (let j = data.length - 1; j >= 1; j--) {
      if (data[j][5] && data[j][5].toString().trim() === busNumber.toString().trim()) { 
        let val = parseFloat(data[j][10]); // कोलम K (आजको KM)
        if(!isNaN(val)) return val;
      }
    }
  }
  return 0;
}

function process(data, photoObj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "२०८३ " + data.nepMonthName;
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    if (sheet.getLastRow() === 0) {
      const headers = [
        "मिति (BS)", "बार", "मिति (AD)", "प्रकार", "संस्था/रुट", "बस नं", "ड्राइभर", 
        "लिटर", "रेट", "डिजल रकम", "आजको KM", "चलेको KM", "रिजर्भ रकम", "बैना/खर्च", "बचत", 
        "कुल डिजल लिटर", "कुल डिजल रकम", "कुल रिजर्भ बचत", "विवरण", "फोटो"
      ];
      sheet.appendRow(headers).getRange(1, 1, 1, 20).setFontWeight("bold").setBackground("#22c55e").setFontColor("white").setHorizontalAlignment("center");
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

    // --- गणना लोजिक ---
    let runningDieselLiter = (parseFloat(data.dLiter) || 0);
    let runningDieselAmount = (parseFloat(data.dAmount) || 0);
    let runningReserveBalance = (data.entryType === "Reserve") ? (parseFloat(data.balance) || 0) : 0;
    
    const currentSheetData = sheet.getDataRange().getValues();

    if (currentSheetData.length > 1) {
      for(let r = 1; r < currentSheetData.length; r++) {
        const rowType = currentSheetData[r][3] ? currentSheetData[r][3].toString().trim() : ""; 
        const rowInstName = currentSheetData[r][4] ? currentSheetData[r][4].toString().trim() : "";
        const rowBusNumber = currentSheetData[r][5] ? currentSheetData[r][5].toString().trim() : "";
        const rowLiter = parseFloat(currentSheetData[r][7]) || 0;
        const rowAmount = parseFloat(currentSheetData[r][9]) || 0;
        const rowBal = parseFloat(currentSheetData[r][14]) || 0;

        // १. डिजलको हिसाब (Liter र Amount)
        if (data.entryType === "Institution") {
          // संस्थाको हकमा नाम मिलेमा मात्र जोड्ने
          if (rowType === "संस्था" && rowInstName === data.instName) {
            runningDieselLiter += rowLiter;
            runningDieselAmount += rowAmount;
          }
        } else {
          // रिजर्भको हकमा सबै गाडीको रिजर्भ इन्ट्री जोड्ने
          if (rowType === "रिजर्भ") {
            runningDieselLiter += rowLiter;
            runningDieselAmount += rowAmount;
          }
        }

        // २. कुल रिजर्भ बचत (सबै बसको रिजर्भ मात्र जोड्ने, संस्थाको नजोड्ने)
        if (rowType === "रिजर्भ") {
          runningReserveBalance += rowBal;
        }
      }
    }

    // बसको अन्तिम KM र चलेको KM हिसाब
    const lastKMValue = getLastKM(data.busNumber);
    const todayKM = parseFloat(data.currentKM) || 0;
    const drivenKM = lastKMValue > 0 ? (todayKM - lastKMValue) : 0;

    // संस्थामा कुल बचत खाली राख्ने (R कोलम)
    let finalRunningBalance = (data.entryType === "Reserve") ? Math.round(runningReserveBalance) : "";

    const rowData = [
      toNepNum(data.nepDateRaw), // A
      data.nepDay,               // B
      data.engDate,              // C
      nepType,                   // D
      instName,                  // E
      data.busNumber,            // F
      data.driverName,           // G
      data.dLiter,               // H
      data.dRate,                // I
      data.dAmount,              // J
      data.currentKM,            // K
      drivenKM,                  // L
      data.totalReserveAmount || 0, // M
      data.staffAllowance || 0,     // N
      data.balance || 0,            // O
      runningDieselLiter.toFixed(2),// P
      Math.round(runningDieselAmount),// Q
      finalRunningBalance,          // R: रिजर्भ भए मात्र आउँछ, संस्थामा खाली रहन्छ
      data.remarks || "",           // S
      photoLink                     // T
    ];

    sheet.appendRow(rowData);
    
    const lastRow = sheet.getLastRow();
    if(lastRow > 2) {
      sheet.getRange(2, 1, lastRow - 1, 20).sort({column: 1, ascending: true});
    }

    sheet.autoResizeColumns(1, 20);
    for (let col = 1; col <= 20; col++) {
      let currentWidth = sheet.getColumnWidth(col);
      sheet.setColumnWidth(col, currentWidth + 25);
    }
    sheet.getRange(lastRow, 1, 1, 20).setVerticalAlignment("middle").setHorizontalAlignment("center");

    return "SUCCESS";
  } catch (e) { return "Error: " + e.toString(); }
}
