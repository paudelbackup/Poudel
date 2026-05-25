Rohit Paudel:
const SHEET_ID = '1RpilD_cocEkq4yCpjTtIN07qGU2zmNDX4Lw4W1yFN6g'; // तपाईंको Sheet ID
const SHEET_NAME = 'पोल्ट्री रेकर्ड'; // Sheet को नाम, यदि फरक छ भने परिवर्तन गर्नुहोस्

function doGet() {
  // Faram खुल्नुअघि Sheet को ढाँचा (Headers, Resize, Colors) मिलाउने कार्य सुरु गर्ने
  setupSheetFormatting(); 
  
  return HtmlService.createHtmlOutputFromFile('Index') 
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// नयाँ Function: Sheet को हेडर र ढाँचा स्वचालित रूपमा मिलाउने
function setupSheetFormatting() {
  const ss = SpreadsheetApp.openById(SHEET_ID); 
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // यदि शीट छैन भने बनाउने
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // हेडरहरूको अन्तिम क्रम
  const headers = [
    'मिति', 'बार', 'सुरुवाती स्टक', 'मरेको संख्या', 'बेचेको संख्या', 'बाँकी कुखुरा', 
    'दाना मात्रा (Kg)', 'दानाको दर (रु.)', 'कुल खर्च', 'अन्य खर्च', 
    'मासु को दर', 'बेचेको (Kg)', 'कुल आम्दानी', 'कुल नाफा/नोक्सान', 
    'कामदार', 'औषधि', 'विवरण'
  ];
  const numColumns = headers.length;
  const HEADER_ROW = 1;
  
  // यदि Sheet खाली छ भने हेडर राख्ने (자동 헤더)
  if (sheet.getLastRow() < HEADER_ROW) {
    sheet.getRange(HEADER_ROW, 1, 1, numColumns).setValues([headers]);
  }
  
  // १. Header Formatting: Bold र Center
  const headerRange = sheet.getRange(HEADER_ROW, 1, 1, numColumns);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // २. Auto Resize Columns (डाटा अनुसार कोलम मिलाउने)
  sheet.autoResizeColumns(1, numColumns); 

  // ३. Conditional Formatting (नाफा/नोक्सानको लागि)
  // 'कुल नाफा/नोक्सान' 14 औं स्तम्भमा छ।
  const PROFIT_LOSS_COLUMN = 14; 
  const profitLossRange = sheet.getRange(2, PROFIT_LOSS_COLUMN, sheet.getMaxRows(), 1);

  // यदि नाफा (मूल्य ० भन्दा बढी) भएमा हरियो (Green)
  const ruleProfit = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setFontColor('#008000') // Green
      .setRanges([profitLossRange])
      .build();

  // यदि नोक्सान (मूल्य ० भन्दा कम) भएमा रातो (Red)
  const ruleLoss = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor('#FF0000') // Red
      .setRanges([profitLossRange])
      .build();

  // पुरानो Rule हटाएर नयाँ लागू गर्ने
  const rules = sheet.getConditionalFormatRules();
  // पुराना नाफा/नोक्सानका नियमहरू हटाउने
  const newRules = rules.filter(rule => rule.getRanges()[0].getColumn() !== PROFIT_LOSS_COLUMN);
  
  sheet.setConditionalFormatRules([...newRules, ruleProfit, ruleLoss]);

  Logger.log('Sheet formatting applied successfully.');
}


// अघिल्लो दिनको बाँकी कुखुराको संख्या तान्ने Function
function getPreviousDayStock() {
  const ss = SpreadsheetApp.openById(SHEET_ID); 
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() <= 1) return 0;

  // 'बाँकी कुखुरा' 6 औं स्तम्भमा छ
  const REMAINING_BIRDS_COLUMN = 6; 
  const lastRow = sheet.getLastRow();

  const lastRemainingStock = sheet.getRange(lastRow, REMAINING_BIRDS_COLUMN).getValue();
  
  return typeof lastRemainingStock === 'number' && lastRemainingStock >= 0 ? lastRemainingStock : 0;
}


// फारमको डाटा बचत गर्ने Function
function processPoultryForm(formData) {
  const ss = SpreadsheetApp.openById(SHEET_ID); 
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return { status: 'error', message: 'Google Sheet मा "' + SHEET_NAME + '" नामको पाना (Sheet) फेला परेन।' };
  }

  // नेपालीमा बार निकाल्ने
  const date = new Date(formData.date);
  const dayNamesNepali = ["आइतबार", "सोमबार", "मङ्गलबार", "बुधबार", "बिहीबार", "शुक्रबार", "शनिबार"];
  const dayOfWeek = dayNamesNepali[date.getDay()];
  
  // डाटाको क्रम
  const rowData = [
    formData.date,              // १. मिति

dayOfWeek,                  // २. बार
    formData.chicksCount,       // ३. सुरुवाती स्टक 
    formData.deadBirds,         // ४. मरेको संख्या
    formData.soldCount,         // ५. बेचेको संख्या
    formData.remainingBirds,    // ६. बाँकी कुखुरा (अन्तिम स्टक)
    formData.feedQty,           // ७. दाना मात्रा
    formData.feedRate,          // ८. दानाको दर
    formData.totalExpense,      // ९. कुल खर्च 
    formData.otherExpense,      // १०. अन्य खर्च
    formData.rate,              // ११. मासु को दर
    formData.soldKg,            // १२. बेचेको (Kg)
    formData.totalRevenue,      // १३. कुल आम्दानी
    formData.netProfitLoss,     // १४. कुल नाफा/नोक्सान 
    formData.workers,           // १५. कामदार
    formData.medicine,          // १६. औषधि
    formData.details            // १७. विवरण
  ];

  try {
    sheet.appendRow(rowData);
    // हरेक पटक डाटा इन्ट्री भएपछि कोलमको साइज मिलाउने
    sheet.autoResizeColumns(1, rowData.length); 
    return { status: 'success', message: 'डाटा सफलतापूर्वक बचत भयो।' };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: 'डाटा बचत गर्ने क्रममा त्रुटि: ' + e.message };
  }
}
