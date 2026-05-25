
// ====================================================================
// १. कन्फिगरेसन (Configuration)
// ====================================================================

// ✅ तपाईंको स्प्रेडसिट ID हरू यहाँ राख्नुहोस्। (पहिलो ID अनिवार्य)
const SHEET_IDS = [
  '129T76GujKhJkk6wxHh8HIMcbzSofCjXvgatI-O_0Agc', 
  'YOUR_SECOND_SPREADSHEET_ID_HERE' // यदि आवश्यक छैन भने खाली छोड्नुहोस् वा हटाउनुहोस्
]; 

const MAX_ROWS_PER_SHEET = 33; // ग्राहक पानामा अधिकतम रो संख्या

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 
                     'July', 'August', 'September', 'October', 'November', 'December'];

// ====================================================================
// २. वेब एप (Web App)
// ====================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('बर पिपल कृषि सहकारी संस्था') 
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no') 
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// ====================================================================
// ३. मासिक ट्रिगर सेटअप (Setup Monthly Trigger)
// ====================================================================

/**
 * प्रत्येक महिनाको १ तारिखमा यो फङ्क्शन चल्ने गरी ट्रिगर सेट गर्छ।
 * ⚠️ यसलाई स्क्रिप्ट चलाएपछि एक पटक म्यानुअल रुपमा चलाउनुहोस् (Run > Run function > setupMonthlyInterestTrigger)।
 */
function setupMonthlyInterestTrigger() {
  // पहिलेको ट्रिगर हटाउने
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'applyMonthlyInterest') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // प्रत्येक महिनाको १ तारिख (Midnight) मा चल्ने नयाँ ट्रिगर सेट गर्ने
  ScriptApp.newTrigger('applyMonthlyInterest')
      .timeBased()
      .onDayOfMonth(1)
      .atHour(0) // मध्यरात १२ बजे
      .create();
  
  Logger.log("Monthly interest trigger set up for the 1st day of the month.");
}

// ====================================================================
// ४. पानाको नाम र विभाजन व्यवस्थापन (Sheet Naming and Segmentation)
// ====================================================================

/**
 * मास्टर पानाको लागि मासिक नाम (जस्तै: 25November11) उत्पादन गर्छ।
 */
function getMasterSheetName(dateString) {
  const date = new Date(dateString);
  const year = date.getFullYear() % 100; 
  const monthIndex = date.getMonth();
  const monthName = MONTH_NAMES[monthIndex];
  const monthNumber = monthIndex + 1;

  return ${year}${monthName}${monthNumber}; 
}

/**
 * ग्राहकको लागि सही पाना (Sheet) फेला पार्छ (जस्तै: Rohit, Rohit1, Rohit2)।
 */
function getOrCreateCustomerSheet(ss, baseName) {
    let sheetIndex = 0;
    let sheetName = baseName;
    let sheet = ss.getSheetByName(sheetName);

    // अन्तिम पाना फेला पार्न लुप गर्ने
    while (sheet && sheet.getLastRow() >= MAX_ROWS_PER_SHEET + 1) { // +1 हेडर रो को लागि
        sheetIndex++;
        // index 0 भए नाममा नम्बर जोडिँदैन (e.g. Rohit), index 1 भए Rohit1
        sheetName = baseName + (sheetIndex > 0 ? sheetIndex : ''); 
        sheet = ss.getSheetByName(sheetName);
        if (sheetIndex > 100) break; 
    }
    
    // यदि पाना भेटिएन वा नयाँ पाना आवश्यक छ भने सेटअप गर्ने
    if (!sheet) {
        sheet = setupSheet(ss, sheetName); 
    }
    
    return sheet;
}

// ====================================================================
// ५. पाना सेटअप र फॉर्मेटिंग (Sheet Setup and Formatting)
// ====================================================================

// ✅ १२ स्तम्भको हेडर
const HEADERS = [
  "मिति", "दिन", "ग्राहकको नाम", "आजको जम्मा रकम", "ऋणमा काटिने रकम", 
  "नयाँ कुल जम्मा", "बाँकी ऋण", "कुल लिएको ऋण", 
  "ब्याज लाग्ने रकम", "ब्याज प्रतिशत (%)", "गणना गरिएको ब्याज", "विवरण"
];

function setupSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
      sheet = ss.insertSheet(sheetName);
  }
  
  // हेडर रो खाली छ भने मात्र हेडर सेट गर्ने
  if (sheet.getLastRow() === 0) {
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setValues([HEADERS])
                 .setFontWeight('bold')
                 .setBackground('#f0f8ff')
                 .setHorizontalAlignment('center');
  }

  autoResizeColumnsAndCenter(sheet);
  
  return sheet;
}

function autoResizeColumnsAndCenter(sheet) {
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    
    sheet.autoResizeColumns(1, lastCol);

    for (let i = 1; i <= lastCol; i++) {
        sheet.setColumnWidth(i, 120); 
    }
    
    if (lastRow > 1) {
        // Data rows लाई center alignment दिने
        sheet.getRange(2, 1, lastRow - 1, HEADERS.length).setHorizontalAlignment('center');
    }
}

// ====================================================================
// ६. फारम प्रशोधन (Process Debt Form) 
// ====================================================================

function processDebtForm(formData) {
  try {
    const row = [
      formData.date,              
      ['आइतबार','सोमबार','मंगलबार','बुधबार','बिहीबार','शुक्रबार','शनिबार'][new Date(formData.date).getDay()], 
      formData.customerName,      
      formData.todayDeposit,      
      formData.todayDeduct,       
      formData.newTotalDeposit,   
      formData.remainingDebt,     
      formData.totalLoan,         
      formData.interestAmount,    
      formData.interestRate || 0,  
      formData.calculatedInterest, 
      formData.details            
    ];

    const baseCustomerName = formData.customerName.trim(); 
    const masterSheetName = getMasterSheetName(formData.date);
    
    // प्रत्येक स्प्रेडसिट ID मा डाटा प्रविष्ट गर्ने
    for (const sheetId of SHEET_IDS) {
        if (!sheetId || sheetId === 'YOUR_SECOND_SPREADSHEET_ID_HERE') continue; 

        const ss = SpreadsheetApp.openById(sheetId);
        
        // १. ग्राहक पाना सेटअप र इन्ट्री
        let customerSheet = getOrCreateCustomerSheet(ss, baseCustomerName);
        customerSheet.appendRow(row);
        autoResizeColumnsAndCenter(customerSheet);
        
        // २. मास्टर पाना सेटअप र इन्ट्री
        let masterSheet = setupSheet(ss, masterSheetName); 
        masterSheet.appendRow(row);
        autoResizeColumnsAndCenter(masterSheet);
    }
    
    return { status: 'success' };
  } catch (e) {
    Logger.log(Error in processDebtForm: ${e});
    return { status: 'error', message: e.message };
  }
}

// ====================================================================
// ७. मासिक ब्याज गणना र जोड्ने तर्क (Monthly Interest Application Logic)
// ====================================================================

function applyMonthlyInterest() {
    Logger.log("Starting monthly interest application...");

    const validSheetId = SHEET_IDS.find(id => id && id !== 'YOUR_SECOND_SPREADSHEET_ID_HERE');
    if (!validSheetId) {
        Logger.log("No valid Sheet ID found for monthly interest application.");
        return;
    }
    
    const ss = SpreadsheetApp.openById(validSheetId);
    
    // ग्राहक पानाहरू पत्ता लगाउने (मासिक पानाहरू बाहेक)
    const allSheets = ss.getSheets();
    const customerSheets = allSheets.filter(sheet => {
        const sheetName = sheet.getName();
        // मास्ट पानाको नाम पहिचान गर्ने रेगुलर अभिव्यक्ति
        return !sheetName.match(/^(2[4-9]|3[0-9])(January|February|March|April|May|June|July|August|September|October|November|December)\d+$/i);
    });

    const today = new Date();
    const dateString = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    const dayName = ['आइतबार','सोमबार','मंगलबार','बुधबार','बिहीबार','शुक्रबार','शनिबार'][today.getDay()];

// स्तम्भ सूचकांकहरू (0-आधारित)
    const COL_CUSTOMER_NAME = 2;   
    const COL_NEW_DEPOSIT = 5;     
    const COL_REMAINING_DEBT = 6;  
    const COL_TOTAL_LOAN = 7;      
    const COL_INTEREST_RATE = 9;   

    for (const sheet of customerSheets) {
        if (sheet.getLastRow() <= 1) continue;
        
        const lastRow = sheet.getLastRow();
        const data = sheet.getRange(lastRow, 1, 1, HEADERS.length).getValues()[0];
        
        const customerName = data[COL_CUSTOMER_NAME];
        const lastRemainingDebt = parseFloat(data[COL_REMAINING_DEBT]) || 0;
        const lastNewTotalDeposit = parseFloat(data[COL_NEW_DEPOSIT]) || 0; 
        const lastTotalLoan = parseFloat(data[COL_TOTAL_LOAN]) || 0; 
        
        const interestRate = parseFloat(data[COL_INTEREST_RATE]) || 0; 

        if (lastRemainingDebt > 0 && interestRate > 0) {
            
            // साधारण मासिक ब्याज गणना
            const monthlyInterest = (lastRemainingDebt * interestRate * (1 / 12)) / 100;
            
            // ⚠️ ब्याजलाई कुल लिएको ऋण र बाँकी ऋण दुवैमा जोड्ने
            const newTotalLoan = lastTotalLoan + monthlyInterest; 
            const newRemainingDebt = lastRemainingDebt + monthlyInterest; 

            Logger.log(Applying interest for ${customerName} (Sheet: ${sheet.getName()}));
            
            // नयाँ रेकर्ड पङ्क्ति (मासिक ब्याज इन्ट्री) सिर्जना गर्ने
            const newRow = [
                dateString,
                dayName,
                customerName,
                0, // आजको जम्मा रकम
                0, // ऋणमा काटिने रकम
                lastNewTotalDeposit.toFixed(2), // नयाँ कुल जम्मा (अपरिवर्तित)
                newRemainingDebt.toFixed(2),    // नयाँ बाँकी ऋण
                newTotalLoan.toFixed(2),        // नयाँ कुल लिएको ऋण
                lastRemainingDebt,              // ब्याज लाग्ने रकम
                interestRate,                   // ब्याज प्रतिशत
                monthlyInterest.toFixed(2),     // गणना गरिएको ब्याज
                "मासिक ब्याज जोडिएको (Auto-Applied)"
            ];
            
            sheet.appendRow(newRow);
            autoResizeColumnsAndCenter(sheet);
            
            // मास्टर शीटमा इन्ट्री पनि गर्ने
            const masterSheetName = getMasterSheetName(dateString);
            let masterSheet = setupSheet(ss, masterSheetName); 
            masterSheet.appendRow(newRow);
            autoResizeColumnsAndCenter(masterSheet);
        }
    }
    Logger.log("Monthly interest application finished.");
}

// ====================================================================
// ८. अघिल्लो रेकर्ड तान्ने (Customer Previous Record - Segmentation Awareness)
// ====================================================================

function getCustomerPreviousRecord(customerName, selectedDate) {
  try {
    const validSheetId = SHEET_IDS.find(id => id && id !== 'YOUR_SECOND_SPREADSHEET_ID_HERE');
    if (!validSheetId) throw new Error("No valid Sheet ID found for fetching record.");

    const ss = SpreadsheetApp.openById(validSheetId); 
    const baseName = customerName.trim();
    
    const defaultResponse = { 
        previousRemainingDebt: 0.00, 
        previousNewTotalDeposit: 0.00, 
        previousTotalLoan: 0.00 
    };

    const targetDate = new Date(selectedDate);
    const targetDateOnly = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate()).getTime();
    
    let lastFoundRecord = null;
    let sheetIndex = 0;
    
    // ⭐ सुधार: अब स्तम्भ A देखि L सम्म (12 स्तम्भ) डाटा तान्ने।
    const DATA_COLUMN_COUNT = 12; 

    while (true) {
        const sheetName = baseName + (sheetIndex > 0 ? sheetIndex : '');
        const sheet = ss.getSheetByName(sheetName);
        
        if (!sheet || sheet.getLastRow() <= 1) {
            if (sheetIndex === 0 && !sheet) return defaultResponse;
            if (!sheet) break;
        }

const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
            // अन्तिम रो भन्दा एक कम सम्म मात्र हेर्ने
            const data = sheet.getRange(2, 1, lastRow - 1, DATA_COLUMN_COUNT).getValues();
            
            for (let i = data.length - 1; i >= 0; i--) {
                const rowDate = new Date(data[i][0]);
                const rowDateOnly = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate()).getTime();
                
                if (rowDateOnly < targetDateOnly) {
                    const record = { 
                        date: rowDateOnly,
                        previousNewTotalDeposit: parseFloat(data[i][5]) || 0, // Col F (Index 5)
                        previousRemainingDebt: parseFloat(data[i][6]) || 0,   // Col G (Index 6)
                        previousTotalLoan: parseFloat(data[i][7]) || 0        // Col H (Index 7)
                    };
                    
                    if (!lastFoundRecord || record.date > lastFoundRecord.date) {
                        lastFoundRecord = record;
                    }
                    // अघिल्लो रेकर्ड भेटिएपछि यस पानाको जाँच रोक्ने
                    break; 
                }
            }
        }
        
        if (lastRow < MAX_ROWS_PER_SHEET + 1 && sheetIndex > 0) break;
        
        sheetIndex++;
        if (sheetIndex > 100) break;
    }
    
    if (!lastFoundRecord) return defaultResponse;
    
    return {
        previousNewTotalDeposit: lastFoundRecord.previousNewTotalDeposit,
        previousRemainingDebt: lastFoundRecord.previousRemainingDebt,
        previousTotalLoan: lastFoundRecord.previousTotalLoan
    };
    
  } catch (e) {
    Logger.log(Error in getCustomerPreviousRecord for ${customerName}: ${e});
    return { previousRemainingDebt: 0.00, previousNewTotalDeposit: 0.00, previousTotalLoan: 0.00, error: e.message };
  }
}
