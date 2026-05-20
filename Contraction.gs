const SPREADSHEET_ID = "1XeQLRfiqrgpMrsWcdi3XllyLc3p7DJu5eHrFFk3PqNM"; 

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Khadka Construction Pro')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let settingsSheet = ss.getSheetByName("सामानहरू") || ss.insertSheet("सामानहरू");
  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.appendRow(["आयोजना / ठेक्का", "विवरण (Item)", "कम्पनीको नाम"]);
    settingsSheet.appendRow(["", "", "KHADKA CONSTRUCTION"]); // Default Name
  }
  
  const lastRow = settingsSheet.getLastRow();
  const lastCol = settingsSheet.getLastColumn();
  if(lastRow > 0) {
    // सुरुवाती सेटिङ सिटलाई पनि सेन्टर र प्रष्ट बनाउने
    settingsSheet.getRange(1, 1, lastRow, lastCol).setHorizontalAlignment("center").setVerticalAlignment("middle");
    settingsSheet.autoResizeColumns(1, lastCol);
    for(let i=1; i<=lastCol; i++) {
      settingsSheet.setColumnWidth(i, settingsSheet.getColumnWidth(i) + 40);
    }
  }

  const data = settingsSheet.getDataRange().getValues();
  let companyName = data.slice(1).map(r => r[2]).filter(Boolean)[0] || "KHADKA CONSTRUCTION";

  return { 
    contracts: data.slice(1).map(r => r[0]).filter(Boolean), 
    items: data.slice(1).map(r => r[1]).filter(Boolean),
    companyName: companyName
  };
}

function saveEntry(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const days = ["आइतबार", "सोमबार", "मंगलबार", "बुधबार", "बिहीबार", "शुक्रबार", "शनिबार"];
    const dayName = days[new Date(data.englishDate).getDay()];
    
    let fileLink = "फोटो छैन";
    if (data.imageFile && data.imageFile.includes("base64")) {
      const folder = DriveApp.getRootFolder();
      const bytes = Utilities.base64Decode(data.imageFile.split(',')[1]);
      const blob = Utilities.newBlob(bytes, "image/jpeg", "Bill_" + Date.now() + ".jpg");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileLink = '=HYPERLINK("' + file.getUrl() + '", "फोटो हेर्नुहोस्")';
    }

    const year = data.nepaliDate.split("/")[0];
    const targetSheetName = data.contract + " " + year;
    const targetSheet = getOrCreateSheet(ss, targetSheetName);
    
    let finalItem = (data.type === "आम्दानी" && !data.isLabor) ? "-" : data.item;

    if (data.isLabor) {
      const laborMaster = getOrCreateSheet(ss, "लेबर_मास्टर");
      const wage = (Number(data.mistryQty) * Number(data.mistryRate)) + (Number(data.laborQty) * Number(data.laborRate));
      
      const row = [
        data.nepaliDate, 
        dayName, 
        data.contract, 
        data.mistryName, 
        data.mistryQty, 
        data.laborName,  
        data.laborQty, 
        data.amt, 
        wage, 
        wage - data.amt, 
        data.remarks, 
        data.englishDate, 
        fileLink
      ];
      appendAndFormat(laborMaster, row);
      
      const pRow = [data.nepaliDate, dayName, data.contract, "लेबर भुक्तानी", "-", data.amt, "", "", "", "खर्च", data.remarks, fileLink, data.englishDate];
      appendWithFormula(targetSheet, pRow);
    } else {
      const master = getOrCreateSheet(ss, "मास्टर");
      const mRow = [data.nepaliDate, dayName, data.contract, finalItem, data.qty, data.amt, data.type, data.remarks, fileLink, data.englishDate];
      appendAndFormat(master, mRow);
      
      const pRow = [data.nepaliDate, dayName, data.contract, finalItem, data.qty, data.amt, "", "", "", data.type, data.remarks, fileLink, data.englishDate];
      appendWithFormula(targetSheet, pRow);
    }
    return "OK";
  } catch (e) { return "Error: " + e.toString(); }
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    let headers = name.includes("लेबर_मास्टर") ? 
      ["मिति", "बार", "ठेक्का", "मिस्त्रीको नाम", "मिस्त्री संख्या", "लेबरको नाम", "लेबर संख्या", "भुक्तानी", "कुल ज्याला", "बाँकी", "कैफियत", "AD", "फोटो"] :
      (name === "मास्टर" ? ["मिति", "बार", "ठेक्का", "विवरण", "परिमाण", "रकम", "प्रकार", "कैफियत", "फोटो", "AD"] :
      ["मिति", "बार", "ठेक्का", "विवरण", "परिमाण", "रकम", "कुल आम्दानी", "कुल खर्च", "नेट बचत", "प्रकार", "कैफियत", "फोटो", "AD"]);

    sheet.appendRow(headers);
    let headerRange = sheet.getRange(1, 1, 1, headers.length);
    // कमजोर आँखाले पनि स्पष्ट देख्ने बोल्ड अक्षर र #1e293b आकर्षक ब्याकग्राउन्ड
    headerRange.setFontWeight("bold")
                .setBackground("#1e293b")
                .setFontColor("white")
                .setHorizontalAlignment("center")
                .setVerticalAlignment("middle");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function appendAndFormat(sheet, row) {
  sheet.appendRow(row);
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(lastRow, 1, 1, row.length);
  
  // डाटालाई तेर्सो र ठाडो दुवै तर्फबाट ट्याक्क बीचमा (Center) पार्ने र बोर्डर दिने
  range.setHorizontalAlignment("center")
       .setVerticalAlignment("middle")
       .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  
  // केवल तेर्सो (Horizontal) चौडाइ मात्र बढाउने, भर्टिकल साइज नर्मल नै राख्ने
  sheet.autoResizeColumns(1, row.length);
  for(let i=1; i<=row.length; i++) {
    sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 40); // छेउछाउमा पर्याप्त ग्याप ताकि डाटा नछोपियोस्
  }
}

function appendWithFormula(sheet, rowData) {
  const lastRow = sheet.getLastRow() + 1;
  sheet.appendRow(rowData);
  
  sheet.getRange(lastRow, 7).setFormula(`=SUMIF($J$2:$J$${lastRow}, "आम्दानी", $F$2:$F$${lastRow})`);
  sheet.getRange(lastRow, 8).setFormula(`=SUMIF($J$2:$J$${lastRow}, "खर्च", $F$2:$F$${lastRow})`);
  sheet.getRange(lastRow, 9).setFormula(`=G${lastRow}-H${lastRow}`);
  
  const range = sheet.getRange(lastRow, 1, 1, rowData.length);
  
  // नयाँ गणना भएको रोलाई पनि पूर्ण रूपमा सेन्टर र प्रष्ट ग्याप दिने
  range.setHorizontalAlignment("center")
       .setVerticalAlignment("middle")
       .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
       
  sheet.autoResizeColumns(1, rowData.length);
  for(let i=1; i<=rowData.length; i++) {
    sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 40); // दायाँ-बायाँ ग्याप मात्र थपिने
  }
}

function updateSettings(type, list) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("सामानहरू") || ss.insertSheet("सामानहरू");
  
  if (type === 'company') {
    sheet.getRange(2, 3).setValue(list[0] || "KHADKA CONSTRUCTION");
  } else {
    const col = type === 'contract' ? 1 : 2;
    sheet.getRange(2, col, 1000, 1).clearContent();
    if (list.length > 0) sheet.getRange(2, col, list.length, 1).setValues(list.map(i => [i]));
  }
  
  sheet.getDataRange().setHorizontalAlignment("center").setVerticalAlignment("middle");
  return "Updated";
}

function getContractSummary(contractName, selectedEngDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const selDate = new Date(selectedEngDate);
  let income = 0, matExp = 0, labExp = 0;
  
  const master = ss.getSheetByName("मास्टर");
  if (master) {
    const data = master.getDataRange().getValues().slice(1);
    data.forEach(r => {
      if (r[2] === contractName && new Date(r[9]) <= selDate) {
        if (r[6] === "आम्दानी") income += Number(r[5] || 0);
        else matExp += Number(r[5] || 0);
      }
    });
  }
  const labor = ss.getSheetByName("लेबर_मास्टर");
  if (labor) {
    const data = labor.getDataRange().getValues().slice(1);
    data.forEach(r => {
      if (r[2] === contractName && new Date(r[11]) <= selDate) labExp += Number(r[7] || 0);
    });
  }
  return { matExp, labExp, savings: income - (matExp + labExp) };
}
