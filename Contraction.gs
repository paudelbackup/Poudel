// प्रारम्भिक डाटा लोड गर्ने (आयोजना र सामानहरूको सूची)
function getInitialData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName("Settings") || createSettingsSheet(ss);
  
  var contracts = settingsSheet.getRange("A2:A" + settingsSheet.getLastRow()).getValues().flat().filter(String);
  var items = settingsSheet.getRange("B2:B" + settingsSheet.getLastRow()).getValues().flat().filter(String);
  var companyName = settingsSheet.getRange("C2").getValue() || "KHADKA CONSTRUCTION";
  
  return {
    contracts: contracts.length ? contracts : ["मुख्य सडक आयोजना", "भवन निर्माण ठेक्का"],
    items: items.length ? items : ["ढुङ्गा", "बालुवा", "सिमेन्ट", "रड", "ईट्टा"],
    companyName: companyName
  };
}

// सेटिङहरू अपडेट गर्ने
function updateSettings(type, list) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Settings") || createSettingsSheet(ss);
  
  if (type === 'company') {
    sheet.getRange("C2").setValue(list[0]);
  } else if (type === 'contract') {
    sheet.getRange("A2:A" + Math.max(2, sheet.getLastRow())).clearContent();
    if(list.length) sheet.getRange(2, 1, list.length, 1).setValues(list.map(v => [v]));
  } else if (type === 'item') {
    sheet.getRange("B2:B" + Math.max(2, sheet.getLastRow())).clearContent();
    if(list.length) sheet.getRange(2, 2, list.length, 1).setValues(list.map(v => [v]));
  }
  return "OK";
}

function createSettingsSheet(ss) {
  var sheet = ss.insertSheet("Settings");
  sheet.getRange("A1:C1").setValues([["Contracts", "Items", "CompanyName"]]).setFontWeight("bold");
  sheet.hideSheet();
  return sheet;
}

// नेपाली अंकलाई अंग्रेजी स्ट्रिङ वा नम्बरमा बदल्ने फंक्शन
function nepToEngNum(nepNumStr) {
  if(!nepNumStr) return "0";
  var nep = ['०','१','२','३','४','५','६','७','८','९'];
  var eng = ['0','1','2','3','4','5','6','7','8','9'];
  return nepNumStr.toString().split('').map(function(ch) {
    var idx = nep.indexOf(ch);
    return idx > -1 ? eng[idx] : ch;
  }).join('');
}

// अंग्रेजी नम्बरलाई नेपाली अंक स्ट्रिङमा बदल्ने फंक्शन
function engToNepNum(engNum) {
  if(engNum === undefined || engNum === null) return "०";
  var nep = ['०','१','२','३','४','५','६','७','८','९'];
  var eng = ['0','1','2','3','4','5','6','7','8','9'];
  return engNum.toString().split('').map(function(ch) {
    var idx = eng.indexOf(ch);
    return idx > -1 ? nep[idx] : ch;
  }).join('');
}

function getNepaliMonthName(monthNum) {
  var stdMonths = ["वैशाख", "जेठ", "असार", "साउन", "भदौ", "असोज", "कात्तिक", "मंसिर", "पुस", "माघ", "फागुन", "चैत"];
  return stdMonths[parseInt(monthNum, 10) - 1] || "महिना";
}

// फारम सुरक्षित गर्ने मुख्य फंक्शन
function saveEntry(d) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var engDateStr = nepToEngNum(d.nepaliDate);
    var dateParts = engDateStr.split('/');
    if(dateParts.length < 3) return "त्रुटि: नेपाली मितिको ढाँचा मिलेन";
    
    var shortYear = dateParts[0].slice(-2); 
    var nepShortYear = engToNepNum(shortYear); 
    var monthName = getNepaliMonthName(dateParts[1]);
    
    var sheetName = nepShortYear + "-" + monthName; 
    
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        "मिति (अंग्रेजी)", "मिति (नेपाली)", "आयोजना (ठेक्का)", "प्रकार", 
        "विवरण/सामान", "परिमाण", "रकम", "मिस्त्री नाम", "मिस्त्री संख्या", 
        "मिस्त्री दर", "लेबर नाम", "लेबर संख्या", "लेबर दर", "कैफियत", "फोटो लिङ्क"
      ]);
      sheet.getRange("A1:O1").setFontWeight("bold").setBackground("#f1f5f9");
    }
    
    var imgUrl = "";
    if (d.imageFile && d.imageFile.includes("base64,")) {
      var contentType = d.imageFile.split(",")[0].split(":")[1].split(";")[0];
      var base64Data = d.imageFile.split(",")[1];
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, "Receipt_" + engDateStr.replace(/\//g,"-") + ".jpg");
      var folder;
      var folders = DriveApp.getFoldersByName("Khadka_Construction_Photos");
      if (folders.hasNext()) { folder = folders.next(); } else { folder = DriveApp.createFolder("Khadka_Construction_Photos"); }
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imgUrl = file.getUrl();
    }
    
    sheet.appendRow([
      d.englishDate,
      d.nepaliDate,
      d.contract,
      d.type,
      d.item,
      d.qty,
      engToNepNum(d.amt),
      d.mistryName,
      d.mistryQty,
      d.mistryRate,
      d.laborName,
      d.laborQty,
      d.laborRate,
      d.remarks,
      imgUrl
    ]);
    
    var lastRow = sheet.getLastRow();
    if(lastRow > 2) {
      var range = sheet.getRange(2, 1, lastRow - 1, 15);
      range.sort({column: 1, ascending: true});
    }
    
    return "OK";
  } catch(e) {
    return "त्रुटि: " + e.toString();
  }
}

// कुल समरी गणना गर्ने फंक्शन
function getContractSummary(contract, targetName, isLaborMode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  var matExp = 0, labExp = 0, savings = 0, personBalance = 0; 

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheet.getName() === "Settings") continue;
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) continue;
    
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (row[2] === contract) {
        var type = row[3]; 
        var amt = Number(nepToEngNum(row[6])) || 0;
        
        if (row[4] === "लेबर भुक्तानी" || row[7] !== "" || row[10] !== "") {
          var mQ = Number(nepToEngNum(row[8])) || 0;
          var mR = Number(nepToEngNum(row[9])) || 0;
          var lQ = Number(nepToEngNum(row[11])) || 0;
          var lR = Number(nepToEngNum(row[12])) || 0;
          var totalWageCalculated = (mQ * mR) + (lQ * lR);
          
          labExp += totalWageCalculated;
          savings -= totalWageCalculated;
        } else {
          if (type === "आम्दानी") { savings += amt; } 
          else if (type === "खर्च") { matExp += amt; savings -= amt; }
        }

        if (targetName && targetName.trim() !== "") {
          var searchName = targetName.trim().toLowerCase();
          
          if (isLaborMode) {
            var mName = String(row[7]).trim().toLowerCase();
            var lName = String(row[10]).trim().toLowerCase();
            
            if (mName === searchName) {
              var wage = (Number(nepToEngNum(row[8])) || 0) * (Number(nepToEngNum(row[9])) || 0);
              personBalance += (amt > 0 ? amt : 0) - wage; 
            }
            if (lName === searchName) {
              var wage = (Number(nepToEngNum(row[11])) || 0) * (Number(nepToEngNum(row[12])) || 0);
              personBalance += (amt > 0 ? amt : 0) - wage;
            }
          } else {
            var itemName = String(row[4]).trim().toLowerCase();
            if (itemName === searchName) {
              if (type === "आम्दानी") { personBalance += amt; } else { personBalance -= amt; }
            }
          }
        }

      }
    }
  }
  
  return { matExp: matExp, labExp: labExp, savings: savings, personBalance: personBalance };
}
