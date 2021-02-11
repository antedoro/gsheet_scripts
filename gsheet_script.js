// @ts-nocheck
/** @OnlyCurrentDoc */
// This Google apps Script make some work on Google GSheet sperdsheet like insert, clear, etc, row on archive from sheet form
// author V. Antedoro (www.antedoro.it)

//****GLOBALS****
const SPREADSHEET_ID = "insert spreadsheetID";
const DASHBOARD_V1_SHEET = "Dashboard";
const DASHBOARD_V2_SHEET = "Dashboard v2";
const ARCHIVE_SHEET = "Archive";
const LISTS_SHEET = "Lists";

function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menubuttons = [ {name: "Open Archive", functionName: "openArchive"},
                  {name: "GoTo Dashboard", functionName: "backToDashboard"},
                  {name: "Clear Form", functionName: "clearForm"},
                  {name: "Insert Order", functionName: "insert_order"}];
    ss.addMenu("Custom", menubuttons);
} // note you also have to have functions called clearForm and highFive as list below

//****MOVING TO SHEETS****
function goTosheet(sheet_name) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sheet_name), true);
};

function openArchive() {
  goTosheet(ARCHIVE_SHEET);  
};

function backToDashboard() {
  goTosheet(DASHBOARD_V2_SHEET);  
};

//****FORM MANAGEMENT****
function clearForm() {
  //Clear form on Dashboard2
  var sheet = SpreadsheetApp.getActive().getSheetByName(DASHBOARD_V2_SHEET);
  sheet.getRange('D34:D39').clearContent();
  sheet.getRange('D48:D49').clearContent();
  sheet.getRange('C52:E55').clearContent();
  sheet.getRange('C52:E55').clearContent();
  sheet.getRange('C56:E56').clearContent();

  SpreadsheetApp.flush();
  Utilities.sleep(10000); // You have 10 seconds to check that the cell has cleared
}

//experimental function not used
function insertRowBefore(sheet, rowIndex, rowData) {
  sheet.insertRowBefore(rowIndex);
  sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
}

//This function copy format from firstrow to lastrow
function formatting_row(sheet_name, rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheet_name);

  //Copy formatting to last row from reference row
  sheet.getRange(rowIndex,1).activate();
  sheet.getRange(2, 1,1,36).copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
};


// Core function: inserte formdata to archive gsheet 
function insert_order() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var copySheet = ss.getSheetByName(DASHBOARD_V2_SHEET);
  var pasteSheet = ss.getSheetByName(ARCHIVE_SHEET);
  var lock = LockService.getScriptLock(); 

  lock.waitLock(4000); // lock 4 seconds
  // Getting all values from sheet form
  var source = [];
  source.push([copySheet.getRange("D34").getValues()]);
  source.push([copySheet.getRange("D35").getValues()]);
  source.push([copySheet.getRange("D36").getValues()]);
  source.push([copySheet.getRange("D37").getValues()]);
  source.push([copySheet.getRange("D41").getValues()]);
  source.push([copySheet.getRange("D38").getValues()]);
  source.push('EUR');
  source.push([copySheet.getRange("D40").getValues()]);
  source.push([copySheet.getRange("D39").getValues()]);
  source.push([copySheet.getRange("D44").getValues()]);
  source.push([copySheet.getRange("E44").getValues()]);
  source.push([copySheet.getRange("D42").getValues()]);
  source.push([copySheet.getRange("E42").getValues()]);
  source.push([copySheet.getRange("D43").getValues()]);
  source.push([copySheet.getRange("E43").getValues()]);
  source.push([copySheet.getRange("D45").getValues()]); //RiskReward ratio
  source.push([copySheet.getRange("D46").getValues()]);
  source.push([copySheet.getRange("D47").getValues()]);
  source.push([copySheet.getRange("D48").getValues()]); //Venduto a
  source.push([copySheet.getRange("D50").getValues()]);
  source.push([copySheet.getRange("D49").getValues()]);
  source.push([copySheet.getRange("D51").getValues()]);
  source.push([copySheet.getRange("E51").getValues()]);
  source.push([copySheet.getRange("D52").getValues()]);
  source.push([copySheet.getRange("D53").getValues()]);
  source.push([copySheet.getRange("D54").getValues()]);
  source.push([copySheet.getRange("D55").getValues()]);
  source.push([copySheet.getRange("C56:E56").getValues()]);

  // Get last row on archive sheet
  var currentRow = pasteSheet.getLastRow()

  // setting a particular formula on first column ID i
  var sourceFormulas = pasteSheet.getRange(currentRow,1).getFormulasR1C1(); //formula copy
  currentRow++; //goto next row
  var targetRange = pasteSheet.getRange(currentRow, 1);
  targetRange.setFormulasR1C1(sourceFormulas); //pasting formula

  // Now setting values cell per cell
  for (var i = 0; i < source.length; i++) {
  pasteSheet.getRange(currentRow,i+2).setValue(source[i]);
  }

  //Copy formatting to last row from reference row
  formatting_row(ARCHIVE_SHEET, currentRow)

  lock.releaseLock();

  // clear source values
  //source.clearContent();

  //Browser.msgBox('Success copy!');
}

//Function to delete a row
function delete_row() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ARCHIVE_SHEET);
  var lock = LockService.getScriptLock();

  lock.waitLock(3000); // lock 3 seconds
  for (var i = sheet.getLastRow(); i>=2; i--) {
    if(sheet.getRange(i, 26).getValue() == 'ORDINE'){    
      sheet.deleteRow(i); }
  }
  lock.releaseLock();

}



//****REALTIME CRYPTO VALUE****
/** Imports JSON data to your spreadsheet Ex: IMPORTJSON("http://myapisite.com","city/population")
* @param url URL of your JSON data as string
* @param xpath simplified xpath as string
* @customfunction
*/
//not working
function IMPORTJSON(url,xpath){ 
  try{
    // /rates/EUR
    var res = UrlFetchApp.fetch(url);
    var content = res.getContentText();
    var json = JSON.parse(content);
    
    var patharray = xpath.split("/");
    //Logger.log(patharray);
    
    for(var i=0;i<patharray.length;i++){
      json = json[patharray[i]];
    }
    
    //Logger.log(typeof(json));
    
    if(typeof(json) === "undefined"){
      return "Node Not Available";
    } else if(typeof(json) === "object"){
      var tempArr = [];
      
      for(var obj in json){
        tempArr.push([obj,json[obj]]);
      }
      return tempArr;
    } else if(typeof(json) !== "object") {
      return json;
    }
  }
  catch(err){
      return "Error getting data";  
  }
  
}

//not working
function GetCryptoValue(crypto) {
  var url = "https://api.coinmarketcap.com/v1/ticker/" + crypto + "/?convert=EUR"; 
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  return data[0].price_eur;
}