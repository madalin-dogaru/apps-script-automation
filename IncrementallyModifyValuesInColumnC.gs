/*
 * Will incrementally modify values in column C of the 
 * "mitre_technique_scoring" sheet in a Google Spreadsheet. It specifically targets cells 
 * with a value of 0 and updates them with a decreasing series of values, starting from 990 
 * and decrementing by 9 with each successive cell. This functionality is useful for 
 * automatically adjusting values in a dataset, particularly in scenarios where a 
 * sequential decremental change is required based on certain conditions.
 */

function modifyColumnC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("mitre_technique_scoring");
  var data = sheet.getRange("C2:C708").getValues();
  
  var addedValue = 990;
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === 0) {
      sheet.getRange(i + 2, 3).setValue(addedValue);
      addedValue -= 9;
    }
  }
}