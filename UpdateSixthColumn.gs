/*
 * Will selectively update the sixth column in the 
 * "mitre_technique_scoring" sheet. It targets specific rows based on predefined technique IDs 
 * (e.g., T1071, T1092, T1132, etc.) listed in the first column. If a row's first column matches 
 * one of these IDs, the script sets a specific value (10000) in the sixth column of that row. 
 * This is useful for batch updating specific data points across a large dataset in Google Sheets.
 */

 function updateSixthColumn() {
  // Access the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Access the sheet you want to modify
  var sheet = ss.getSheetByName("mitre_technique_scoring"); // Replace with your sheet name
  
  // Define the target values
  var targetValues = ["T1071", "T1092", "T1132", "T1001", "T1568", "T1573", "T1008", "T1105", "T1104", "T1095", "T1571", "T1572", "T1090", "T1219", "T1205", "T1102"];

  // Get the range of cells to check
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(1, 1, lastRow, 6); // Assuming the data starts at row 1
  
  // Fetch values from the range
  var values = range.getValues();
  
  // Loop through each row in the range
  for (var i = 0; i < values.length; i++) {
    // Check if the value in the first column is in the targetValues array
    if (targetValues.indexOf(values[i][0]) !== -1) {
      // Update the value in the 6th column of this row
      values[i][5] = 10000;
    }
  }
  
  // Update the sheet with the modified values
  range.setValues(values);
}
