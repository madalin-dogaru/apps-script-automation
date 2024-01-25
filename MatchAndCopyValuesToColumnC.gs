/*
 * Will match and copy values in a Google Sheet. 
 * It first creates a lookup from values in columns E and G, then iterates through column A, 
 * checking for matches in the lookup. When a match is found, the corresponding value from 
 * column G is copied into a new array, which is eventually written back to column C. 
 * The script also sorts the rows based on these new values in column C, providing 
 * an organized view of the matched and updated data. This functionality is especially 
 * useful for data reconciliation and organization tasks where matching and sorting 
 * based on specific criteria are required.
 */

function matchAndCopyValues() {
  // Get the active Google Sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Fetch values in columns A, E, and G for the first 200 rows
  var columnA = sheet.getRange("A1:A200").getValues();
  var columnE = sheet.getRange("E1:E200").getValues();
  var columnG = sheet.getRange("G1:G200").getValues();

  // Create an object to store values from column E and their corresponding values from column G
  var lookup = {};
  
  for (var i = 0; i < columnE.length; i++) {
    var valueE = columnE[i][0];
    var valueG = columnG[i][0];
    
    if (valueE) {
      lookup[valueE] = valueG;
    }
  }

  // Create an array to store new values for column C
  var newColumnC = [];
  
  // Loop through each value in column A
  for (var j = 0; j < columnA.length; j++) {
    var valueA = columnA[j][0];
    
    // Initialize new C value as empty (it will stay empty if there's no match)
    var newCValue = [""];
    
    if (valueA && lookup.hasOwnProperty(valueA)) {
      newCValue = [lookup[valueA]];
    }
    
    newColumnC.push(newCValue);
  }

  // Clear the original column C for the first 200 rows
  sheet.getRange("C1:C200").clearContent();

  // Write the new values back into column C
  sheet.getRange(1, 3, newColumnC.length, 1).setValues(newColumnC);

  // Sort the rows based on the values in column C from highest to lowest
  sheet.getRange("A1:C200").sort({column: 3, ascending: false});
}