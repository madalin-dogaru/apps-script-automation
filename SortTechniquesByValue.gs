/*
 * Sorts techniques in a Google Spreadsheet based on specific values. 
 * It operates on the active sheet of the active spreadsheet, filtering out subtechniques 
 * and focusing on main techniques. The script identifies techniques using a regular expression 
 * to match technique IDs in column A and then sorts them by corresponding values in column E. 
 * The sorted list of techniques, along with their values, is then written back to the sheet, 
 * specifically in columns J and K. This functionality is particularly useful for organizing 
 * MITRE techniques by various metrics, such as impact, frequency, or other custom criteria.
 */

function sortTechniques() {
  // Get the active spreadsheet and the active sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  // Get data from column A and E
  var data = sheet.getRange("A1:E800").getValues();
  
  // Filter and map the data to keep only techniques (not subtechniques) and their values
  var filteredData = data.filter(function(row) {
    return /^[T]\d+$/.test(row[0]);  // Regex to match "T" followed by digits
  }).map(function(row) {
    return { technique: row[0], value: row[4] };  // 0 and 4 correspond to columns A and E
  });
  
  // Sort the data by the value in descending order
  filteredData.sort(function(a, b) {
    return b.value - a.value;
  });
  
  // Prepare the data to be set back to the sheet
  var sortedData = filteredData.map(function(row) {
    return [row.technique, row.value];
  });
  
  // Write the sorted data to columns J and K
  sheet.getRange(1, 10, sortedData.length, 2).setValues(sortedData);  // 10 and 11 correspond to columns J and K
}