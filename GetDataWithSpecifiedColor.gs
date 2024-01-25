/*
 * Will filter and retrieve data from a Google Sheet based on cell 
 * background color. It examines cells in a specified column (E in this case), comparing their 
 * background color to a given color. When a match is found, the corresponding cell's data is 
 * added to an array. The function ultimately returns this array of filtered data. This script 
 * is particularly useful for extracting and working with data categorized or highlighted 
 * by specific colors, which is a common practice in visually organized spreadsheets.
 */

function getDataWithColorx(backgroundColor) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("E1:E" + sheet.getLastRow()).getValues();
  var colors = sheet.getRange("E1:E" + sheet.getLastRow()).getBackgrounds();
  
  var filteredData = [];
  
  for (var i = 0; i < data.length; i++) {
    if (colors[i][0] === backgroundColor) {
      filteredData.push([data[i][0]]);
    }
  }
  
  return filteredData;
}