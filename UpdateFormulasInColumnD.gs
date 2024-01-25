/*
 * Will update specific formulas in column D of the active sheet 
 * within a Google Spreadsheet. It searches for formulas matching a certain pattern 
 * (SUM of three terms divided by 3) and modifies them by applying a calculation 
 * that adjusts the terms based on certain multipliers. This automated modification 
 * of formulas is useful for dynamically recalibrating calculations, such as weighted 
 * averages or adjusted scores, based on new criteria or changed datasets.
 */

function updateFormulas() {
  // Get the active spreadsheet and the active sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  // Define the range for column D (assuming 1000 rows; adjust as needed)
  var range = sheet.getRange("D1:D1000");
  var formulas = range.getFormulas();
  
  // Define the regex pattern to look for
  var pattern = /^=SUM\((\d+(\.\d+)?)\+(\d+(\.\d+)?)\+(\d+(\.\d+)?)\)\/3$/;
  
  // Loop through all rows to find and update matching formulas
  for (var i = 0; i < formulas.length; i++) {
    var formula = formulas[i][0];
    var match = formula.match(pattern);
    
    if (match) {
      // Extract the value of X from the matched formula
      var X = parseFloat(match[1]);
      
      // Calculate the new formula
      var newFormula = `=SUM(${X}+(${X}*0.85)+(${X}*1.05))/3`;
      
      // Update the cell with the new formula
      sheet.getRange(i + 1, 4).setFormula(newFormula);
    }
  }
}