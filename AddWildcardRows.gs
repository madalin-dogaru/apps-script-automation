/*
 * It dynamically inserts 'wildcard' rows into the "mitre-techniques" sheet 
 * of a Google Spreadsheet. It adds a new row after each technique row, provided the next row isn't 
 * a new technique. This is primarily used to facilitate structured data organization and maintain 
 * consistency in formula calculations across the sheet, particularly updating SUM formulas to 
 * include the newly inserted rows.
 */

function addWildcardRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("mitre-techniques");
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Start from row 2
  var formulas = sheet.getRange(2, 3, lastRow - 1, 1).getFormulas();
  
  for (var i = lastRow - 1; i >= 1; i--) { // Start from row 2 (i.e., index 1 in the array)
    var isTechnique = data[i - 1][0].indexOf(".") === -1 && data[i - 1][0].startsWith("T");
    var isNextNewTechnique = (i < lastRow - 1) ? (data[i][0].indexOf(".") === -1 && data[i][0].startsWith("T")) : true;

    if (isTechnique && !isNextNewTechnique) {
      sheet.insertRowAfter(i + 1);
      sheet.getRange(i + 2, 1).setValue("wildcard row");
      
      // Update the formula for the technique row
      if (formulas[i - 1][0] !== '') {
        var formula = formulas[i - 1][0];
        var updatedFormula = updateSumFormula(formula, i + 1);
        sheet.getRange(i + 1, 3).setFormula(updatedFormula);
      }
    }
  }
}

// Function to update the SUM formula to include the new row
function updateSumFormula(formula, newRowNumber) {
  var match = formula.match(/SUM\((C\d+):(C\d+)\)/);
  if (match) {
    var startRow = parseInt(match[1].substr(1));
    var endRow = parseInt(match[2].substr(1));
    startRow++; // Increment start row by 1 because we inserted a new row
    endRow++;  // Increment end row by 1 to include the new row
    return formula.replace(match[0], `SUM(C${startRow}:C${endRow})`);
  }
  return formula;
}