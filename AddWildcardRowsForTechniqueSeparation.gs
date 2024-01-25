/*
 * It dynamically inserts 'wildcard' rows between distinct MITRE 
 * technique entries to improve data organization and visualization. The script identifies 
 * each technique and adds a new row immediately after it, provided the following row 
 * isn't a new distinct technique. Additionally, it updates any existing SUM formulas 
 * in the sheet to include these newly added wildcard rows, thereby maintaining the 
 * accuracy of calculated values. This functionality is particularly useful for 
 * visually separating techniques and ensuring seamless integration of new data entries.
 */

function addWildcardRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("mitre_technique_scoring");
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); 
  var formulas = sheet.getRange(2, 3, lastRow - 1, 1).getFormulas();

  for (var i = lastRow - 1; i >= 1; i--) {
    var isTechnique = data[i - 1][0].indexOf(".") === -1 && data[i - 1][0].startsWith("T");
    var isNextNewTechnique = (i < lastRow - 1) ? (data[i][0].indexOf(".") === -1 && data[i][0].startsWith("T")) : true;

    if (isTechnique && !isNextNewTechnique) {
      sheet.insertRowAfter(i + 1);
      var newRowRange = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn());
      newRowRange.getCell(1, 1).setValue("wildcard row");
      newRowRange.setBackground("#FFFFFF");

      if (formulas[i - 1][0] !== '') {
        var formula = formulas[i - 1][0];
        var updatedFormula = updateSumFormula(formula);
        sheet.getRange(i + 1, 3).setFormula(updatedFormula);
      }
    }
  }
}

function updateSumFormula(formula) {
  var match = formula.match(/SUM\((C\d+):(C\d+)\)\/(\d+)/);
  if (match) {
    var startRow = parseInt(match[1].substr(1));
    var endRow = parseInt(match[2].substr(1)) + 1;
    var divisor = parseInt(match[3]) + 1;
    return `=SUM(C${startRow}:C${endRow})/${divisor}`;
  }
  return formula;
}