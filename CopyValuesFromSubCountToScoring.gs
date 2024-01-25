/*
 * It will synchronize data between two sheets 
 * within a Google Spreadsheet: "scoring_mitre_technique" and "subcount_techniques". 
 * The script scans the "scoring_mitre_technique" sheet for techniques with a zero value 
 * in a specified column and then searches for these techniques in the "subcount_techniques" 
 * sheet. Once found, it copies corresponding values from the "subcount_techniques" sheet 
 * back to the "scoring_mitre_technique" sheet. This process is particularly useful for 
 * updating scoring metrics based on sub-technique counts or similar criteria.
 */

function copyValuesFromSubCount() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scoringSheet = ss.getSheetByName("scoring_mitre_technique");
  var subCountSheet = ss.getSheetByName("subcount_techniques");
  
  // Get the data ranges for both sheets
  var scoringRange = scoringSheet.getRange("C1:C" + scoringSheet.getLastRow());
  var subCountRange = subCountSheet.getRange("A1:C" + subCountSheet.getLastRow());
  
  // Get all values in column C of the scoring sheet
  var scoringValues = scoringRange.getValues();
  
  // Iterate through the scoring sheet to find techniques with a value of 0
  for (var i = 0; i < scoringValues.length; i++) {
    if (scoringValues[i][0] === 0) {
      // Get the corresponding technique name from the same row in column A
      var techniqueName = scoringSheet.getRange(i + 1, 1).getValue();
      
      // Search for the technique name in the subCount sheet
      var subCountData = subCountRange.getValues();
      for (var j = 0; j < subCountData.length; j++) {
        if (subCountData[j][0] === techniqueName) {
          // Copy the value from column C of the subCount sheet to the scoring sheet
          scoringSheet.getRange(i + 1, 3).setValue(subCountData[j][2]);
          break; // Break the loop once the value is found
        }
      }
    }
  }
}