/*
 * Will update the 'sheet_1' sheet in Google Sheets with 
 * MITRE TTP data. It first reads technique data from the 'sheet_2' sheet, 
 * structuring it for ease of access, and then cross-references this data with the technique IDs
 * listed in the 'sheet_1' sheet. It updates the analysis sheet with values 
 * corresponding to 'Process Injection', 'LOTL', 'Memory Execution', 'Obfuscation', 
 * 'EDR Disabling', and 'Encrypted Channels' for each technique. The script uses regular 
 * expressions to ensure accurate matching and logs processing details for troubleshooting.
 */

function updateTechniqueData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_2Sheet = ss.getSheetByName("sheet_2");
  var analysisSheet = ss.getSheetByName("sheet_1");

  // Read data from 'sheet_2' sheet
  var sheet_2Data = sheet_2Sheet.getRange(1, 1, sheet_2Sheet.getLastRow(), 2).getValues();

  // Process and structure the data for easy access
  var structuredData = {};
  var idRegex = /^T\d+(\.\d+)?/; // Regular expression to match both technique and sub-technique formats
  
  for (var i = 0; i < sheet_2Data.length; i++) {
    var match = sheet_2Data[i][0].match(idRegex); // Extract the technique ID using regex
    if (match) {
      var techniqueId = match[0];
      console.log("Processing: " + techniqueId);
      if (i + 6 < sheet_2Data.length) {
        structuredData[techniqueId] = {
          'Process Injection': sheet_2Data[i + 1][1],
          'LOTL': sheet_2Data[i + 2][1],
          'Memory Execution': sheet_2Data[i + 3][1],
          'Obfuscation': sheet_2Data[i + 4][1],
          'EDR Disabling': sheet_2Data[i + 5][1],
          'Encrypted Channels': sheet_2Data[i + 6][1]
        };
      }
    }
  }

  // Read technique IDs from 'sheet_1' sheet
  var analysisData = analysisSheet.getRange(2, 1, analysisSheet.getLastRow(), 1).getValues();

  // Update 'sheet_1' sheet with the corresponding values
  for (var j = 0; j < analysisData.length; j++) {
    var technique = analysisData[j][0];
    if (structuredData[technique]) {
      analysisSheet.getRange(j + 2, 3).setValue(structuredData[technique]['Process Injection']);
      analysisSheet.getRange(j + 2, 4).setValue(structuredData[technique]['LOTL']);
      analysisSheet.getRange(j + 2, 5).setValue(structuredData[technique]['Memory Execution']);
      analysisSheet.getRange(j + 2, 6).setValue(structuredData[technique]['Obfuscation']);
      analysisSheet.getRange(j + 2, 7).setValue(structuredData[technique]['EDR Disabling']);
      analysisSheet.getRange(j + 2, 8).setValue(structuredData[technique]['Encrypted Channels']);
    } else {
      console.log("No match found for: " + technique);
    }
  }
}