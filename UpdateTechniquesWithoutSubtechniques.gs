/*
 * Updates the scores of MITRE techniques in the 
 * "mitre_technique_scoring" sheet, specifically targeting those techniques that do not 
 * have associated subtechniques. It scans the sheet to identify such techniques and 
 * then recalculates their scores using a custom formula. The script ensures that techniques 
 * without subtechniques are scored appropriately, reflecting their standalone impact 
 * in the scoring model.
 */

function updateTechniquesWithoutSubtechniques() {
  // Open the Google Sheet and select the relevant sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("mitre_technique_scoring");
  
  // Get values from columns TECHNIQUE_ID and TACTICS_SCORE
  var techniqueIds = sheet.getRange("A2:A608").getValues();
  var tacticsScores = sheet.getRange("F2:F608").getValues();
  
  // Loop through each row to identify techniques
  for (var i = 0; i < techniqueIds.length; i++) {
    var currentTechniqueId = techniqueIds[i][0];
    var currentTacticsScore = tacticsScores[i][0];
    
    var hasSubtechnique = false;
    
    // Check if the current row contains a technique ID and has a TACTICS_SCORE
    if (currentTechniqueId && currentTacticsScore) {
      
      // Loop through the following rows to check for subtechniques for the current technique
      for (var j = i + 1; j < techniqueIds.length; j++) {
        var nextTechniqueId = techniqueIds[j][0];
        
        // If the next row is a subtechnique of the current technique
        if (nextTechniqueId.startsWith(currentTechniqueId + ".")) {
          hasSubtechnique = true;
          break;
        } else if (!nextTechniqueId.startsWith(currentTechniqueId)) {
          // Exit the inner loop if the next row is not a subtechnique
          break;
        }
      }
      
      // If no subtechnique was found
      if (!hasSubtechnique) {
        // Construct the new formula for the technique
        var newFormula = `=SUM(${currentTacticsScore}+20+90)/3`;
        
        // Set the new formula for the TE_SCORE column of the technique
        sheet.getRange(i + 2, 4).setFormula(newFormula); // i + 2 because the array is 0-based and the range starts from row 2
      }
    }
  }
}