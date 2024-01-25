/*
 * Update scores for subtechniques in the 
 * "technique_value" sheet. It iterates through rows to identify main techniques and their 
 * corresponding subtechniques. For each subtechnique, the script calculates a new score by a 
 * formula that incorporates the score of its parent technique. This automated approach ensures 
 * consistency in score calculation for subtechniques based on their associated main technique scores.
 */

function updateSubtechniques() {
  // Open the Google Sheet and select the relevant sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("technique_value");
  
  // Get values from columns TECHNIQUE_ID and TACTICS_SCORE
  var techniqueIds = sheet.getRange("A2:A608").getValues();
  var tacticsScores = sheet.getRange("F2:F608").getValues();
  
  // Loop through each row to identify techniques and their corresponding subtechniques
  for (var i = 0; i < techniqueIds.length; i++) {
    var currentTechniqueId = techniqueIds[i][0];
    var currentTacticsScore = tacticsScores[i][0];
    
    // Check if the current row contains a technique ID and has a TACTICS_SCORE
    if (currentTechniqueId && currentTacticsScore) {
      
      // Loop through the following rows to find subtechniques for the current technique
      for (var j = i + 1; j < techniqueIds.length; j++) {
        var nextTechniqueId = techniqueIds[j][0];
        
        // If the next row is a subtechnique of the current technique
        if (nextTechniqueId.startsWith(currentTechniqueId + ".")) {
          
          // Construct the new formula for the subtechnique
          var newFormula = `=SUM(${currentTacticsScore}+50+20)/3`;
          
          // Set the new formula for the TE_SCORE column of the subtechnique
          sheet.getRange(j + 2, 4).setFormula(newFormula); // j + 2 because the array is 0-based and the range starts from row 2
          
        } else if (!nextTechniqueId.startsWith(currentTechniqueId)) {
          // Exit the inner loop if the next row is not a subtechnique
          break;
        }
      }
    }
  }
}