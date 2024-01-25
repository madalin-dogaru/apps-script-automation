/*
 * Deletes rows in the active sheet of a Google Spreadsheet where the first 
 * column's value matches a specific condition. It identifies rows where the first column equals 
 * 'sum' and deletes them. The script operates from the bottom to the top of the sheet to 
 * ensure consistent row indexing during deletion.
 */
// delete all rows that have a specific value on the first column. 

function deleteRowsBasedOnCondition() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  let rowsToDelete = [];

  // Loop through all rows in the data
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === 'sum') {
      // Store the row number that needs to be deleted
      rowsToDelete.push(i + 1);
    }
  }

  // Delete rows from the bottom of the sheet to the top
  rowsToDelete.forEach(function(row) {
    sheet.deleteRow(row);
  });
}