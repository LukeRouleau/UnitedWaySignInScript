/* ==================== SIGN-IN FORM AUTOCOMPLETION SCRIPT ====================

This script was written to help reduce the time it takes for return visitors for a Family Fundamentals Class
to sign-in. Before this script, returning attendees would have to fill out the entire google form again,
exactly like they did the first time they attended the class.

Now, going combined with a change made to the Google Forms sheet
we returning attendees fill out a shorter form, and this script will fill out the rest of their informaiton in 
the spreadsheet according to the last time they attended the class.

Developer:
  Luke Rouleau
  
Date Developed:
  July 17th, 2022
================================================================================*/
// Global Varibles
// Grab the active sheet:
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];

// Grab the sheet data:
var range = sheet.getDataRange();
var values = range.getValues(); // a 2D array of all the data

// Grab the # corresponding to the autoCompleteColumn: AC: "Is this your first time attending this class?":
var autoCompleteColumn = range.getLastColumn() - 2;

// Grab the # corresponding to the verifyAutoCompleteColumn: AD: "Auto-Completed Success?":
var verifyAutoCompleteColumn = range.getLastColumn() - 1;

// Grab the three match-criteria columns, class, first, and last
var classCol = 1; // Which Class are you attending?
var firstNameCol = 2; // First Name
var lastNameCol = 4; // Last Name

// Grab the # corresponding to the last row in the sheet, i.e. the most recent response
var lastRow = range.getLastRow()-1;

// Grab the # corresponding to the first row in the sheet, i.e. the first response ever in row 2
var firstRow = 2;

// Add a button to the tool
function onOpen() {
  console.log("Opened the sheet.");
  ui.createMenu('Auto-Complete')
  .addItem('Run', 'autoComplete')
  .addToUi();
};

// Function to perform the auto-completion:
function autoComplete() {  
  // Loop over the rows in reverse:
  for (var row = lastRow; row > 0; row--) {
    // This holds the value in the current row of column AC: "Is this your first time attending this class?"
    var autoCompleteCellValue = values[row][autoCompleteColumn];
    
    // If we have hit empty cells while looping through the auto-complete row, we can exit:
    if (autoCompleteCellValue == ''){
      break;
    }
    // Now we have to call the auto complete function:
    else if (values[row][autoCompleteColumn] == "No, use my responses from last time") {
      // Only replace if we have not already replaced for them before:
      if(values[row][verifyAutoCompleteColumn] == '') {
        console.log(values[row][firstNameCol] + " " + values[row][lastNameCol] + " requested to autocomplete, beginning search for prior class visit.");
        var res = performCopy(values, row);
        if (res == -1){
          // Failure, mark it in the cell
          sheet.getRange(row+1, verifyAutoCompleteColumn+1).setValue('FAILED, NO MATCH.').setBackground("#ffcccb");
        }
        else{
          // Success, mark it in the cell
          sheet.getRange(row+1, verifyAutoCompleteColumn+1).setValue('Success!').setBackground("#90ee90");
        }
      }
    }
  }  
}

// Search in reverse order for a first name, last name, class match then copy the range over: 
function performCopy(values, row){
  var input = values[row];
  // Iterate in reverse starting from row-1 
  for (var searchRow = row-1; searchRow > 0; searchRow--){
    var search = values[searchRow];
    // Look for class match:
    if (input[classCol] == search[classCol]) {
      // Look for first name match (ignore whitespace & Upper case):
      if (input[firstNameCol].replace(/\s+/g, '').toLowerCase() == 
          search[firstNameCol].replace(/\s+/g, '').toLowerCase()) 
      {
        // Look for last name match:
        if (input[lastNameCol].replace(/\s+/g, '').toLowerCase() == 
            search[lastNameCol].replace(/\s+/g, '').toLowerCase())
        {
          var printRow = row + 1;
          var printSearchRow = searchRow + 1;
          console.log("Auto-complete Match Found: Copying " + input[firstNameCol] + " " + input[lastNameCol] + '\'s auto-complete info from row ' + printSearchRow + " to " + printRow);
          var copyRange = sheet.getRange(searchRow + 1, 6, 1,21);
          var destRange = sheet.getRange(row + 1, 6, 1,21);
          copyRange.copyTo(destRange);
          return 1; // Exit, we've done our job
        }
      }
    }
  }
  return -1;
}
