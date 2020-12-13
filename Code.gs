/****************************************************************************************************************************************
*
* Save the value at G2 to the next empty cell in Column D.
*
****************************************************************************************************************************************/

function saveValue() {
  
  //  Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var g2Range = sheet.getRange("G2");
  var g2RangeValue = g2Range.getDisplayValue();
  var colDRange = sheet.getRange("D:D");
  
  //  Get values from Column D, only count String values (no empty/blank values), return length
  var lastRowColD = colDRange.getValues().filter(String).length;
  
  //  append value to next blank cell in ColD
  sheet.getRange(lastRowColD + 1, colDRange.getColumn()).setValue(g2RangeValue);
}
