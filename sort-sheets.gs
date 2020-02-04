// function: sorts sheets alphabetically
// author: Matthew Hall
// created: 04-Feb-20
// modified:

// global variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();
var ui = SpreadsheetApp.getUi();

function onOpen() {
  // creates menu to easily run script(s) from the spreadsheet
  ui.createMenu('Sort')
    .addItem('Alphabetically', 'alpha')
    .addItem('Reverse Alpha', 'ahpla')
    .addToUi();
}

// sort sheets alphabetically
function alpha() {
  // get the sheet names as an array
  var sheetNames = [];
  for (var i = 0; i < sheets.length; i++) {
    sheetNames.push(sheets[i].getName());
  }

  // sort the array
  sheetNames.sort();

  // reorder the sheets
  for (var i = 0; i < sheets.length; i++) {
    // only moving of active sheet is supported so we will cycle through the sorted
    // sheet names and make them active
    // then set the position based on the index i
    ss.setActiveSheet(ss.getSheetByName(sheetNames[i]));
    ss.moveActiveSheet(i + 1);
  }
}

// sort sheets reverse alphabetically
function ahpla() {

}
