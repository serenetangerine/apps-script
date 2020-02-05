// function: searches all sheets for a user input string
// author: Matthew Hall
// created: 04-Feb-20
// last modified: 04-Feb-20 Matthew Hall

// TODO:
// generalize script for arbitrary search
// find next instance
// automatically make result the active cell/sheet

// global variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();
var ui = SpreadsheetApp.getUi();

function onOpen() {
  // creates menu to easily run script(s) from the spreadsheet
  ui.createMenu('Scripts')
    .addItem('Search', 'prompt')
    .addToUi();
}

function prompt() {
  // generate prompt for search text
  var query = ui.prompt(
    'Enter search string: ',
    ui.ButtonSet.OK_CANCEL);
  var button = query.getSelectedButton();
  var queryText = query.getResponseText();
  if (button == ui.Button.OK) {
    search_next(queryText);
  }
}

function search_next(query) {
  // instead of searching column 4 we want to search the grid starting at
  // column 1 and going to sheet.getLastColumn();
  var found = 0;
  for each (var sheet in sheets) {
    if (found == 0) {
      ss.toast('Searching ' + sheet.getName());

      // get how many rows and columns are in the sheet
      var rows = sheet.getLastRow();
      var columns = sheet.getLastColumn();

      //gets values in the sheet and then tests with query
      var range = sheet.getRange(1, 1, rows, columns);
      var values = range.getValues();
      for each (var value in values) {
        if (value == query) {
          ss.setActiveSheet(sheet);
          found = 1;
        }
      }
    }
  }
  if (found == 0) {
    var message = 'Not found.';
    ss.toast(message)
    Browser.msgBox(message);
  }
}

// eventually need to add ability to find next
function search(query) {
  var searchColumn = 4;
  var found = 0;
  for each (var sheet in sheets) {
    if (found == 0) {
      ss.toast('Searching ' + sheet.getName());

      // get how many rows are in the sheet (usually bounded by # of hosts)
      var rows = sheet.getLastRow();
      // catches sheets without any host names
      // otherwise will throw an error in the getRange call later on
      if (rows < 3) {
        rows = 1;
      } else {
        rows = rows - 2;
      }

      // get values in range then test values with query
      var range = sheet.getRange(3, searchColumn, rows, 1);
      var values = range.getValues();
      for each (var value in values) {
        if (value == query) {
          Browser.msgBox(sheet.getName());
          // set active sheet to location of match
          // need to find an efficient way to set the active cell to the match
          ss.setActiveSheet(sheet);
          found = 1;
        }
      }
    }
  }
  if (found == 0) {
    var message = 'Not found.';
    ss.toast(message)
    Browser.msgBox(message);
  }
}
