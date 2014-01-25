/*
 * gDoc Sub System
 *
 */


/*** Global Params ***/
var numHeaderRows = 1;
var numHeaderCols = 2;
var numSubColumnPairs = 5;
var docUrl = '';
var toEmail = 'group@listserv.com';


/*** Optional Params ***/
var toastTimeout = 3;


/*** Main Functions ***/

// adds a 'Subs' menu
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({
    name: "Send Email",
    functionName: "sendEmail"
  });
  menuEntries.push({
    name: "Update Colors",
    functionName: "updateColors"
  });
  menuEntries.push(null);
  menuEntries.push({
    name: "About",
    functionName: "about"
  });
  ss.addMenu("Subs", menuEntries);
}

// updates the colors of edited rows
function onEdit(e) {
  if (e.range.getGridId() == 0) {
    for (var row = e.range.getRow(); row <= e.range.getLastRow(); row++) {
      colorRow(row);
    }
  }
}

function sendEmail() {
  SpreadsheetApp.getActiveSpreadsheet().toast(toEmail, 'Sending Email', toastTimeout);
}

// colors all rows
function updateColors() {
  var startRow = 1 + numHeaderRows;
  var endRow = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getMaxRows();
  SpreadsheetApp.getActiveSpreadsheet().toast('updating rows ' + startRow + ' - ' + endRow, 'Started', toastTimeout);

  for (var row = startRow; row <= endRow; row++) {
    colorRow(row);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('updating rows ' + startRow + ' - ' + endRow, 'Finished', toastTimeout);
}

// colors a row based on the date in first column
function colorRow(row) {
  if (row <= numHeaderRows) {
    return;
  }

  firstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  date = firstSheet.getRange(row, 1).getValue()
  rowRange = firstSheet.getRange(row, 1 + numHeaderCols, 1, numSubColumnPairs * 2)
  var rowData = rowRange.getValues()[0];
  var rowColors = [];

  if (!(date instanceof Date)) {
    rowRange.setBackground('white');
    return;
  }

  if (dateTodayDiff(date) < -1) {
    rowRange.setBackground('grey');
    return;
  }

  for (var col = 0; col < rowData.length; col += 2) {
    var orig = rowData[col];
    var sub = rowData[col + 1];

    if (orig != '' && sub == '') {
      rowColors.push('pink');
    } else if (orig != '' && sub != '') {
      rowColors.push('lightgreen');
    } else {
      rowColors.push('white');
    }
    rowColors.push('#F3F3F3');
  }
  rowRange.setBackgrounds([rowColors]);
}

function about() {
  SpreadsheetApp.getActiveSpreadsheet().toast('http://ultimateanu.github.io/gdoc-sub-system', 'gDoc Sub System');
}


/*** Helper Functions ***/

function dateTodayDiff(d) {
  var utcDate = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.floor((utcDate - Date()) / (1000 * 60 * 60 * 24));
}