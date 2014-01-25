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
    name: "Reset",
    functionName: "reset"
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

function updateColors() {
  dataRange = getDataRange()
  SpreadsheetApp.getActiveSpreadsheet().toast(dataRange.getHeight(), 'h');
  dataRange.setBackground("blue");

  colorRow(3);
  colorRow(5);
  colorRow(7);

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

function reset() {
  SpreadsheetApp.getActiveSpreadsheet().toast('', 'Resetting', toastTimeout);
}

function about() {
  SpreadsheetApp.getActiveSpreadsheet().toast('http://ultimateanu.github.io/gdoc-sub-system', 'gDoc Sub System');
}

function getDataRange() {
  firstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  return firstSheet.getRange(1 + numHeaderRows, 1 + numHeaderCols, firstSheet.getMaxRows() - numHeaderRows, numSubColumnPairs * 2)
}


/*** Helper Functions ***/

function DateDiffInDays(a, b) {
  var _MS_PER_DAY = 1000 * 60 * 60 * 24;
  // Discard the time and time-zone information.
  var utc1 = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  var utc2 = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.floor((utc2 - utc1) / _MS_PER_DAY);
}

function dateTodayDiff(d) {
  return DateDiffInDays(new Date(), d);
}