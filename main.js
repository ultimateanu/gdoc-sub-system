/*
 * gDoc Sub System
 *
 */


/*** Global Params ***/
var numHeaderRows = 1;
var numHeaderCols = 2;
var numSubColumnPairs = 5;
var toEmail = 'group@listserv.com';
var emailSubject = 'Group Name';


/*** Optional Params ***/
var toastTimeout = 3;
var headerBgColor = '#465C8E'; //darkBlue
var borderColor = '#A2A3A4'; //lightBlue
var altColor = '#EAEBED'; //veryLightBlue


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
  for (var row = e.range.getRow(); row <= e.range.getLastRow(); row++) {
    colorRow(row);
  }
}

// colors all rows
function updateColors() {
  var startRow = 1 + numHeaderRows;
  var endRow = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getMaxRows();
  for (var row = startRow; row <= endRow; row++) {
    colorRow(row);
  }
}

// colors a row based on the date in first column
function colorRow(row) {
  if (row <= numHeaderRows) {
    return;
  }

  var firstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var date = firstSheet.getRange(row, 1).getValue();
  var rowRange = firstSheet.getRange(row, 1 + numHeaderCols, 1, numSubColumnPairs * 2);

  if (!(date instanceof Date)) {
    rowRange.setBackground('white');
    return;
  }

  if (dateTodayDiff(date) < 0) {
    rowRange.setBackground('grey');
    return;
  }

  var rowData = rowRange.getValues()[0];
  var rowColors = [];

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

// sends an email digest with sub requests
function sendEmail() {
  var firstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var names = [],
    dates = [],
    reasons = [];

  for (var row = 1 + numHeaderRows; row <= firstSheet.getMaxRows(); row++) {
    var date = firstSheet.getRange(row, 1).getValue();

    if ((date instanceof Date) && dateTodayDiff(date) >= 0) {
      var rowRange = firstSheet.getRange(row, 1 + numHeaderCols, 1, numSubColumnPairs * 2);
      var rowData = rowRange.getValues()[0];
      var rowNotes = rowRange.getNotes()[0];
      for (var col = 0; col < rowData.length; col += 2) {
        var orig = rowData[col];
        var sub = rowData[col + 1];
        if (orig != '' && sub == '') {
          names.push(orig);
          dates.push(date);
          var r = rowNotes[col];
          if (r == '') {
            r = '{reason not specified. please add reason as a NOTE in gdoc}';
          }
          reasons.push(r);
        }
      }
    }
  }


  if (names.length > 0) {
    var htmlMsg = createHTMLTable(names, dates, reasons);
    htmlMsg += '<br><br><br><br> *This message was automatically generated at ' + formatToday('h:mm a (M/d/yy)') + ' by ' + appName + ' [' + appUrl + '].';
    var subject = emailSubject + ' ' + formatToday('M/d') + ' Digest [' + names.length + ' needed]';
    MailApp.sendEmail(toEmail, subject, "Your email client doesn't support HTML :(", {
      htmlBody: htmlMsg
    });
    SpreadsheetApp.getActiveSpreadsheet().toast(names.length + ' people need subs', 'Emailed ' + toEmail, toastTimeout);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('No one needs a sub', 'Email not sent', toastTimeout);
  }
}

// returns a formatted html table
function createHTMLTable(nameList, dateList, reasonList) {
  var HTMLtable = wrapHTML('Sub Requests', 'h2', 'style="color:' + headerBgColor + '"');
  HTMLtable += '<table  cellpadding="5" style="border:1px solid ' + headerBgColor + ';border-collapse:collapse;width:100%">';
  var hDate = wrapHTML('Date', 'th', 'width="15%" align="center"');
  var hName = wrapHTML('Name', 'th', 'width="15%" align="center"');
  var hMessage = wrapHTML('Message', 'th', 'width="70%" align="center"');
  HTMLtable += wrapHTML(hDate + hName + hMessage, 'tr', 'style="background-color:' + headerBgColor + ';color:white"');

  for (var i = 0; i < nameList.length; i++) {
    var bgColor = '';
    if (i % 2 != 0) bgColor = 'bgcolor="' + altColor + '" ';

    var cellStyle = 'style="border:1px solid ' + borderColor + '"';
    hDate = wrapHTML(formatDate(dateList[i], 'M/d (EEE)'), 'td', cellStyle);
    hName = wrapHTML(nameList[i], 'td', cellStyle);
    hMessage = wrapHTML(reasonList[i], 'td', cellStyle);
    HTMLtable += wrapHTML(hDate + hName + hMessage, 'tr', bgColor + cellStyle);
  }

  HTMLtable += "</table>";
  HTMLtable += '<br>Sign up: ';
  HTMLtable += wrapHTML('LINK', 'a', 'href="' + SpreadsheetApp.getActive().getUrl() + '"');
  return HTMLtable;
}

// displays app info
function about() {
  SpreadsheetApp.getActiveSpreadsheet().toast(appUrl, appName);
}


/*** Helper Functions ***/

function dateTodayDiff(d) {
  var today = new Date();
  var dateUtc = Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  var todayUtc = Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate());
  return Math.floor((dateUtc - todayUtc) / (1000 * 60 * 60 * 24));
}

function wrapHTML(inside, html, options) {
  return '<' + html + ' ' + options + '>' + inside + '</' + html + '>';
}

function formatDate(d, f) {
  return Utilities.formatDate(d, Session.getTimeZone(), f);
}

function formatToday(f) {
  return formatDate(new Date(), f);
}

/*** App Params ***/
var appName = 'gDoc Sub System';
var appUrl = 'http://ultimateanu.github.io/gdoc-sub-system';