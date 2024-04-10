var URL_BOOK = 'https://docs.google.com/spreadsheets/d/your_sheet_id/';
var SHEETNAME = 'fromSlack';
var CALENDAR_ID = 'your_calender_id@group.calendar.google.com';

function onPermit() {
  //var sheet = getSheet(URL_BOOK, SHEETNAME);
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveRange();
Logger.log(cell.getColumn() +" "+cell.getRow() + " "+cell.getValue());
  if(cell.getColumn() != 1) { // Column A:Permit
    return;
  }
  if(cell.getValue() != "OK") { // OK or Reject
    return;
  }
  
  var thisRow = cell.getRow();
  var json = convertSheet2Json(sheet, thisRow);

  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  for(var i=0; i<json.length; i++) {
    var event = json[i];
    var title = event.Name;
    var starttime = event.Date + " " + event.StartTime + ".00";
    var endtime = event.Date + " " + event.EndTime + ".00";
    var start = new Date(starttime);
    var end = new Date(endtime);
    var purpose = event.ProjectName + " / " + event.Purpose;
    var options = {description: purpose};
    calendar.createEvent(title, start, end, options);
  }
}

function getSheet(bookUrl, sheetName) {
  var book = SpreadsheetApp.openByUrl(bookUrl);
  return book.getSheetByName(sheetName);
}

// https://gist.github.com/daichan4649/8877801#file-convertsheet2json-gs
function convertSheet2Json(sheet, row) {
  // first line(title)
  var colStartIndex = 2;
  var rowNum = 1;
  var firstRange = sheet.getRange(1, colStartIndex, rowNum, sheet.getLastColumn());
  var firstRowValues = firstRange.getValues();
  var titleColumns = firstRowValues[0];

  // data
  var rowValues = [];
  var range = sheet.getRange(row, colStartIndex, rowNum, sheet.getLastColumn());
  var values = range.getValues();
  var rowValues = [];
  rowValues.push(values[0]);

  // create json
  var jsonArray = [];
  for(var i=0; i<rowValues.length; i++) {
    var line = rowValues[i];
    var json = new Object();
    for(var j=0; j<titleColumns.length; j++) {
      json[titleColumns[j]] = line[j];
    }
    jsonArray.push(json);
  }
  return jsonArray;
}
