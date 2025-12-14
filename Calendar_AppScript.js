function gettoEvents() {
var ss=SpreadsheetApp.getActiveSpreadsheet()
var as=ss.getSheetByName("Data")
var cal = CalendarApp.getCalendarById("jeevanandam702@gmail.com");
var start=as.getRange("F2").getValue()
var end=as.getRange("G2").getValue()
Logger.log(start)
Logger.log(end)
var events = cal.getEvents(new Date(start), new Date(end))
// Logger.log(events)
var lr = as.getLastRow()
as.getRange(2,1,lr-1,4).clearContent()
as.getRange(2,1,lr-1,4).clearFormat()


for(let i=0; i<events.length; i++){
  var title = events[i].getTitle();
  var sd = events[i].getStartTime();
  var ed = events[i].getEndTime();
  var des = events[i].getDescription();
  as.getRange(i+2,1).setValue(title)
  as.getRange(i+2,2).setValue(sd)
  as.getRange(i+2,2).setNumberFormat("dd/mm/yyyy h:mm AM/PM")
  as.getRange(i+2,3).setValue(ed)
  as.getRange(i+2,3).setNumberFormat("dd/mm/yyyy h:mm AM/PM")
  as.getRange(i+2,4).setValue(des)
}
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Approve")
    .addItem("Approve", "getEvents")
    .addToUi();
}

//with Status Option
/**
 function getEvents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Data");
  var cal = CalendarApp.getCalendarById("jeevanandam702@gmail.com");

  var start = sheet.getRange("F2").getValue();
  var end = sheet.getRange("G2").getValue();

  // Safety check
  if (!(start instanceof Date) || !(end instanceof Date)) {
    SpreadsheetApp.getUi().alert("Please select valid Start and End dates");
    return;
  }

  var events = cal.getEvents(start, end);

  // Clear old data safely
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 5).clearContent().clearFormat();
    sheet.getRange(2, 5, lastRow - 1).clearDataValidations();
  }

  // Dropdown rule
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pending", "Completed", "Cancelled"], true)
    .setAllowInvalid(false)
    .build();

  // Write events + add dropdown only when event exists
  events.forEach((event, i) => {
    var row = i + 2;

    sheet.getRange(row, 1).setValue(event.getTitle());
    sheet.getRange(row, 2)
      .setValue(event.getStartTime())
      .setNumberFormat("dd/MM/yyyy h:mm AM/PM");
    sheet.getRange(row, 3)
      .setValue(event.getEndTime())
      .setNumberFormat("dd/MM/yyyy h:mm AM/PM");
    sheet.getRange(row, 4).setValue(event.getDescription());

    // ✅ Add dropdown to Status column
    sheet.getRange(row, 5).setDataValidation(statusRule);
  });
}
function onEdit(e) {
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Only Column A (Events), row 2+
  if (col !== 1 || row < 2) return;

  var statusCell = sheet.getRange(row, 5);
  var value = e.range.getValue();

  if (value !== "") {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Pending", "Completed", "Cancelled"], true)
      .setAllowInvalid(false)
      .build();

    statusCell.setDataValidation(rule);
  } else {
    statusCell.clearDataValidations();
    statusCell.clearContent();
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Approve")
    .addItem("Approve", "getEvents")
    .addToUi();
}
 */