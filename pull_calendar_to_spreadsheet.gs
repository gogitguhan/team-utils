function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Partner Meetings')
    .addItem('Sync with Calendar', 'getEvents')
    .addToUi();
}

function getEvents() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  let calendar;
  const calendarId = 'guhan.sivaji@neotechnology.com';
  const now = new Date();

  now.setHours(08);

  var startDate = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();

  if (startDate == null && startDate == '') {
    //SpreadsheetApp.getUi().alert("Default start date to pull partner meetings  is " + now);
    startDate = now;
  }

  var cal = CalendarApp.getCalendarById(calendarId);
  console.log("Obtained Calendar=" + cal.getDescription());
  var today = new Date();
  console.log('today=' + today.toDateString());
  var events = cal.getEvents(new Date(startDate), new Date());
  var eventsCount = events.length;

  var nextRow = getFirstEmptyRowByColumnArray();
  var n = 0;
  for (var i = 0; i <= eventsCount - 1; i++) {
    var myEvent = events[i];
    var myDescription = myEvent.getDescription();
    var myStartTime = myEvent.getStartTime();
    var myEndTime = myEvent.getEndTime();
    var title = myEvent.getTitle();

    var sheetMeetings = ss.getSheetByName("Meetings");


    var guests = myEvent.getGuestList();
    var guestCount = guests.length;

    if (guestCount > 0) {
      console.log("Guest:" + guests[0].getEmail() + "Guest Name: " + guests[0].getName());

      if (guests[0].getEmail().indexOf("google.com") > -1) {

        sheetMeetings.getRange(nextRow + n, 1).setValue(myStartTime);
        sheetMeetings.getRange(nextRow + n, 2).setValue("GCP");
        sheetMeetings.getRange(nextRow + n, 3).setValue(title);
        sheetMeetings.getRange(nextRow + n, 4).setValue(guests[0].getName());
        sheetMeetings.getRange(nextRow + n, 5).setValue(guests[0].getEmail());
        if (guestCount > 1 && guests[1].getEmail().indexOf("google.com") > -1) {
          sheetMeetings.getRange(nextRow + n, 6).setValue(guests[1].getName());
          sheetMeetings.getRange(nextRow + n, 7).setValue(guests[1].getEmail());
        }
        n++;
      }
      else if (guests[0].getEmail().indexOf("amazon.com") > -1) {

        sheetMeetings.getRange(nextRow + n, 1).setValue(myStartTime);
        sheetMeetings.getRange(nextRow + n, 2).setValue("AWS");
        sheetMeetings.getRange(nextRow + n, 3).setValue(title);
        sheetMeetings.getRange(nextRow + n, 4).setValue(guests[0].getName());
        sheetMeetings.getRange(nextRow + n, 5).setValue(guests[0].getEmail());
        if (guestCount > 1 && guests[1].getEmail().indexOf("amazon.com") > -1) {
          sheetMeetings.getRange(nextRow + n, 6).setValue(guests[1].getName());
          sheetMeetings.getRange(nextRow + n, 7).setValue(guests[1].getEmail());
        }
        n++;
      }
      else if (guests[0].getEmail().indexOf("microsoft.com") > -1) {

        sheetMeetings.getRange(nextRow + n, 1).setValue(myStartTime);
        sheetMeetings.getRange(nextRow + n, 2).setValue("Azure");
        sheetMeetings.getRange(nextRow + n, 3).setValue(title);
        sheetMeetings.getRange(nextRow + n, 4).setValue(guests[0].getName());
        sheetMeetings.getRange(nextRow + n, 5).setValue(guests[0].getEmail());
        if (guestCount > 1 && guests[1].getEmail().indexOf("microsoft.com") > -1) {
          sheetMeetings.getRange(nextRow + n, 6).setValue(guests[1].getName());
          sheetMeetings.getRange(nextRow + n, 7).setValue(guests[1].getEmail());
        }
        n++;
      }
    }

  }
  function getFirstEmptyRowByColumnArray() {
    var spr = SpreadsheetApp.getActiveSpreadsheet();
    var column = spr.getRange('A:A');
    var values = column.getValues(); // get all data in one call
    var ct = 0;
    while (values[ct] && values[ct][0] != "") {
      ct++;
    }
    return (ct + 1);
  }

}
