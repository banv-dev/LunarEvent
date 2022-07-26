var DATETIME_SEPARATOR = "/";

function getListCalendars() {
  var listCalendar = [];
  var calendars;
  var pageToken;
  do {
    calendars = Calendar.CalendarList.list({
      maxResults: 100,
      pageToken: pageToken
    });
    if (calendars.items && calendars.items.length > 0) {
      for (var i = 0; i < calendars.items.length; i++) {
        var calendar = calendars.items[i];
        Logger.log('%s (ID: %s)', calendar.summary, calendar.id);
        listCalendar.push(calendar);
      }
    } else {
      Logger.log('No calendars found.');
    }
    pageToken = calendars.nextPageToken;
  } while (pageToken);

  return listCalendar;
}


function dateString(day, month, year) {
  return day + DATETIME_SEPARATOR + month + DATETIME_SEPARATOR + year;
}

function array2DateString(dateArray) {
  return dateArray[0] + DATETIME_SEPARATOR + dateArray[1] + DATETIME_SEPARATOR + dateArray[2];
}

function array2Date(dateArray) {
  var date =  new Date();
  //month range: 0->11
  date.setFullYear(dateArray[2], dateArray[1] - 1, dateArray[0]);
  return date;
}

