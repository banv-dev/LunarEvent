function onSheetHomepage() {
    return createCard();
}


function createCard() {
  var cardHeader = CardService.newCardHeader()
    .setTitle("Tạo lịch âm")
    .setImageStyle(CardService.ImageStyle.CIRCLE)
    .setImageUrl("https://play-lh.googleusercontent.com/fjnNV-3KBihOVr50aYvKVGRiqeD2gRxH5S2_CnaA9LiZ8TTxY0R7NnFmmL8q5YrBmv0=s128-rw");

  var listCalendars = getListCalendars();
  var selectCalendar = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Chọn lịch âm")
    .setFieldName("calendar_list");

  for(i = 0; i < listCalendars.length; i++) {
    selectCalendar.addItem(listCalendars[i].summary, listCalendars[i].id, false);
  }

  var actionUpdate = CardService.newAction()
      .setFunctionName('onUpdateEvent');

  var buttonUpdate = CardService.newTextButton()
      .setText('Cập nhật')
      .setOnClickAction(actionUpdate)
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  var actionDelete = CardService.newAction()
      .setFunctionName('onDeleteEvent');

  var buttonDelete = CardService.newTextButton()
      .setText('Xoá')
      .setOnClickAction(actionDelete)
      .setTextButtonStyle(CardService.TextButtonStyle.TEXT);

var buttonSet = CardService.newButtonSet()
    .addButton(buttonUpdate)
    .addButton(buttonDelete);

  var section = CardService.newCardSection()
    .addWidget(selectCalendar)
    .addWidget(buttonSet);

   var card = CardService.newCardBuilder().setHeader(cardHeader).addSection(section);

  return card.build();
}

function onUpdateEvent(e) {
  var calendarId = e.commonEventObject.formInputs.calendar_list[""].stringInputs.value[0];
  notificationText   = updateLunarEvent(calendarId);
  
  return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
            .setText(notificationText)
            .setType(CardService.NotificationType.INFO))
        .build();      // Don't forget to build the response!

}

function updateLunarEvent(calendarId) {
  var calendar = CalendarApp.getCalendarById(calendarId);
  var message = "";
  if (calendar == null) {
    message = "Không tìm thấy lịch với id = " + calendarId;
    Logger.log(message);
    return message; 
  }

  Logger.log("User select calendar with id %s, name %s", calendarId, calendar.getName());
  year = new Date().getFullYear();

  var count = 0;
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == null || data[i][0]=="") {
      continue;
    }
    dateMonth = data[i][0].toString().split("/");
    var date = parseInt(dateMonth[0]);
    var month = parseInt(dateMonth[1]);
    var solarDate = getSolarDate(date, month, year);
    //update excel column Solar date
    dataRange.getCell(i + 1, 4).setValue(array2DateString(solarDate));

    var eventDate = array2Date(solarDate);
    //delete old event
    deleteLunarEventInCalendar(calendar, eventDate, eventDate);

    //create calendar event
    createLunarEvent(calendar, data[i][1], data[i][2], eventDate);

    count ++;
    if (count % 10 == 0) {
      Utilities.sleep(2000); 
      //to avoid error: invoked too many times in a short time
      //https://developers.google.com/apps-script/guides/services/quotas
    }

  }

  message = "Đã tạo " + count + " sự kiện trên lịch âm của bạn";

  return message;

}

function createLunarEvent(calendar,title, description, date) {
  var event = calendar.createAllDayEvent(title, date, 
  {
        description: description,
  });
  event.addEmailReminder(7 * 24 * 60); //1 week before
  event.addEmailReminder(14 * 24 * 60); //2 weeks before
  event.setTag("project", "lunar");
  event.setTag("version", "v1");
}

function deleteLunarEventInCalendar(calendar, start, end) {
  var count = 0;
  var events = calendar.getEvents(start, end);
  for (var i = 0; i < events.length; i++) {
    var e = events[i];
    if (e.getTag("project") == "lunar") {
      count ++;
      e.deleteEvent();

    }
  }
  return count;
}

function onDeleteEvent(e) {
  var calendarId = e.commonEventObject.formInputs.calendar_list[""].stringInputs.value[0];
  notificationText = deleteLunarEvent(calendarId);
  return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
            .setText(notificationText)
            .setType(CardService.NotificationType.INFO))
        .build();      
}

function deleteLunarEvent(calendarId) {
  var calendar = CalendarApp.getCalendarById(calendarId);
    // var calendar = CalendarApp.getCalendarById("4pkra3i2u3hpeqoekjfa2fppu8@group.calendar.google.com");
  var message = null;
  if (calendar == null) {
    message = "Không tìm thấy lịch với id = " + calendarId;
    Logger.log(message);
    return message; 
  }

  var start = new Date();
  start.setFullYear(start.getFullYear(), 1, 1);
  var end = new Date()
  end.setFullYear(start.getFullYear() + 1, 11, 01);
  var count = deleteLunarEventInCalendar(calendar, start, end);
  message =  "Đã xoá " + count + " sự kiện trên lịch " + calendar.getName();
  Logger.log(message);
  return message;

}

function storeEvent(calendarId, eventId) {
  var lunarEventString = PropertiesService.getUserProperties().getProperty("LUNAR_EVENT");
  var lunarEvents = JSON.parse(lunarEventString);
  Logger.log(lunarEvents);
  
  if (lunarEvents == null) {
    lunarEvents = [];
  }

  lunarEvents.push({
    calendar: calendarId,
    event: eventId
  });

  PropertiesService.getUserProperties().setProperty("LUNAR_EVENT", JSON.stringify(lunarEvents));

}

function testProps() {
  PropertiesService.getUserProperties().deleteAllProperties();
  var scriptProperties = PropertiesService.getUserProperties();
  var data = scriptProperties.getProperties();
  for (var key in data) {
    Logger.log('Key: %s, Value: %s', key, data[key]);
  }
}


