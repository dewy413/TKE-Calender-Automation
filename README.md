# TKE-Calender-Automation


This was a calendar automation program I created along with a Google Calendar for my fraternity. The idea is that if a member had an event planned they would submit a Google Form, and my script would create an event on the calendar at the click of a button. This is still being used in my chapter, and it allowed me to learn Google Apps Script.

```
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    var mainMenu = ui.createMenu("TKE Xi Tau Scripts");
    mainMenu.addItem("Update Calender", "createEvents");
    mainMenu.addToUi();
};

function createEvents() {
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var eventCal = CalendarApp.getCalendarById("c_8id7jc5cbi26hf914glpadmcu8@group.calendar.google.com");
  var items = spreadsheet.getDataRange().getValues();

  for (x = 1; x < items.length; x++) {
    
    var details = items[x];
    var name = details[1];
    var eventdetails = details[2];
    var starttime = new Date(details[3]);
    var endtime = new Date(details[4]);
    var place = details[5];
    var host = details[6]; 
    var nameOfHost = details[7];
    var notiAmount = details[8].split(", ");

    var event = eventCal.createEvent(name, starttime, endtime, {location: place, description: "Person in Charge: " + nameOfHost + "\nDetails: " + eventdetails});

    switch(host) {
      case "Prytanis/Weekly Meeting/Officer Meeting":
        event.setColor(7);
        break;
      case "Epiprytanis/Chair Meeting/Chair":
        event.setColor(6);
        break;
      case "Crysophylos":
        event.setColor(5);
        break;
      case "Histor":
        event.setColor(10);
        break;
      case "Hypophetes":
        event.setColor(11)
        break;
      case "Hegemon":
        event.setColor(8)
        break;
      case "Pylortes": 
        event.setColor(4);
        break;
    }

    for(j = 0; j < notiAmount.length; j++) {
      switch(notiAmount[j]) {
        case "1hr before":
          event.addPopupReminder(60);
          break;
        
        case "3hr before":
          event.addPopupReminder(180);
          break;

        case "1d before":
          event.addPopupReminder(1440);
          break;

        case "3d before":
          event.addPopupReminder(4320);
          break;

        case "1w before":
          event.addPopupReminder(10080);
          break;

      }
    }
  
    spreadsheet.deleteRow(2); //Deletes item from Excel sheet so it doesn't stay there when a new execution of this code is ran
  }

}

```
