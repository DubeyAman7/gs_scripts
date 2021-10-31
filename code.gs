function doGet(e) {
  var params = JSON.stringify(e);
  return HtmlService.createHtmlOutput(params);
}



function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Use Case Name')
      .addItem('Select your Mentor', 'SelectMentor')
      .addItem('Book session with your mentor', 'BookSession')
      .addToUi();
}






// The mentee should be able to see the slots available with his mentor only

function SelectMentor(a){

  var today = Utilities.formatDate(new Date(), "GMT + 0530", "dd/MM/yyyy");


  



// Arriving at the name of the mentor from Sheet 'Mentee Mentor Mapping'
  var mentee_name= String(Session.getActiveUser());
  //console.log(mentee_name)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Mentee Mentor Mapping");


  //var row = Number(mentee_name);
  var range = sheet.getRange(2,1,10,2);
  var values = range.getValues();

  for (var i in values) {
    var row = values[i];
    var mentee = row[0];
    if (mentee === mentee_name) {
      var mentor_name = row[1];

    }
  }

  
    

  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Available Slots');
  var range = ss.getRange(10,1,10,10);
  var filter = range.getFilter() || range.createFilter()
  var mentor_1 = SpreadsheetApp.newFilterCriteria().whenTextContains(mentor_name); 
  var booked = SpreadsheetApp.newFilterCriteria()
  .whenNumberEqualTo(0)
  var date_1 = SpreadsheetApp.newFilterCriteria()
  .whenNumberEqualTo(0) //this will filter on amount greater than 50000
  .build();
  filter.setColumnFilterCriteria(1, mentor_1);
  filter.setColumnFilterCriteria(6, booked);
  filter.setColumnFilterCriteria(7,date_1)

}



function BookSession(e){
  //var range = e.range
  //Logger.log(range.getA1Notation())
  // The code below sets the value of name to the name input by the user, or 'cancel'.
  var name = Browser.inputBox('Slot Selection', 'Row Number of slot to be booked', Browser.Buttons.OK_CANCEL);
  console.log(name)
  handleCalenderInvites(name)

  
}

function handleCalenderInvites(name){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Available Slots');
  var row = Number(name);
  var range = sheet.getRange(row,1,1,10);
  var values = range.getValues();
  console.log(values);
  meeting_detail = values[0];
  if(meeting_detail[5]==0 && meeting_detail[6]==0){
  sendCalenderinvite(meeting_detail);
  sheet.getRange(row,6).setValue(1) 
  sheet.getRange(row,8).setValue(String(Session.getActiveUser()))
  }
  else{
  Browser.msgBox('Sorry the session you want to book is already booked');
  }
}

function joinDateAndTime_(date, time) {
  date = new Date(date);
  console.log(time)
  date.setHours(time.getHours());
  date.setMinutes(time.getMinutes());
  return date;
}

function sendCalenderinvite(details){
  var cal = CalendarApp.createCalendar('Mentoring Calender');
  var title = 'Mentorship session with ' + details[0];
  var start_time = joinDateAndTime_(details[2], details[3]);
  var end_time = joinDateAndTime_(details[2], details[4]);
  var options = {description:"Your session with the mentor has been booked " , location: "Virtual", sendInvites: true, guests: details[1]+','+'firstsalary.in@gmail.com' +','+ String(Session.getActiveUser())};
  var event = cal.createEvent(title, start_time, end_time, options)
        .setGuestsCanSeeGuests(true);
  console.log(event.getId)
}