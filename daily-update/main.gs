/*
  -- Calendar variables
*/
// An array of the Google calendars to be checked each day
var calendarIds = ["3x6m9l3@import.calendar.google.com", "6n0t4er@import.calendar.google.com" ];

/*
  -- Weather variables
*/
var metOfficeURLprefix = "http://datapoint.metoffice.gov.uk/public/data/val/";
// Update this to your location ID, found when using the Met Office website
var metOfficeLocation = "322953";
var metOfficeURLsuffix = "/all/json/capabilities?res=3hourly&key=";
// Update this to your own API key from the Met Office website
var metOfficeAPIkey = "YOUR-MET-OFFICE-API-KEY";


/* 
  -- Email Variables
*/
// Update this to your email(s) which will receive the daily update
var recipients = "YOUREMAIL@example.com; anotheremail@example.com;";
var subject = "Daily Update Email - ";
var body = "";

// Moving these out so they can be used in other functions
var thisPrettyDate = "";
var tmwPrettyDate = "";
var thisDate = "";
var tmwDate = "";
  
function setupDates(){

  var today = new Date();
  var thisMonth = convertMonth(today.getMonth());
  var thisDay = convertDay(today.getDate());
  thisDate = today.getFullYear() + "-" + thisMonth + "-" + thisDay;
  
  var thisShortYear = String(today.getFullYear()).substr(2,4);
  // Used in multiple functions
  thisPrettyDate = thisDay + '/' + thisMonth + '/' + thisShortYear;
  
  var tmw = new Date();
  tmw.setDate(today.getDate()+1);
  
  var tmwMonth = convertMonth(tmw.getMonth());
  var tmwDay = convertDay(tmw.getDate());
  tmwDate = tmw.getFullYear() + "-" + tmwMonth + "-" + tmwDay;
  
  var tmwShortYear = String(tmw.getFullYear()).substr(2,4);
  tmwPrettyDate = tmwDay + '/' + tmwMonth + '/' + tmwShortYear;

}




// This is the function called at each trigger, calling the other functions required
function generateEmail(){

  setupDates();
  
  body = "Your daily update email.";
  
  body += checkWasteCollection();
  
  body += "<h2>Humidity Forecast</h2>";
  
  getWeatherData();
  
  body += "<br><em>For full weather information, visit the <a href='https://www.metoffice.gov.uk/weather/forecast/gctqz7t1r'>Met Office website</a></em><br>";
  body += "<h2>Diary for Today</h2>";
  
  body += getTodayEventAllCals();
  
  body += "<br>END OF YOUR BODY EMAIL";
  
  GmailApp.sendEmail(recipients, subject + thisPrettyDate , "", {htmlBody:body})

}
