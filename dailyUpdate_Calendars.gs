function getTodayEventAllCals(){

  var calendarId = "";
  var now = new Date();
  var events = null;
  var roomStr = "";
  var diaryStr = "";
  var bookedRooms = 0;
  
  //var maxTime = new Date(now.getDate() + 1);
  
  var maxTime = new Date();
  maxTime.setDate(now.getDate() + 1);
  maxTime.setHours(0);
  maxTime.setMinutes(0);
  maxTime.setSeconds(0);
  
  for (var i = 0; i < calendarIds.length; i++){
  
    calendarId = calendarIds[i];
    roomStr = "Room " + (i+1);
    
    events = Calendar.Events.list(calendarId, {
      timeMin: now.toISOString(),
      timeMax: maxTime.toISOString(),
      singleEvents: true,
      orderBy: 'startTime',
      maxResults: 1
    });
    
    if (events.items && events.items.length > 0) {
      
      var event = events.items[0];
      diaryStr += "<b>" + roomStr + '</b> - ' + event.summary + "<br>";
      
      // Testing detect blocks/bookings
      //Logger.log('Original ' + roomStr + ' ' + event.summary);
      var arrSummary = (event.summary).split('-');
      var strSummary = arrSummary.join('');
      var intSummary = parseInt(strSummary);
      
      if (strSummary.length == 9){
      //if (intSummary > 100000){
        // Detect as booking reference
        //Logger.log('Booking ' + ' ' + roomStr + ' ' + event.summary);
        bookedRooms++;
      }else{
        // Detect block
        //Logger.log('Block ' + roomStr + ' ' + event.summary);
      }
        
    } else {
      //Logger.log(roomStr + ' - No bookings today');
      diaryStr += "<b>" + roomStr + '</b> - No bookings today<br>';
    }
    
  }
  
  // Optional - flag if more than 2 calendars have events for today
  if (bookedRooms > 2){
    diaryStr += "<br><b style='color: red;'>You have " + bookedRooms + " rooms booked tonight - should you turn on the hot water boost this evening?</b>";
  }
  
  return diaryStr;

}
