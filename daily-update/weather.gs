function getWeatherData() {

  // Uncomment this line if using function on its own!
  // setupDates();
  
  var timestamps = [];
  var prettyTimes = [];
  
  // Timestamps for 6pm and 9pm tonight
  timestamps.push(thisDate + "T18Z");
  prettyTimes.push(thisPrettyDate + " 6pm");
  timestamps.push(thisDate + "T21Z");
  prettyTimes.push(thisPrettyDate + " 9pm");
  
  // Timestamps for 12am, 3am, 6am tomorrow
  timestamps.push(tmwDate + "T00Z");
  prettyTimes.push(tmwPrettyDate + " midnight");
  timestamps.push(tmwDate + "T03Z");
  prettyTimes.push(tmwPrettyDate + " 3am");
  timestamps.push(tmwDate + "T06Z");
  prettyTimes.push(tmwPrettyDate + " 6am");
  
  body += "<table style='text-align: center; border: 1px solid black; border-collapse: collapse;'><tr>";
  
  for (var i = 0; i < prettyTimes.length; i++){
    body += "<th style='border: 1px solid black;'>" + prettyTimes[i] + "</th>";
  }
  
  body += "</tr><tr>";
  var humidFlag = false;
  var maxHumid = 0;
  
  for (var i = 0; i < timestamps.length; i++){
  
    var humidityVal = getHumidityAtTime(timestamps[i],false);
    var humidCol = 'green';
    
    if (humidityVal >= 90){
      humidCol = 'red';
      humidFlag = true;
      if (humidityVal > maxHumid){ maxHumid = humidityVal }
      Logger.log('OVER 90 ' + humidityVal + ' ' + timestamps[i]);
    }else{
      Logger.log('UNDER 90 ' + humidityVal + ' ' + timestamps[i]);
    }
    
    body += "<td style='border: 1px solid black; color: " + humidCol + ";'>" + humidityVal + "%</td>";
    
  }
  
  body += "</tr></table>";
  
  if (humidFlag){
    body += "<br><b style='color: red;'>Warning, the humidity is expected to reach " + maxHumid + "% tonight - can you take any corrective action?</b>";
  }
  
}


function convertMonth(month){
  var thisMonth = ( month + 1 < 10) ? "0" + ( month + 1) : month + 1;
  return thisMonth;
}

function convertDay(day){
  var thisDay = ( day < 10) ? "0" + day : day;
  return thisDay;
}

function getHumidityAtTime(thisTime,booText){

  //Logger.log("-------------------------------------------");
  //Logger.log("Getting data at " + thisTime);
  var urlPre = "http://datapoint.metoffice.gov.uk/public/data/val/wxfcs/all/json/322953?res=3hourly&time=";
  // Update this to your own API Key from the Met Office
  var urlSuf = "&key=YOUR-API-KEY";
  
  var getUrl = urlPre + thisTime + urlSuf;

  var resp = UrlFetchApp.fetch(getUrl);
  var respStr = resp.getContentText().toString();
  var respJson = JSON.parse(respStr);
  
  var weatherData = respJson["SiteRep"]["DV"]["Location"]; 
  var thisTimeData = weatherData["Period"]["Rep"];  
  var humidityValue = thisTimeData["H"];
  
  if (booText){
    if (humidityValue >= 90){
      // Show warning!
      //Logger.log("WARNING! The forecast humidity is " + humidityValue + "%");
      return "WARNING! The forecast humidity is " + humidityValue + "%";
    }else{
      //Logger.log("No worries, the forecast humidity is " + humidityValue + "%");
      return "No worries, the forecast humidity is " + humidityValue + "%";
    }
  }else{
    return humidityValue;// + "%";
  }
  
  //SpreadsheetApp.getActive().insertRowsBefore(1, 5);
  //SpreadsheetApp.setActiveRange("A1:B5");
  
  
}
