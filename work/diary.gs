// Array holding the names of the sheets to exclude from the execution
var arrExcludedSheets = ['Summary'];
var strMyEmail = 'YOUR_EMAIL_HERE';
// The format of my timesheet, data starting at row 3 - row 30/33 depending on month length 
var arrHeadings = [
  //NAME       COL# Alpha Array
  'Timesheet',  //1   A   0
  'Date',       //2   B   1
  'Start',      //3   C   2
  'BreakS',     //4   D   3
  'BreakE',     //5   E   4
  'End',        //6   F   5
  'Hrs',        //7   G   6
  '£apx',       //8   H   7
  'Notes'       //9   I   8
];
var arrDayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

/**
 * @param string monthSheetName The name of a specific month's sheet to limit output
 */
function generateDiary(monthSheetName = null) {

  var ss = SpreadsheetApp.getActive();
  var allSheets = ss.getSheets();
  var sheetNameArray = [];

  // Sort sheets by name
  for (var i = 0; i < allSheets.length; i++) {

    // If the sheet name is found in the excluded array, continue 
    if(arrExcludedSheets.indexOf(allSheets[i].getName()) !== -1) continue;

    // If a month has been specified
    if (monthSheetName !== null){

      // If the sheet name does not match, continue
      if (monthSheetName !== allSheets[i].getName()) continue;
    }

    sheetNameArray.push(allSheets[i].getName());
  }
  
  sheetNameArray.sort();
  Logger.log(sheetNameArray);
  Logger.log('Sheet count: ' + sheetNameArray.length);

  var strDiary = '<h1>Diary Output</h1>';

  for (var i = 0; i < sheetNameArray.length; i++) {

    thisSheetName = sheetNameArray[i];
    strDiary += '<br>------------------------------<br>';
    strDiary += '<h2>' + thisSheetName + '</h2>';
    Logger.log(thisSheetName);
    sheet = ss.getSheetByName(thisSheetName);
    dataRange = 'A3:I33';
    dataValues = sheet.getRange(dataRange).getValues();

    // Loop through data array
    for (var j = 0; j < 31; j++){

      // Get the values for this row
      var row = dataValues[j];
      // Get a string output and concat onto strDiary
      strDiary += getRowOutput(row);
      
    } // end data array loop

  } // end of loop

  //Logger.log(strDiary);
  sendEmailCheckQuotas(strDiary);

} // end of function

function getRowOutput(row){

  var strOut = '';
  //SKIP IF NO NOTES FOUND (DID NOT WORK THAT DAY)
  if (String(row[8]).length === 0) return strOut;

  //OUTPUT DATE INFORMATION
  var thisDate = new Date(row[1]);
  var thisDayOfMonth = String(thisDate.getDate()).padStart(2, '0');
  var thisDayOfWeek = String(arrDayNames[thisDate.getDay()]).substr(0,3);
  var thisMonth = thisDate.getMonth() + 1;
  var thisYear = thisDate.getFullYear();
  // Concatenate to output
  strOut += '<b>Date: ' + thisDayOfWeek + ', ' + thisYear + '-' + thisMonth + '-' + thisDayOfMonth + '</b><br>';

  //OUTPUT START/END TIMES
  var startHour = new Date(row[2]).getHours();
  var startMins = String(new Date(row[2]).getMinutes()).padEnd(2,'0');
  var endHour = new Date(row[5]).getHours();
  var endMins = String(new Date(row[5]).getMinutes()).padEnd(2,'0');
  strOut += '<em>Start: ' + startHour + ':' + startMins;
  strOut += ' - End: ' + + endHour + ':' + endMins;
  //CHECK IF A BREAK WAS TAKEN
  if (String(row[3]).length !== 0){
    var breakStartHour = new Date(row[3]).getHours();
    var breakStartMins = String(new Date(row[3]).getMinutes()).padEnd(2, '0');
    var breakEndHour = new Date(row[3]).getHours();
    var breakEndMins = String(new Date(row[3]).getMinutes()).padEnd(2,'0');
    strOut += ' (break ' + breakStartHour + ':' + breakStartMins;
    strOut += ' - ' + breakEndHour + ':' + breakEndMins + ')';
  }
  strOut += '</em><br>';

  //OUTPUT DAILY NOTES
  strOut += 'Notes: ' + row[8] + "<br><br>"; 

  return strOut; 
}

function sendEmailCheckQuotas(message){

  /* EDIT THESE */
  var emailToSendTo = strMyEmail;
  var subject = "Diary Summary";
  var message = "<html><body>" + message + "</body></html>";
  var fromName = strMyEmail;

  /* TRY-CATCH FOR EMAIL QUOTA LIMITS */
  try {
    MailApp.sendEmail(emailToSendTo, subject, message, { name: fromName, htmlBody: message });
  }catch(e){
      Browser.msgBox("Sorry, an error occurred sending the update email: \n\r" + e + "\n\rThe email text will be shown in a browser pop-up (please copy and paste)");
      Browser.msgBox(message);
  }

}

