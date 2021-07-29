/**
 * Gathers details of timesheet hours, gets totals by week
 * and outputs by email for sending to line manager
 */

//WHERE THE EMAIL COMES FROM
var strEmailSender = 'SENDER_EMAIL_HERE';
//WHERE THE SUMMARY EMAIL IS SENT
var strEmailReceiver = 'RECEIVER_EMAIL_HERE';

/**
 * The main function to generate and send timesheet hours
 */
function sendTimesheet() {
  
  //Get the most recent sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //Get the latest spreadsheet
  var latestSheet = ss.getSheets()[0];

  //GET THE RANGE (COLUMN) WHERE "Submitted" IS RECORDED (last timesheet)
  var submittedFlagRangeStr = 'A3:A33';
  var submittedValues = latestSheet.getRange(submittedFlagRangeStr).getValues();

  //THE LAST ROW WHERE "Submitted" FOUND
  var lastSubmittedRow = false;

  //LOOP THROUGH TO FIND THE LAST SUBMITTED (lastSubmittedRow gets overwritten if multiple)
  for (var i = 0; i < submittedValues.length; i++){

    //GET THIS CELL'S VALUE
    var thisCell = submittedValues[i];
    //IF THE CELL CONTAINS "Submitted"
    if (String(thisCell) === 'Submitted'){

      //SET THE LAST SUBMITTED ROW = 3 (start row, ignore headings) + i (index of rows)
      lastSubmittedRow = 3 + i;
    }
  }

  //#-#-#-#-#-#-#-#-#-#-#
  //Determine if submission
  //starts in this/last month
  //#-#-#-#-#-#-#-#-#-#-#
  if (lastSubmittedRow === false){

    //Last submit was in the previous month - get last month's spreadsheet
    var previousSheet = ss.getSheets()[1];

    //Loop through the "Submitted" column to get the last "Submitted" row
    submittedValues = previousSheet.getRange(submittedFlagRangeStr).getValues();

    //START FROM THE END OF THIS SHEET
    for (var i = submittedValues.length; i > 0; i--){

      var thisCell = submittedValues[i];

      if (String(thisCell) === 'Submitted'){

        //SET THE LAST SUBMITTED ROW = 3 (start row, ignore headings) + i (index of rows)
        lastSubmittedRow = 3 + i;
        //BREAK OUT OF THE FOR-LOOP - LAST SUBMITTED FOUND
        break;
      }
    }
    
    Logger.log('Last submitted Last Month, Row = ' + lastSubmittedRow);
    //GET THE VALUES OF CELLS IN LAST MONTH'S SUBMITTABLE RANGE
    var lastMonthData = getMonthValues(previousSheet, lastSubmittedRow + 1);
    //GET THE VALUES OF CELLS IN THIS MONTH'S SUBMITTABLE RANGE
    var thisMonthData = getMonthValues(latestSheet, false);
    //MERGE THE TWO ARRAYS TO PROCESS AT ONCE
    var timesheetData = lastMonthData.concat(thisMonthData);
  }else{

    //Last submit was this month, easy!
    Logger.log('Last submitted This Month, Row = ' + lastSubmittedRow);
    //GET THE VALUES OF CELLS IN THIS MONTH'S SUBMITTABLE RANGE
    var timesheetData = getMonthValues(latestSheet, lastSubmittedRow + 1);
  }

  //Logger.log(timesheetData);

  //OUTPUT THE TIMESHEET DATA IN THE REQUIRED FORMAT
  outputTimesheetData(timesheetData);  
}

/**
 * Get the timesheet values for this month
 * @param Spreadsheet sheet The sheet to check
 * @param int firstRow      The first row to start checking from
 * 
 * @return array The timesheet data in array format
 */
function getMonthValues(sheet, firstRow){

  var returnVals = [];

  var firstEmptyRow = getFirstEmptyRow(sheet, 3, 3);
  //Logger.log('First empty row = ' + firstEmptyRow);

  //IF NO FIRST ROW SPECIFIED
  if (firstRow === false){

    //JUST GET EVERYTHING UP TO THE FIRST EMPTY ROW
    var thisSubmissionRange = 'B3:G' + firstEmptyRow;
    returnVals = sheet.getRange(thisSubmissionRange).getValues();
  }else{

    //Note: setting the full range, but will only output if dates noted
    var thisSubmissionRange = 'B' + firstRow + ':G' + firstEmptyRow;
    returnVals = sheet.getRange(thisSubmissionRange).getValues();
  }

  return returnVals;
}

/**
 * Output the timesheet hours in the required format
 * @param Array data The values from the sheet in the submittable range
 */
function outputTimesheetData(data){

  //START STRING OUTPUT - TABLE FORMAT
  //var strOut = '<table><tr><th style="border: 1px solid black;">Week</th><th style="border: 1px solid black;">Hours</th></tr>';
  //START STRING OUTPUT - LIST FORMAT
  var strOut = '<ul>';

  //SET THE START DATE TO THE FIRST ROW OF DATA
  var startDate = new Date(data[0][0]);
  //FORMAT THE DATE STRING (YYYY-MM-DD)
  var startDateString = startDate.getFullYear() + '-'
      + String(startDate.getMonth() + 1).padStart(2, '0') + '-' 
      + String(startDate.getDate()).padStart(2, '0');

  //SET THE WEEKLY TOTAL
  var weeklyTotal = 0;

  //LOOP THROUGH THE DATA VALUES
  for (var row = 0; row < data.length; row++){

    //GET THE CURRENT DATE
    var fullDate = new Date(data[row][0]);
    //FORMAT THE DATE STRING (YYYY-MM-DD)
    var thisDate = fullDate.getFullYear() + '-'
      + String(fullDate.getMonth() + 1).padStart(2, '0') + '-' 
      + String(fullDate.getDate()).padStart(2, '0');

    //DISPLAY HOURS TOTAL TO 2DP
    var thisHours = parseFloat(data[row][5]);
    //ADD TODAY'S HOURS TO THE WEEKLY TOTAL (AS FLOAT, FORMAT toFixed LATER)
    weeklyTotal = weeklyTotal + thisHours;

    //IF A SUNDAY, FINISH WEEK
    if (fullDate.getDay() === 0){

      //strOut += outputTableRow(startDateString, thisDate, weeklyTotal);
      strOut += outputListItem(startDateString, thisDate, weeklyTotal);

      //IF THERE ARE STILL ROWS TO PROCESS
      if (row < data.length){

        //RESET VALUES
        weeklyTotal = 0;
        startDate = new Date(data[row + 1][0]);
        startDateString = startDate.getFullYear() + '-'
        + String(startDate.getMonth() + 1).padStart(2, '0') + '-' 
        + String(startDate.getDate()).padStart(2, '0');
      } // end if row < data.length      
    } // end if getDay === 6
  } // end for (data)

  //IF WE HAVE REACHED THE END OF THE DATA, BUT HAVEN'T CLOSED OUT THE CURRENT PERIOD
  if (weeklyTotal !== 0){

    //GET THE FINAL DATE IN THE DATA ARRAY
    fullDate = data[data.length - 1][0];
    //FORMAT THE DATE STRING (YYYY-MM-DD)
    thisDate = fullDate.getFullYear() + '-'
      + String(fullDate.getMonth() + 1).padStart(2, '0') + '-' 
      + String(fullDate.getDate()).padStart(2, '0');

    //strOut += outputTableRow(startDateString, thisDate, weeklyTotal);
    strOut += outputListItem(startDateString, thisDate, weeklyTotal);
  }

  //COMPLETE STRING OUTPUT
  //strOut += '</table>';
  strOut += '</ul><br>';
  strOut += 'No holiday/sickdays/furlough/training/overtime during this period.';
  
  //SEND VIA EMAIL
  sendEmailCheckQuotas(strOut, 'Hours');
}


/**
 * Output the current period's hours in list item format
 * @param string startDate  The start of this period
 * @param string endDate    The end of this period
 * @param float hours       The total hours in this period
 * 
 * @return string A formatted <li> element
 */
function outputListItem(startDate, endDate, hours){
  
  strReturn = '<li>';
  strReturn += startDate + ' to ' + endDate + ' = ';
  strReturn += parseFloat(hours).toFixed(2) + ' hours';
  strReturn += '</li>';
  return strReturn;
}

/**
 * Output the current period's hours in table row format
 * @param string startDate  The start of this period
 * @param string endDate    The end of this period
 * @param float hours       The total hours in this period
 * 
 * @return string A formatted <tr> element
 */

function outputTableRow(startDate, endDate, hours){
  
  strReturn = '<tr><td style="border: 1px solid black;">';
  strReturn += startDate + ' -> ' + endDate + '</td>';
  strReturn += '<td style="border: 1px solid black;">' + parseFloat(hours).toFixed(2) + '</td></tr>';
  return strReturn;
}

/**
 * Get the first empty row in a specific sheet and column
 * @param Spreadsheet sheet The sheet to check
 * @param int column        The column to check
 * @param int startRow      The row to start checking from
 * 
 * @return int The first empty row found
 */
function getFirstEmptyRow(sheet, column, startRow = 1){
  lastCellVal = false;
  for (var i = startRow; i < sheet.getLastRow(); i++){
    //Logger.log('Checking cell ' + String.fromCharCode(64 + column) + i);
    var thisCellVal = sheet.getRange(String.fromCharCode(64 + column) + i).getValues()[0];
    if (String(thisCellVal).length === 0){
      var dateCellVal = sheet.getRange(String.fromCharCode(64 + column - 1) + i).getValues()[0];
      var testDate = new Date(dateCellVal);
      var testDayOfWeek = testDate.getDay();
      //IGNORE BLANKS IN WEEKENDS
      if (testDayOfWeek === 0 || testDayOfWeek === 6){
        //Weekend - expected
        //lastCellVal = i;
      }else{
        return i;
        //lastCellVal = i;
      }
    }
  }
  return lastCellVal;
}
