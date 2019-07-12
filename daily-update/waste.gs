var arrHousehold = [ 
  "12/06/19", "19/06/19", "26/06/19", "03/07/19", "10/07/19", "17/07/19", "24/07/19", 
  "31/07/19", "07/08/19", "14/08/19", "21/08/19", "29/08/19", "04/09/19", "11/09/19", 
  "18/09/19", "25/09/19", "02/10/19", "09/10/19", "16/10/19", "23/10/19", "30/10/19", 
  "06/11/19", "13/11/19", "20/11/19", "27/11/19", "04/12/19", "11/12/19", "18/12/19", 
  "27/12/19", "02/01/20" 
];

var arrBlackBox = [
 "19/06/19", "03/07/19", "17/07/19", "31/07/19", "14/08/19", "29/08/19", "11/09/19", 
 "25/09/19", "09/10/19", "23/10/19", "06/11/19", "20/11/19", "04/12/19", "18/12/19", 
 "02/01/20" 
];

var arrGreenBag = [
 "12/06/19", "10/07/19", "07/08/19", "05/09/19", "02/10/19", "30/10/19", "27/11/19", 
 "27/12/19"
];


// Run AFTER getWeatherData and use
// thisPrettyDate, tmwPrettyDate
function checkWasteCollection(){

  var strWaste = "";
  var booWasteFlag = false;
  
  // Uncomment this line if using on its own
  //setupDates();
    
    // TESTING ONLY
    // tmwPrettyDate = "12/06/19";
  
  // Check tomorrow - HH,BB,GB
  if (arrHousehold.indexOf(tmwPrettyDate) >= 0){
    // Household TOMORROW
    strWaste += "Household waste collection tomorrow!<br>";
    booWasteFlag = true;
  }
  
  if (arrBlackBox.indexOf(tmwPrettyDate) >= 0 ){
    // Black Box TOMORROW
    strWaste += "Glass, Cans and Plastic Box collection tomorrow!<br>";
    booWasteFlag = true;
  }
  
  if (arrGreenBag.indexOf(tmwPrettyDate) >= 0){
    // Green Bag TOMORROW
    strWaste += "Paper and Card Bag collection tomorrow!<br>";
    booWasteFlag = true;
  }
  
  if (booWasteFlag){
    strWaste = "<h2>Waste Collection</h2>" + strWaste + "<br>";
  }
  
  //Logger.log(strWaste);
  
  return strWaste;

}
