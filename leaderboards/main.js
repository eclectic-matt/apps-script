/* ---------------------

WORKING (LAZY) CODE (25-03-2019)

Gets data from a spreadsheet (leaderboardId)
Processes this into arrays (questAllTime, questThisMonth etc)
Uses this data to update a COPY of the template presentation (templatePREId)
Replaces the slides from a saved presentation (thisPresentationId)
Then deletes the copy to save space/tidy up

------------------- */

var leaderboardId = "YOUR_GOOGLE_SHEET_ID";
var templatePREId = "YOUR_TEMPLATE_PRESENTATION_ID";
var thisPresentationId = 'YOUR_CURRENT_PRESENTATION_ID';

var questAllTime = [[]];
var questThisMonth = [[]];

var frankAllTime = [[]];
var frankThisMonth = [[]];

var bakerAllTime = [[]];
var bakerThisMonth = [[]];

var sheetNames = ["Quest for the Throne" , "Dr. Frankenscraft", "Baker Street" ];
var rangeAllTime = "E5:F9";
var rangeThisMonth = "E13:F15";


function fullUpdate(){
  getUpdateData();
  pushDataToPresentation();
}

function getThisPresentationID(){
  var thisId = SlidesApp.getActivePresentation().getId();
  //Logger.log('This ID = ' + thisId);
  return thisId;
}

function getUpdateData() {
  
  
  var shtLeaderboard = SpreadsheetApp.openById(leaderboardId);
  /* QUEST DATA */
  
  // All Time Top 5
  var questDataRange = 'Quest for the Throne!E5:F9';
  var dataValues = shtLeaderboard.getRange(questDataRange).getValues();
  // Add to data array
  for (var i = 0; i < 5; i++){   
    questAllTime[i] = [ dataValues[i][0] , dataValues[i][1] ];  
  }
  //Logger.log(questAllTime);
  
  // This Month Top 3
  questDataRange = 'Quest for the Throne!E13:F15';
  dataValues = shtLeaderboard.getRange(questDataRange).getValues();
  // Add to data array
  for (var i = 0; i < 3; i++){   
    questThisMonth[i] = [ dataValues[i][0] , dataValues[i][1] ];  
  }
  //Logger.log(questThisMonth);
  
  
  /* FRANK DATA */
  
  // All Time Top 5
  var frankDataRange = 'Dr. Frankenscraft!E5:F9';
  var dataValues = shtLeaderboard.getRange(frankDataRange).getValues();
  // Add to data array
  for (var i = 0; i < 5; i++){   
    frankAllTime[i] = [ dataValues[i][0] , dataValues[i][1] ];  
  }
  //Logger.log(frankAllTime);
  
  // This Month Top 3
  frankDataRange = 'Dr. Frankenscraft!E13:F15';
  dataValues = shtLeaderboard.getRange(frankDataRange).getValues();
  // Add to data array
  for (var i = 0; i < 3; i++){   
    frankThisMonth[i] = [ dataValues[i][0] , dataValues[i][1] ];  
  }
  //Logger.log(frankThisMonth);
  
  
  /* BAKER STREET DATA */
  
  // All Time Top 5
  var bakerDataRange = 'Baker Street!E5:F9';
  var dataValues = shtLeaderboard.getRange(bakerDataRange).getValues();
  // Add to data array
  for (var i = 0; i < 5; i++){   
    bakerAllTime[i] = [ dataValues[i][0] , dataValues[i][1] ];  
  }
  //Logger.log(bakerAllTime);
  
  // This Month Top 3
  bakerDataRange = 'Baker Street!E13:F15';
  dataValues = shtLeaderboard.getRange(bakerDataRange).getValues();
  // Add to data array
  for (var i = 0; i < 3; i++){   
    bakerThisMonth[i] = [ dataValues[i][0] , dataValues[i][1] ];  
  }
  //Logger.log(bakerThisMonth);
  
}

function pushDataToPresentation(){
  
  var copyTitle = 'Lakes Escapes - LIVE leaderboard (' + Date() + ')';
  var copyFile = {
    title: copyTitle,
    parents: [{id: 'root'}]
  };
  // A duplicate of the template presentation (the current app)
  copyFile = Drive.Files.copy(copyFile, templatePREId );
  // Copied presentation file ID
  var presentationCopyId = copyFile.id;

  var requests = [];
  
  var e = 0;
  
  // ALL TIME UPDATES
  for (var i = 0; i < 5; i++){
    
    // QUEST
    var thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{questAT-Name' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: questAllTime[i][0]
    }
    requests[e] = thisRequest;
    e += 1;
    //
    thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{questAT-Time' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: questAllTime[i][1]
    }
    requests[e] = thisRequest;
    e += 1;
    
    // BAKER
    var thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{bakerAT-Name' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: bakerAllTime[i][0]
    }
    requests[e] = thisRequest;
    e += 1;
    //
    thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{bakerAT-Time' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: bakerAllTime[i][1]
    }
    requests[e] = thisRequest;
    e += 1;

    // FRANK    
    var thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{frankAT-Name' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: frankAllTime[i][0]
    }
    requests[e] = thisRequest;
    e += 1;
    //
    thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{frankAT-Time' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: frankAllTime[i][1]
    }
    requests[e] = thisRequest;
    e += 1;
    
  }
  
  // ------------------
  // THIS MONTH UPDATES
  // ------------------
  for (var i = 0; i < 3; i++){
    
    // QUEST
    var thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{questTM-Name' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: questThisMonth[i][0]
    }
    requests[e] = thisRequest;
    e += 1;
    //
    thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{questTM-Time' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: questThisMonth[i][1]
    }
    requests[e] = thisRequest;
    e += 1;
    
    // BAKER
    var thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{bakerTM-Name' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: bakerThisMonth[i][0]
    }
    requests[e] = thisRequest;
    e += 1;
    //
    thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{bakerTM-Time' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: bakerThisMonth[i][1]
    }
    requests[e] = thisRequest;
    e += 1;
    
    // FRANK
    var thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{frankTM-Name' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: frankThisMonth[i][0]
    }
    requests[e] = thisRequest;
    e += 1;
    //
    thisRequest = {};
    thisRequest.replaceAllText = {
      containsText: {
        text: '{{frankTM-Time' + (i+1) + '}}',
        matchCase: true
      },
      replaceText: frankThisMonth[i][1]
    }
    requests[e] = thisRequest;
    e += 1;
    
  }
  
  // Execute the requests onto the copy presentation
  var result = Slides.Presentations.batchUpdate({
    requests: requests
  }, presentationCopyId);
  
  // Count the total number of replacements made.
  var numReplacements = 0;
  result.replies.forEach(function(reply) {
    numReplacements += reply.replaceAllText.occurrencesChanged;
  });
  
  // REMOVE AND THEN APPEND
  var livePresentation = SlidesApp.getActivePresentation();
  var copyPresentationSlides = SlidesApp.openById(presentationCopyId).getSlides();
  
  livePresentation.getSlides()[0].remove();
  livePresentation.getSlides()[0].remove();
  livePresentation.getSlides()[0].remove();
  
  livePresentation.appendSlide(copyPresentationSlides[0]);
  livePresentation.appendSlide(copyPresentationSlides[1]);
  livePresentation.appendSlide(copyPresentationSlides[2]);
  
  Drive.Files.remove(presentationCopyId);
  
  
}
