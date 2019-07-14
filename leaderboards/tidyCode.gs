// PLEASE NOTE - this code is being updated and has not been tested as working
// Use main.gs for any untidy but functional implementation of the leaderboards

/* ---------------------

TIDY CODE (27-03-2019)

Gets data from a spreadsheet (leaderboardId)
Processes this into an object holding all data
Uses this data to update a COPY of the template presentation (templatePREId)
Replaces the slides from a saved presentation (thisPresentationId)
Then deletes the copy to save space/tidy up

------------------- */

  // The leaderboard source spreadsheet
var leaderboardId = "YOUR_GOOGLE_SHEET_ID";
  // The template presentation
var templatePREId = "YOUR_GOOGLE_SLIDES_TEMPLATE_PRESENTATION_ID";
  // The main presentation (this slides presentation)
var mainPREId = 'YOUR_GOOGLE_SLIDES_LIVE_PRESENTATION_ID';

  // The main object being used to collate the data
var objLeaderboard = {
  
  // Where the leaderboard data will be collated
  data: {
    allTime: [[]],
    thisMonth: [[]]
  },
  
  // Template uses moustaches for data to be updated
  moustache: {
    gameNames: ["quest", "frank", "baker"],
    typeNames: ["AT-", "TM-"],
    fieldNames: ["Name", "Time"]
  },
  
  // Information about the source spreadsheet and data
  sheetValues: {
    names: ["Quest for the Throne" , "Dr. Frankenscraft", "Baker Street" ],
    ranges: {
      allTime: "E5:F9",
      thisMonth: "E13:F15"
    },
    limits: {
      thisMonth: 3,
      allTime: 5
    }
  }
  
};

var numAllTime = objLeaderboard.sheetValues.limits.allTime;
var numThisMonth = objLeaderboard.sheetValues.limits.thisMonth;
var numGames = objLeaderboard.moustache.gameNames.length;

// This function gets the data from the source spreadsheet 
// and collates this into objLeaderboard.data
function getUpdateData() {
  
  var shtLeaderboard = SpreadsheetApp.openById(leaderboardId);
  var sheetValues = objLeaderboard.sheetValues;
  var leaderboardEntries = sheetValues.names.length;
  
  for (var i = 0; i < leaderboardEntries; i++){
   
    // ALL TIME DATA
    var rngAT = sheetValues.names[i] + "!" + sheetValues.ranges.allTime;
    var dataValues = shtLeaderboard.getRange(rngAT).getValues();
    
    objLeaderboard.data.allTime[i] = [];
    
    for (var j = 0; j < numAllTime; j++){
      objLeaderboard.data.allTime[i][j] = [ dataValues[j][0] , dataValues[j][1] ];
    }
    
    // THIS MONTH DATA
    var rngTM = sheetValues.names[i] + "!" + sheetValues.ranges.thisMonth;
    var dataValues = shtLeaderboard.getRange(rngTM).getValues();
    
    objLeaderboard.data.thisMonth[i] = [];
    
    for (var j = 0; j < numThisMonth; j++){
      objLeaderboard.data.thisMonth[i][j] = [ dataValues[j][0] , dataValues[j][1] ];
    }
    
  }
  
  //Logger.log(objLeaderboard.data);
  
}



function pushDataToPresentation(){
  
  /*
  var copyTitle = 'Lakes Escapes - LIVE leaderboard (' + Date() + ')';
  var copyFile = {
    title: copyTitle,
    parents: [{id: 'root'}]
  };
  // A duplicate of the template presentation (the current app)
  copyFile = Drive.Files.copy(copyFile, templatePREId );
  // Copied presentation file ID
  var presentationCopyId = copyFile.id;
  */
  
  var requests = [];
  
  var e = 0;
  
  //var arrLeaderboardData = Object.prototype.values(objLeaderboard.data);
  //var arrLeaderboardData = Object.values(objLeaderboard.data);
  //var arrLeaderboardData = objLeaderboard.data.values();
  //var arrLeaderboardData = Array(objLeaderboard.data); //Single element, the object
  //var arrLeaderboardData = Array.from(objLeaderboard.data);
  //var arrLeaderboardData = Array(objLeaderboard.data).values();
  //var arrLeaderboardData = Object.keys(objLeaderboard.data).map(function(key) {
  //  return [Number(key), objLeaderboard.data[[key]]];
  //});
  var arrLeaderboardData = Object.entries(objLeaderboard["data"]);

  Logger.log(arrLeaderboardData);
  
  for (var i = 0; i < numGames; i++){
    
    var strGameName = objLeaderboard.moustache.gameNames[i];
    
    /* ----
    
    NEED TO UPDATE FROM HERE ONWARDS
    
    */
    
    var scoreTypes = objLeaderboard.moustache.typeNames;
    var scoreFields = objLeaderboard.moustache.fieldNames;
    
    for (var gameType = 0; gameType < scoreTypes.length; gameType++){
     
      var thisType = scoreTypes[gameType];
      
      
      
      for (var scoreField = 0; scoreField < scoreFields.length; scoreField++){
       
        var thisScore = scoreFields[scoreField];
        
        Logger.log('This Score'+ arrLeaderboardData[gameType][scoreField] );
        
        Logger.log('type = '+thisType+'. field = '+thisScore+'.');
        
      }
      
    }
    
    // Update all time scores
    for (var j = 0; j < numAllTime; j++){
    
      var thisRequest = {};
      thisRequest.replaceAllText = {
        containsText: {
          text: '{{questAT-Name' + (j+1) + '}}',
          matchCase: true
        },
        replaceText: questAllTime[j][0]
      }
      requests[e] = thisRequest;
      
    }
    
  }
  
  
  
  
  
  
  
  
  
  
  
  
  
  // ------------- OLD ------------------
  
  
  
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
  /*
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
  */
  
}

function fullUpdate(){
  getUpdateData();
  pushDataToPresentation();
}

function getThisPresentationID(){
  var thisId = SlidesApp.getActivePresentation().getId();
  //Logger.log('This ID = ' + thisId);
  return thisId;
}
