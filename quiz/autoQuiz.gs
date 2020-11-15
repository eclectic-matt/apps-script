/**
 * NEW QUIZ SETUP
 * Edit values below to automate the sheet output
**/

// REPLACE THIS WITH THE SPREADSHEET ID TO GET QUESTIONS
const SHEET_QUESTIONS_ID = 'REPLACE_WITH_YOUR_SPREADSHEET_ID';

// ADD THE QUIZ TITLE (SHOWN ON FIRST SLIDE AND PRESENTATION FILENAME)
 const TEXT_QUIZ_TITLE = 'Family Quiz';

// THE NAME OF THE SHEET TO GET THE QUESTIONS (ALSO USED IN TITLE SLIDE)
var shtName = '24 June 2020';

// WHAT WOULD YOU LIKE THE ANSWERS SLIDE TO SAY?
var TEXT_ANSWERS_HEADER = 'Ready for answers?';
var TEXT_ANSWERS_FOOTER = 'Here they come!';

// DOES THE QUIZ HAVE A PICTURE ROUND?
// NOTE ERROR AT PRESENT - MUST BE SET TO TRUE AND DELETED, IDENTIFYING ERROR WHEN HAVE TIME!
var flgPictureRound = true;
var strPictureRoundTitle = 'DELETE THIS SLIDE!';
const TEXT_PICTURE_ROUND = 'Picture Round';

// DO YOU WANT A SLIDE TO TELL TEAMS TO MUTE THEMSELVES DURING EACH ROUND
// Include slides to show you are muted?
var flgMutedSlide = false;
 const TEXT_MUTE_HEADER = 'You are now muted!';
 const TEXT_MUTE_REMINDER = 'Remember to un-mute yourself at the end of the questions!';
 
// DO YOU WANT A BONUS ROUND SLIDE?
var flgBonusRound = false;
var TEXT_BONUS_HEADER = 'BONUS ROUND!';
var TEXT_BONUS_INFO = 'In this round the first team to give me the answer wins the points!';

// DO YOU WANT TO RANDOMISE THE ANSWERS TO RUN AS A BINGO QUIZ?
var flgBingoAnswers = false;

// DO YOU WANT A SCORES TABLE SLIDE AT THE END?
var flgScoresTable = true;
var TEXT_SCORES_HEADER = 'Scores';
var TEXT_SCORES_TABLE = 'Scores Table:';

// DO YOU WANT A FAREWELL MESSAGE SLIDE?
var flgFarewellSlide = true;
var TEXT_FAREWELL_HEADER = 'Thank you for joining us!';
var TEXT_FAREWELL_FOOTER = 'We hope to see you for the next quiz';

// THEMEING OPTIONS (BG AND TEXT COLOUR)
var colBG = '#5c2c86';
var colText = '#ffffff';

function genericQuiz() {
  
  // Create a new presentation
  
  // Set up the MAIN title page
  //   - date = sheet name
  
  // Set up the PICTURE ROUND title page
  // Add a blank page to copy the picture round into
  
  // Loop through the sheet
  // If round # different 
      // Add a new ROUND title page
      //    - title = round name
      // Add a new ROUND desc page
      //    - desc = round description
      // Add a YOU ARE MUTED page
  // Add question pages for each question (row)
  
  SpreadsheetApp.flush();
  
  var quizTitle = TEXT_QUIZ_TITLE;
  
  var shtQuestionsId = SHEET_QUESTIONS_ID;
  var shtQuestions = null;
  shtQuestions = SpreadsheetApp.openById(shtQuestionsId);
  
  var shtThisQuiz = null;
  //shtThisQuiz = shtQuestions.getSheetByName(shtName);
  var allSheets = shtQuestions.getSheets();
  
  for (var i = 0; i < allSheets.length; i++){
    //Logger.log('Checking sheet called ' + allSheets[i] + ' when i = ' + i);
    if (allSheets[i].getSheetName() === shtName){
      Logger.log('Found sheet called ' + shtName + ' so i = ' + i);
      shtThisQuiz = allSheets[i].activate();
      break;
    }
  }
  
  var rowLastQuestion = shtThisQuiz.getLastRow();
  var rngQuestions = shtThisQuiz.getRange('A1:D' + rowLastQuestion);
  
  var thisPre = SlidesApp.create(quizTitle + ' - ' + shtName);
  var preId = thisPre.getId();
  var slides = thisPre.getSlides();
  //slides[0] = SlidesApp.PageType.MASTER;
  
  // SLIDE - Title page
  Logger.log('Adding title page');
  slides[0].getBackground().setSolidFill(colBG);
  // Title Text
  slides[0].getShapes()[0].getText().clear();
  slides[0].getShapes()[0].getText().appendText(quizTitle);
  slides[0].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(72);
  slides[0].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
  // Add a box with the Date text
  slides[0].getShapes()[1].getText().clear();
  slides[0].getShapes()[1].getText().appendText(shtName);
  slides[0].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(36);
  slides[0].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);  
  
  var curRound = 1;
  var curSlide = 1;
  
  if (flgPictureRound){
    
    // SLIDE - Picture Round
    Logger.log('Adding picture round');
    slides[0].duplicate();
    // Save/Close to flush changes and re-open
    thisPre.saveAndClose();
    thisPre = SlidesApp.openById(preId);
    slides = thisPre.getSlides();
    
    // Title Text
    slides[1].getShapes()[0].getText().clear();
    slides[1].getShapes()[0].getText().appendText(TEXT_PICTURE_ROUND);
    // Title description
    slides[1].getShapes()[1].getText().clear();
    slides[1].getShapes()[1].getText().appendText(strPictureRoundTitle);
    
  }else{
    
    // SLIDE - First round details
    var roundTitle = shtQuestions.getRange('B2').getValue();
    var roundDesc = shtQuestions.getRange('F2').getValue();
    
    // SLIDE - ROUND TITLE
    thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE)
    // Save/Close to flush changes and re-open
    thisPre.saveAndClose();
    thisPre = SlidesApp.openById(preId);
    slides = thisPre.getSlides();
    curSlide++;
    // Update format
    slides[curSlide].getBackground().setSolidFill(colBG);
    // Title Text
    slides[curSlide].getShapes()[0].getText().clear();
    slides[curSlide].getShapes()[0].getText().appendText('ROUND ' + curRound);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(60);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
    slides[curSlide].getShapes()[0].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    // Body text
    slides[curSlide].getShapes()[1].getText().clear();
    slides[curSlide].getShapes()[1].getText().appendText(roundTitle.toUpperCase());
    slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(36);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
    slides[curSlide].getShapes()[1].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // SLIDE - ROUND DESCRIPTION
    thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    // Save/Close to flush changes and re-open
    thisPre.saveAndClose();
    thisPre = SlidesApp.openById(preId);
    slides = thisPre.getSlides();
    curSlide++;
    // Update format
    slides[curSlide].getBackground().setSolidFill(colBG);
    // Title Text
    slides[curSlide].getShapes()[0].getText().clear();
    slides[curSlide].getShapes()[0].getText().appendText(roundTitle);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(42);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
    slides[curSlide].getShapes()[0].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    // Body text
    slides[curSlide].getShapes()[1].getText().clear();
    slides[curSlide].getShapes()[1].getText().appendText(roundDesc);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(20);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setItalic(true);
    slides[curSlide].getShapes()[1].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    slides[curSlide].getShapes()[1].getContentAlignment().MIDDLE;
    
  }
  
  /**
   * FIRST LOOP THROUGH THE SHEET TO GET QUESTION TEXT
  **/
  for (var i = 2; i <= rowLastQuestion; i++){
   
    //var rngThisQ = shtQuestions.getRange('A' + i + 'D' + i);
    var thisRound = shtQuestions.getRange('A' + i).getValue();
    var roundTitle = shtQuestions.getRange('B' + i).getValue();
    
    if (thisRound !== curRound){
      
      curRound = thisRound;
      var roundDesc = shtQuestions.getRange('F' + i).getValue();
      
      Logger.log('Processing ' + roundTitle + ' = Round ' + thisRound);
      // SLIDE - ROUND TITLE
      //slides[curSlide].duplicate();
      thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE)
      // Save/Close to flush changes and re-open
      thisPre.saveAndClose();
      thisPre = SlidesApp.openById(preId);
      slides = thisPre.getSlides();
      curSlide++;
      // Update format
      slides[curSlide].getBackground().setSolidFill(colBG);
      // Title Text
      slides[curSlide].getShapes()[0].getText().clear();
      slides[curSlide].getShapes()[0].getText().appendText('ROUND ' + curRound);
      slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(60);
      slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
      slides[curSlide].getShapes()[0].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      // Body text
      slides[curSlide].getShapes()[1].getText().clear();
      slides[curSlide].getShapes()[1].getText().appendText(roundTitle.toUpperCase());
      slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(36);
      slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
      slides[curSlide].getShapes()[1].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      //slides[curSlide].getShapes()[1].getContentAlignment().BOTTOM;// .setParagraphAlignment(SlidesApp.AlignmentPosition.VERTICAL_CENTER);
      
      // SLIDE - ROUND DESCRIPTION
      thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      // Save/Close to flush changes and re-open
      thisPre.saveAndClose();
      thisPre = SlidesApp.openById(preId);
      slides = thisPre.getSlides();
      curSlide++;
      // Update format
      slides[curSlide].getBackground().setSolidFill(colBG);
      // Title Text
      slides[curSlide].getShapes()[0].getText().clear();
      slides[curSlide].getShapes()[0].getText().appendText(roundTitle);
      slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(42);
      slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
      slides[curSlide].getShapes()[0].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      // Body text
      slides[curSlide].getShapes()[1].getText().clear();
      slides[curSlide].getShapes()[1].getText().appendText(roundDesc);
      slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(20);
      slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
      slides[curSlide].getShapes()[1].getText().getTextStyle().setItalic(true);
      slides[curSlide].getShapes()[1].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      //slides[curSlide].getShapes()[1].getContentAlignment().BOTTOM;// .setParagraphAlignment(SlidesApp.AlignmentPosition.VERTICAL_CENTER);
      //slides[curSlide].getShapes()[1].setContentAlignment(SlidesApp.AlignmentPosition.VERTICAL_CENTER);
      slides[curSlide].getShapes()[1].getContentAlignment().MIDDLE;
      
      if (flgMutedSlide){
        
        // SLIDE - MUTED REMINDER
        thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE)
        // Save/Close to flush changes and re-open
        thisPre.saveAndClose();
        thisPre = SlidesApp.openById(preId);
        slides = thisPre.getSlides();
        curSlide++;
        // Update format
        slides[curSlide].getBackground().setSolidFill(colBG);
        // Title Text
        slides[curSlide].getShapes()[0].getText().clear();
        slides[curSlide].getShapes()[0].getText().appendText(TEXT_MUTE_HEADING);
        slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(60);
        slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
        // Body text
        slides[curSlide].getShapes()[1].getText().clear();
        slides[curSlide].getShapes()[1].getText().appendText(TEXT_MUTE_REMINDER);
        slides[curSlide].getShapes()[1].getText().getTextStyle().setItalic(true);
        slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(36);
        slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
        
      }
      
    }
    
    // SLIDE - QUESTION TEXT
    thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    // Save/Close to flush changes and re-open
    thisPre.saveAndClose();
    thisPre = SlidesApp.openById(preId);
    slides = thisPre.getSlides();
    
    curSlide++;
    var qNum = shtQuestions.getRange('C' + i).getValue();
    var qText = shtQuestions.getRange('D' + i).getValue();
    // Update format
    slides[curSlide].getBackground().setSolidFill(colBG);
    // Title Text
    slides[curSlide].getShapes()[0].getText().clear();
    slides[curSlide].getShapes()[0].getText().appendText('QUESTION ' + qNum);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(42);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
    slides[curSlide].getShapes()[0].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    // Body text
    slides[curSlide].getShapes()[1].getText().clear();
    slides[curSlide].getShapes()[1].getText().appendText(qText);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(30);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
    slides[curSlide].getShapes()[1].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    //slides[curSlide].getShapes()[1].getContentAlignment().BOTTOM;// .setParagraphAlignment(SlidesApp.AlignmentPosition.VERTICAL_CENTER);
    //slides[curSlide].getShapes()[1].setContentAlignment(SlidesApp.AlignmentPosition.VERTICAL_CENTER);
    slides[curSlide].getShapes()[1].getContentAlignment().MIDDLE;
    
  }
  
  // BONUS ROUND?
  if (flgBonusRound){
    
    // InsureThat Bonus Round?
    thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE);
    // Save/Close to flush changes and re-open
    thisPre.saveAndClose();
    thisPre = SlidesApp.openById(preId);
    slides = thisPre.getSlides();
    curSlide++;
    // Update format
    slides[curSlide].getBackground().setSolidFill(colBG);
    // Title Text
    slides[curSlide].getShapes()[0].getText().clear();
    slides[curSlide].getShapes()[0].getText().appendText(TEXT_BONUS_HEADER);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(60);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
    // Body text
    slides[curSlide].getShapes()[1].getText().clear();
    slides[curSlide].getShapes()[1].getText().appendText(TEXT_BONUS_INFO);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(36);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
    
  }
  
  /// ANSWER SLIDES
  
  thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE);
  // Save/Close to flush changes and re-open
  thisPre.saveAndClose();
  thisPre = SlidesApp.openById(preId);
  slides = thisPre.getSlides();
  curSlide++;
  // Update format
  slides[curSlide].getBackground().setSolidFill(colBG);
  // Title Text
  slides[curSlide].getShapes()[0].getText().clear();
  slides[curSlide].getShapes()[0].getText().appendText(TEXT_ANSWERS_HEADER);
  slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(60);
  slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
  // Body text
  slides[curSlide].getShapes()[1].getText().clear();
  slides[curSlide].getShapes()[1].getText().appendText(TEXT_ANSWERS_FOOTER);
  slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(36);
  slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
  
  // Add the picture round again - separately
  if (flgPictureRound){
    
    thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE);
    // Save/Close to flush changes and re-open
    thisPre.saveAndClose();
    thisPre = SlidesApp.openById(preId);
    slides = thisPre.getSlides();
    curSlide++;
    // Update format
    slides[curSlide].getBackground().setSolidFill(colBG);
    // Title Text
    slides[curSlide].getShapes()[0].getText().clear();
    slides[curSlide].getShapes()[0].getText().appendText('PICTURE ROUND!');
    slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(60);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
    // Body text
    slides[curSlide].getShapes()[1].getText().clear();
    slides[curSlide].getShapes()[1].getText().appendText('Answers');
    slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(36);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
    
  }
  
  // Then loop back through to display the answers!
  // Either randomised (bingo) or straight (normal)
  
  if (flgBingoAnswers){
    
    /**
    * RANDOMISED LOOP THROUGH THE SHEET 
    * TO GET ANSWERS IN RANDOM ORDER
    **/
    
    // Generate an array of question rows
    var arrQs = [];
    for (var i = 2; i <= rowLastQuestion; i++){
      arrQs[i - 2] = i;  
    }
    // Then shuffle that array
    arrQs = shuffle(arrQs);
    
    
    for (var i = 2; i <= rowLastQuestion; i++){
      
      //arrQs[i] = the row number to process
      var thisRow = arrQs[i - 2];
      
      //var rngThisQ = shtQuestions.getRange('A' + i + 'D' + i);
      var thisRound = shtQuestions.getRange('A' + thisRow).getValue();
      var roundTitle = shtQuestions.getRange('B' + thisRow).getValue();
      
      if (thisRound !== curRound){
        
        curRound = thisRound;
        var roundDesc = shtQuestions.getRange('F' + thisRow).getValue();
        
      }
      
      // Show round name & Question number
      thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      // Save/Close to flush changes and re-open
      thisPre.saveAndClose();
      thisPre = SlidesApp.openById(preId);
      slides = thisPre.getSlides();
      
      curSlide++;
      var qNum = shtQuestions.getRange('C' + thisRow).getValue();
      var qText = shtQuestions.getRange('D' + thisRow).getValue();
      var aText = shtQuestions.getRange('E' + thisRow).getValue();
      // Update format
      slides[curSlide].getBackground().setSolidFill(colBG);
      // Title Text
      slides[curSlide].getShapes()[0].getText().clear();
      slides[curSlide].getShapes()[0].getText().appendText('ROUND ' + curRound + ' - QUESTION ' + qNum);
      slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(42);
      slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
      slides[curSlide].getShapes()[0].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      // Body text
      slides[curSlide].getShapes()[1].getText().clear();
      slides[curSlide].getShapes()[1].getText().appendText(aText);
      slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(30);
      slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
      slides[curSlide].getShapes()[1].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      //slides[curSlide].getShapes()[1].getContentAlignment().BOTTOM;// .setParagraphAlignment(SlidesApp.AlignmentPosition.VERTICAL_CENTER);
      //slides[curSlide].getShapes()[1].setContentAlignment(SlidesApp.AlignmentPosition.VERTICAL_CENTER);
      slides[curSlide].getShapes()[1].getContentAlignment().MIDDLE;
      
    }
    
  }else{
    
  
    /**
    * STRAIGHT LOOP THROUGH THE SHEET TO GET QUESTION TEXT
    **/
    for (var i = 2; i <= rowLastQuestion; i++){
      
      //var rngThisQ = shtQuestions.getRange('A' + i + 'D' + i);
      var thisRound = shtQuestions.getRange('A' + i).getValue();
      var roundTitle = shtQuestions.getRange('B' + i).getValue();
      
      if (thisRound !== curRound){
        
        curRound = thisRound;
        var roundDesc = shtQuestions.getRange('F' + i).getValue();
        
      }
      
      // Show round name & Question number
      thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      // Save/Close to flush changes and re-open
      thisPre.saveAndClose();
      thisPre = SlidesApp.openById(preId);
      slides = thisPre.getSlides();
      
      curSlide++;
      var qNum = shtQuestions.getRange('C' + i).getValue();
      var qText = shtQuestions.getRange('D' + i).getValue();
      var aText = shtQuestions.getRange('E' + i).getValue();
      // Update format
      slides[curSlide].getBackground().setSolidFill(colBG);
      // Title Text
      slides[curSlide].getShapes()[0].getText().clear();
      slides[curSlide].getShapes()[0].getText().appendText('ROUND ' + curRound + ' - QUESTION ' + qNum);
      slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(42);
      slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
      slides[curSlide].getShapes()[0].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      // Body text
      slides[curSlide].getShapes()[1].getText().clear();
      slides[curSlide].getShapes()[1].getText().appendText(aText);
      slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(30);
      slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
      slides[curSlide].getShapes()[1].getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      //slides[curSlide].getShapes()[1].getContentAlignment().BOTTOM;// .setParagraphAlignment(SlidesApp.AlignmentPosition.VERTICAL_CENTER);
      //slides[curSlide].getShapes()[1].setContentAlignment(SlidesApp.AlignmentPosition.VERTICAL_CENTER);
      slides[curSlide].getShapes()[1].getContentAlignment().MIDDLE;
      
    }
    
  }
  
  if (flgScoresTable){
  
    // Scores table slide added at the end!
    thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE);
    // Save/Close to flush changes and re-open
    thisPre.saveAndClose();
    thisPre = SlidesApp.openById(preId);
    slides = thisPre.getSlides();
    curSlide++;
    // Update format
    slides[curSlide].getBackground().setSolidFill(colBG);
    // Title Text
    slides[curSlide].getShapes()[0].getText().clear();
    slides[curSlide].getShapes()[0].getText().appendText(TEXT_SCORES_HEADER);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(60);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
    // Body text
    slides[curSlide].getShapes()[1].getText().clear();
    slides[curSlide].getShapes()[1].getText().appendText(TEXT_SCORES_TABLE);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(36);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);

  }
  
  if (flgFarewellSlide){
  
      // Farewell table slide added at the end!
    thisPre.appendSlide(SlidesApp.PredefinedLayout.TITLE);
    // Save/Close to flush changes and re-open
    thisPre.saveAndClose();
    thisPre = SlidesApp.openById(preId);
    slides = thisPre.getSlides();
    curSlide++;
    // Update format
    slides[curSlide].getBackground().setSolidFill(colBG);
    // Title Text
    slides[curSlide].getShapes()[0].getText().clear();
    slides[curSlide].getShapes()[0].getText().appendText(TEXT_FAREWELL_HEADER);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setFontFamily('Lato').setFontSize(60);
    slides[curSlide].getShapes()[0].getText().getTextStyle().setForegroundColor(colText);
    // Body text
    slides[curSlide].getShapes()[1].getText().clear();
    slides[curSlide].getShapes()[1].getText().appendText(TEXT_FAREWELL_FOOTER);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setFontFamily('Lato').setFontSize(36);
    slides[curSlide].getShapes()[1].getText().getTextStyle().setForegroundColor(colText);
  
  }
  
}
