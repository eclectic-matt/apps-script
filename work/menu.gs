function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Xepta Tasks')
      .addItem('Send Diary Email', 'generateDiary')
      .addSeparator()
      .addItem('Send Hours Email', 'sendTimesheet')
      .addToUi();
}
