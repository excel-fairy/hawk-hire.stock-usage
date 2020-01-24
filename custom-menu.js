function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Run script')
      .addItem('Import Task List', 'importTaskList')
      .addItem('Export Sheet and save in drive', 'exportToPdf')
      .addToUi();
}
