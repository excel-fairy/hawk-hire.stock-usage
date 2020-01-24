function onOpen() {
    SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.importTaskListButtonCell).setValue(false);
    SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.exportSheetButtonCell).setValue(false);
    SpreadsheetApp.getUi()
        .createMenu('Run scripts')
        .addItem('Import Task List', 'importTaskList')
        .addItem('Export Sheet and save in drive', 'exportToPdf')
        .addItem('Authorize scripts to access Google drive from smartphone', 'createInstallableTriggers')
        .addToUi();
}

function createInstallableTriggers(){
    deleteAllTriggers();
    ScriptApp.newTrigger('installableOnEdit')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();
}

function installableOnEdit(e){
    var range = e.range;
    if(range.getSheet().getName() === SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.importTaskListButtonCell).getSheet().getName()
        && range.getA1Notation() === SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.importTaskListButtonCell).getA1Notation()
        && range.getValue() === true){
        range.setValue(false);
        importTaskList();
    }
    else if(range.getSheet().getName() === SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.exportSheetButtonCell).getSheet().getName()
        && range.getA1Notation() === SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.exportSheetButtonCell).getA1Notation()
        && range.getValue() === true){
        range.setValue(false);
        exportToPdf();
    }
}

function deleteAllTriggers() {
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
        ScriptApp.deleteTrigger(allTriggers[i]);
    }
}