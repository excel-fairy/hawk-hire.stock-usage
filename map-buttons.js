function onOpen() {
    SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.importTaskListButtonCell).setValue(false);
    SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.exportSheetButtonCell).setValue(false);
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
    if(range.getSheet().getName() === SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.importTaskListButtonCell).getSheet().getName()
        && range.getA1Notation() === SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.importTaskListButtonCell).getA1Notation()
        && range.getValue() === true){
        range.setValue(false);
        importTaskList();
    }
    else if(range.getSheet().getName() === SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.exportSheetButtonCell).getSheet().getName()
        && range.getA1Notation() === SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.exportSheetButtonCell).getA1Notation()
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