var WHITE = '#ffffff';
var BEIGE = '#aacbe3';
var MAX_NB_TASKS = 1000;

function clearTaskList(){
    var taskListMaxRange = getTasksListRange(MAX_NB_TASKS);
    taskListMaxRange.clearContent();
    taskListMaxRange.setFontWeight("normal");
    taskListMaxRange.setBackground(WHITE);
    taskListMaxRange.setBorder(false, false, false, false, false, false);
    taskListMaxRange.setFontSize(10);
}


function importTaskList() {
    clearTaskList();

    SPREADSHEET.sheets.serviceTaskList.sheet.getRange(SPREADSHEET.sheets.serviceTaskList.sourceDataRange)
        .copyTo(SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.topLefCellOfTaskList),
            {contentsOnly: true});

    var nbTasks = getNbTasks();
    var taskRange = getTasksListRange(nbTasks);
    taskRange.setBorder(true, true, true, true, false, false);
    highlightKeyWordCells(taskRange);

    if (serviceSheetIsRepairMode() || serviceSheetIsInspectionMode()){
        var commentCellRow = SPREADSHEET.sheets.serviceTaskList.sheet.getRange(SPREADSHEET.sheets.serviceTaskList.commentCellRowCell).getValue();
        SPREADSHEET.sheets.service.sheet.getRange(commentCellRow, SPREADSHEET.sheets.service.taskListCoordinates.col, 1, SPREADSHEET.sheets.service.taskListCoordinates.nbCols).setBackground(BEIGE);
        if(serviceSheetIsRepairMode()) {
            var firstLineOfColumnBox = SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.taskListCoordinates.row + 1, SPREADSHEET.sheets.service.taskListCoordinates.col, 1, 4);
            firstLineOfColumnBox.setValues([['Part used', null, 'Part no', 'Qty']]); // Second parameter is null because two columns are merged and we need to skip the merged column
            firstLineOfColumnBox.setFontWeight('bold');
            var penultimateLineFirstColOfCommentBox = SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.taskListCoordinates.row + getNbTasks() - 2, SPREADSHEET.sheets.service.taskListCoordinates.col);
            penultimateLineFirstColOfCommentBox.setValue('Total number of hours of the job:');
            penultimateLineFirstColOfCommentBox.setFontWeight('bold');
        }
    }

    SPREADSHEET.sheets.service.sheet.getRange(
        SPREADSHEET.sheets.service.taskListCoordinates.row, SPREADSHEET.sheets.service.taskListCoordinates.col + SPREADSHEET.sheets.service.taskListCoordinates.nbCols - 1, MAX_NB_TASKS, 1)
        .setDataValidation(null);

    if (serviceSheetIsServiceMode())
        setTaskListDataValidationRules();
}


function highlightKeyWordCells(range){
    var keyWords = ['Replace', 'Part no', 'Replaced', 'Qty', 'Additional parts - Description', 'Inspect', 'Comments', 'Completed'];
    var nbRows = range.getNumRows();
    var nbCols = range.getNumColumns();
    for (var i = 0; i < nbRows; i++)
        for (var j = 0; j < nbCols; j++){
            var cell = range.offset(i, j, 1, 1);
            if(keyWords.indexOf(cell.getValue()) !== -1)
                cell.setFontWeight("bold").setBackground(BEIGE);
        }
}

function setTaskListDataValidationRules(){
    var endRange1 = SPREADSHEET.sheets.serviceTaskList.sheet.getRange("AH8").getValue();
    var startRange2 = SPREADSHEET.sheets.serviceTaskList.sheet.getRange("AH9").getValue();
    var endRange2 = SPREADSHEET.sheets.serviceTaskList.sheet.getRange("AH10").getValue();
    var dataValidYesNo1 = SPREADSHEET.sheets.service.sheet.getRange(16,5,endRange1-15,1);
    var dataValidYesNo2 = SPREADSHEET.sheets.service.sheet.getRange(startRange2+1,5,endRange2-startRange2,1);
    var yes = SPREADSHEET.sheets.dataValidation.sheet.getRange("A2").getValue();
    var no = SPREADSHEET.sheets.dataValidation.sheet.getRange("A3").getValue();
    var ruleYesNo = SpreadsheetApp.newDataValidation().requireValueInList([yes,no]).build();
    dataValidYesNo1.setDataValidation(ruleYesNo);
    dataValidYesNo2.setDataValidation(ruleYesNo);
}
