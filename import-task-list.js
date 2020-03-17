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
    highlightHeaderCells(taskRange);
    boldCells(taskRange);

    if (serviceSheetIsRepairMode() || serviceSheetIsInspectionMode()){
        var commentCellRow = SPREADSHEET.sheets.serviceTaskList.sheet.getRange(SPREADSHEET.sheets.serviceTaskList.commentCellRowCell).getValue();
        SPREADSHEET.sheets.service.sheet.getRange(commentCellRow, SPREADSHEET.sheets.service.taskListCoordinates.col, 1, SPREADSHEET.sheets.service.taskListCoordinates.nbCols).setBackground(BEIGE);
        if(serviceSheetIsRepairMode()) {
            var firstLineOfColumnBox = SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.taskListCoordinates.row + 1, SPREADSHEET.sheets.service.taskListCoordinates.col, 1, 4);
            // firstLineOfColumnBox.setValues([['Part used', null, 'Part no', 'Qty']]); // Second parameter is null because two columns are merged and we need to skip the merged column
            firstLineOfColumnBox.setFontWeight('bold');
            var penultimateLineFirstColOfCommentBox = SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.taskListCoordinates.row + getNbTasks() - 2, SPREADSHEET.sheets.service.taskListCoordinates.col);
            // penultimateLineFirstColOfCommentBox.setValue('Total number of hours of the job:');
            penultimateLineFirstColOfCommentBox.setFontWeight('bold');
        }
    }

    SPREADSHEET.sheets.service.sheet.getRange(
        SPREADSHEET.sheets.service.taskListCoordinates.row, SPREADSHEET.sheets.service.taskListCoordinates.col + SPREADSHEET.sheets.service.taskListCoordinates.nbCols - 1, MAX_NB_TASKS, 1)
        .setDataValidation(null);

    if (serviceSheetIsServiceMode())
        setTaskListDataValidationRules();
}


/**
 * Apply a style to all the cells on the range that match the content
 * @param range
 * @param keyWords Keywords to match cells contents against
 * @param styleFunction function to apply to a cell to update its style
 */
function applyStyle(range, keyWords, styleFunction) {
    var nbRows = range.getNumRows();
    var nbCols = range.getNumColumns();
    for (var i = 0; i < nbRows; i++)
        for (var j = 0; j < nbCols; j++){
            var cell = range.offset(i, j, 1, 1);
            if(keyWords.indexOf(cell.getValue()) !== -1)
                styleFunction(cell);
        }
}

function highlightHeaderCells(range){
    var cellsContent = ['Replace', 'Part no', 'Replaced', 'Qty', 'Additional Parts - Description', 'Inspect', 'Comments',
        'Completed', 'Parts used'];
    applyStyle(range, cellsContent, applyStyleOnCell);

    function applyStyleOnCell(cell) {
        cell.setFontWeight("bold").setBackground(BEIGE);
    }
}

function boldCells(range){
    var cellsContent = ['Client Supplied Parts', 'Total number of hours of the job:'];
    applyStyle(range, cellsContent, applyStyleOnCell);

    function applyStyleOnCell(cell) {
        cell.setFontWeight("bold");
    }
}

function setTaskListDataValidationRules(){
    // Yes/No
    var yes = SPREADSHEET.sheets.dataValidation.sheet.getRange(SPREADSHEET.sheets.dataValidation.yesCell).getValue();
    var no = SPREADSHEET.sheets.dataValidation.sheet.getRange(SPREADSHEET.sheets.dataValidation.noCell).getValue();
    var ruleYesNo = SpreadsheetApp.newDataValidation().requireValueInList([yes,no]).build();

    // Number
    var numbersRange = SPREADSHEET.sheets.dataValidation.sheet.getRange(SPREADSHEET.sheets.dataValidation.numbersrange).getValues();
    var ruleNumberList = SpreadsheetApp.newDataValidation().requireValueInList(numbersRange).build();


    var numberValidationStartRange = SPREADSHEET.sheets.serviceTaskList.sheet.getRange(
        SPREADSHEET.sheets.serviceTaskList.numberValidationStartRange).getValue();
    var numberValidationEndRange = SPREADSHEET.sheets.serviceTaskList.sheet.getRange(
        SPREADSHEET.sheets.serviceTaskList.numberValidationEndRange).getValue();
    var numberValidationRangeToApply = SPREADSHEET.sheets.service.sheet.getRange(numberValidationStartRange, 5,
        numberValidationEndRange - numberValidationStartRange + 1, 1);
    numberValidationRangeToApply.setDataValidation(ruleNumberList);

    var yesNoValidationStartRange = SPREADSHEET.sheets.serviceTaskList.sheet.getRange(
        SPREADSHEET.sheets.serviceTaskList.yesNoValidationStartRange).getValue();
    var yesNoValidationEndRange = SPREADSHEET.sheets.serviceTaskList.sheet.getRange(
        SPREADSHEET.sheets.serviceTaskList.yesNoValidationEndRange).getValue();
    var yesNoValidationRangeToApply = SPREADSHEET.sheets.service.sheet.getRange(yesNoValidationStartRange, 5,
        yesNoValidationEndRange - yesNoValidationStartRange + 1, 1);
    yesNoValidationRangeToApply.setDataValidation(ruleYesNo);


    var clientPartValidationRange1 = SPREADSHEET.sheets.serviceTaskList.sheet.getRange(
        SPREADSHEET.sheets.serviceTaskList.clientPartValidationRange1).getValue();
    var clientValidationRange1ToApply = SPREADSHEET.sheets.service.sheet.getRange(clientPartValidationRange1, 5, 1, 1);
    clientValidationRange1ToApply.setDataValidation(ruleYesNo);
    var clientPartValidationRange2 = SPREADSHEET.sheets.serviceTaskList.sheet.getRange(
        SPREADSHEET.sheets.serviceTaskList.clientPartValidationRange2).getValue();
    var clientValidationRange2ToApply = SPREADSHEET.sheets.service.sheet.getRange(clientPartValidationRange2, 5, 1, 1);
    clientValidationRange2ToApply.setDataValidation(ruleYesNo);

}
