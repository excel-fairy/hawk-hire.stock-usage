var SPREADSHEET = {
    spreadSheet: SpreadsheetApp.getActiveSpreadsheet(),
    sheets: {
        serviceSheet:{
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Service sheet"),
            equipmentNumberCell: 'E6',
            taskListNameCell: 'B14',
            taskTypeCell: 'C11',
            typeCell: 'C12',
            topLefCellOfTaskList: 'B15',
            machineHoursCell: 'C13',
            taskDateCell: 'C6',
            importTaskListButtonCell: 'J5',
            exportSheetButtonCell: 'J10',
            partsCol: ColumnNames.letterToColumn('D'),
            quantityCol: ColumnNames.letterToColumn('E'),
            serviceMode: {
                firstEntryRow: 16
            },
            repairMode: {
                firstEntryRow: 17
            }

        },
        servicePerTypeSheet: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Service per type"),
            rowsInTaskListCell: 'E1',
            commentCellRowCell: 'F1',
            sourceDataRange: 'B2:E70'
        },
        dataValidSheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Valid")

    }
};
var SERVICE_REGISTER_SPREADSHEET = {
    valuesRange: 'A3:AA47',
    hoursColumnOffset: 2,
    lastServiceCompletedColumnOffset: 3,
    dateLastServiceCompletedColumnOffset: 4,
    nextServiceDueColumnOffset: 5,
    partsRequiredForNextService: 6
};

var WHITE = '#ffffff';
var BEIGE = '#aacbe3';
var MAX_NB_TASKS = 1000;
var TASK_LIST_COORDINATES = {
    fullDocumentBeginningRow: 4,
    row: 15,
    col: 2,
    nbCols: 4
};

function importTaskList() {
    clearTaskList();

    SPREADSHEET.sheets.servicePerTypeSheet.sheet.getRange(SPREADSHEET.sheets.servicePerTypeSheet.sourceDataRange)
        .copyTo(SPREADSHEET.sheets.serviceSheet.sheet
                .getRange(SPREADSHEET.sheets.serviceSheet.topLefCellOfTaskList),
            {contentsOnly: true});


    var nbTasks = getNbTasks();
    var taskRange = getTasksListRange(nbTasks);
    taskRange.setBorder(true, true, true, true, false, false);
    highlightKeyWordCells(taskRange);

    if (serviceSheetIsRepairMode() || serviceSheetIsInspectionMode()){
        var commentCellRow = SPREADSHEET.sheets.servicePerTypeSheet.sheet.getRange(SPREADSHEET.sheets.servicePerTypeSheet.commentCellRowCell).getValue();
        SPREADSHEET.sheets.serviceSheet.sheet.getRange(commentCellRow, TASK_LIST_COORDINATES.col, 1, TASK_LIST_COORDINATES.nbCols).setBackground(BEIGE);
        if(serviceSheetIsRepairMode()) {
            var firstLineOfColumnBox = SPREADSHEET.sheets.serviceSheet.sheet.getRange(TASK_LIST_COORDINATES.row + 1, TASK_LIST_COORDINATES.col, 1, 4);
            firstLineOfColumnBox.setValues([['Part used', null, 'Part no', 'Qty']]); // Second parameter is null because two columns are merged and we need to skip the merged column
            firstLineOfColumnBox.setFontWeight('bold');
            var penultimateLineFirstColOfCommentBox = SPREADSHEET.sheets.serviceSheet.sheet.getRange(TASK_LIST_COORDINATES.row + getNbTasks() - 2, TASK_LIST_COORDINATES.col);
            penultimateLineFirstColOfCommentBox.setValue('Total number of hours of the job:');
            penultimateLineFirstColOfCommentBox.setFontWeight('bold');
        }
    }

    SPREADSHEET.sheets.serviceSheet.sheet.getRange(TASK_LIST_COORDINATES.row, TASK_LIST_COORDINATES.col + TASK_LIST_COORDINATES.nbCols - 1, MAX_NB_TASKS, 1).setDataValidation(null);

    if (serviceSheetIsServiceMode())
        setTaskListDataValidationRules();
}

function setTaskListDataValidationRules(){
    // var startRange1 = SPREADSHEET.sheets.dataValidSheet.getRange("Z4").getValue();
    var endRange1 = SPREADSHEET.sheets.dataValidSheet.getRange("Z5").getValue();
    var startRange2 = SPREADSHEET.sheets.dataValidSheet.getRange("Z7").getValue();
    var endRange2 = SPREADSHEET.sheets.dataValidSheet.getRange("Z8").getValue();
    var dataValidYesNo1 = SPREADSHEET.sheets.serviceSheet.sheet.getRange(16,5,endRange1-15,1);
    var dataValidYesNo2 = SPREADSHEET.sheets.serviceSheet.sheet.getRange(startRange2+1,5,endRange2-startRange2,1);
    var yes = SPREADSHEET.sheets.servicePerTypeSheet.sheet.getRange("F2").getValue();
    var no = SPREADSHEET.sheets.servicePerTypeSheet.sheet.getRange("F3").getValue();
    var ruleYesNo = SpreadsheetApp.newDataValidation().requireValueInList([yes,no]).build();
    dataValidYesNo1.setDataValidation(ruleYesNo);
    dataValidYesNo2.setDataValidation(ruleYesNo);
}

function getTasksListRange(nbLines){
    return SPREADSHEET.sheets.serviceSheet.sheet.getRange(TASK_LIST_COORDINATES.row, TASK_LIST_COORDINATES.col, nbLines, TASK_LIST_COORDINATES.nbCols);
}

function getTasksListStartLineEndLine(startLineOffset, endLineOffset){
    return SPREADSHEET.sheets.serviceSheet.sheet.getRange(TASK_LIST_COORDINATES.row + startLineOffset, TASK_LIST_COORDINATES.col, endLineOffset, TASK_LIST_COORDINATES.nbCols);
}

function getNbTasks(){
    return SPREADSHEET.sheets.servicePerTypeSheet.sheet.getRange(SPREADSHEET.sheets.servicePerTypeSheet.rowsInTaskListCell).getValue();
}

function clearTaskList(){
    var taskListMaxRange = getTasksListRange(MAX_NB_TASKS);
    taskListMaxRange.clearContent();
    taskListMaxRange.setFontWeight("normal");
    taskListMaxRange.setBackground(WHITE);
    taskListMaxRange.setBorder(false, false, false, false, false, false);
    taskListMaxRange.setFontSize(10);
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

function serviceSheetIsInspectionMode(){
    return getTask() === "Inspection";
}

function serviceSheetIsServiceMode(){
    return getTask() === "Service";
}

function serviceSheetIsRepairMode(){
    return getTask() === "Repair";
}

function getTask(){
    return SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.taskTypeCell).getValue();
}

function getEquipmentNumber(){
    return SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.equipmentNumberCell).getValue();
}
function getMachineHours(){
    return SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.machineHoursCell).getValue();
}
function getTaskType(){
    var type = SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.typeCell).getValue();
    var hrSuffix = " hr";
    if(type.substring(type.length - hrSuffix.length) === hrSuffix)
        type = type.substring(0, type.length - hrSuffix.length);
    return type;
}
function getTaskDate(){
    return SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.taskDateCell).getValue();
}

function exportToPdf() {
    var equipmentNumber = getEquipmentNumber();
    var exportFolderId = EXPORT_FOLDER_ID;
    var fileName = SPREADSHEET.sheets.serviceSheet.sheet.getRange(SPREADSHEET.sheets.serviceSheet.taskListNameCell).getValue();

    var exportOptions = {
        exportFolderId: exportFolderId,
        exportFileName: fileName,
        range: {
            r1: TASK_LIST_COORDINATES.fullDocumentBeginningRow - 1,
            r2: TASK_LIST_COORDINATES.row + getNbTasks(),
            c1: TASK_LIST_COORDINATES.col - 1,
            c2: TASK_LIST_COORDINATES.col + TASK_LIST_COORDINATES.nbCols - 1
        },
        repeatHeader: true,
        fileFormat: 'pdf'
    };
//    ExportSpreadsheet.export(exportOptions);
    var pdfFile = exportspreadsheet.export(exportOptions);
     sendEmail(pdfFile);
     exportPartsToDatabase();
}

function sendEmail(attachment) {
    var recipient = SPREADSHEET.sheets.dataValidSheet.getRange("K21").getValue();
    var subject = SPREADSHEET.sheets.dataValidSheet.getRange("K22").getValue();
    var message = SPREADSHEET.sheets.dataValidSheet.getRange("K23").getValue();
    var emailOptions = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: 'Automatic service sheet form mail sender'
    };
    MailApp.sendEmail(recipient, subject, message, emailOptions);
}



function getFolderToExportPdfTo(folderName){
    var parentFolder = DriveApp.getFolderById(EXPORT_FOLDER_ID);
    var folders = parentFolder.getFolders();
    while (folders.hasNext()){
        var folder = folders.next();
        if(folder.getName() === folderName)
            return folder;
    }
    var otherFolder = parentFolder.getFoldersByName("Other");
    if(otherFolder.hasNext())
        return otherFolder.next();
    else
        return null;
}


function getComments(){
    var i;
    var firstRowOffset, nbRows;
    if(serviceSheetIsServiceMode()){
        var tasksListRange = getTasksListRange(getNbTasks());
        var tasksListValues = tasksListRange.getValues();
        for (i = 0; i < tasksListValues.length; i++) {
            var firstCellContent = tasksListValues[i][0];
            if(!firstRowOffset && firstCellContent === 'Additional parts - Description')
                firstRowOffset = i+1;
            if(!!firstRowOffset && !nbRows && firstCellContent === 'Inspect')
                nbRows = i - firstRowOffset;
        }
        if(!firstRowOffset || !nbRows) // Either beginning or end of comment section not found
            return '';
    }
    if(serviceSheetIsInspectionMode() || serviceSheetIsRepairMode()){
        firstRowOffset = 1;
        nbRows = getNbTasks();
    }
    var commentsRange = getTasksListStartLineEndLine(firstRowOffset, nbRows);
    var commentsValues = commentsRange.getValues();
    var retVal = '';
    for(i=0; i < commentsValues.length; i++){
        var line = '';
        for(var j=0; j < commentsValues[i].length; j++){
            var comment = commentsValues[i][j];
            if(comment !== '')
                line += comment + ' ';
        }
        if(line !== '')
            retVal += line.trim() + '\n';
    }
    retVal = retVal.trim();
    return retVal;
}