var SPREADSHEET = {
    spreadSheet: SpreadsheetApp.getActiveSpreadsheet(),
    sheets: {
        service:{
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
            },
            taskListCoordinates: {
                fullDocumentBeginningRow: 4,
                row: 15,
                col: 2,
                nbCols: 4
            }
        },
        servicePerType: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Service per type"),
            rowsInTaskListCell: 'E1',
            commentCellRowCell: 'F1',
            sourceDataRange: 'B2:E70'
        },
        serviceTaskList: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("service task list"),
            rowsInTaskListCell: 'AE6',
            commentCellRowCell: 'AE6',
            sourceDataRange: 'AC7:AC80'
        },
        dataValidation: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data validation"),
            equipmentsRange: 'B3:B36'
        },
        emailAutomation: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("email automation"),
        },
        references: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("references"),
            serviceRegisterSpreadsheetIdCell: 'AE6',
        },

    }
};

function getTasksListRange(nbLines){
    return SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.taskListCoordinates.row, SPREADSHEET.sheets.service.taskListCoordinates.col, nbLines, SPREADSHEET.sheets.service.taskListCoordinates.nbCols);
}

function getTasksListStartLineEndLine(startLineOffset, endLineOffset){
    return SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.taskListCoordinates.row + startLineOffset, SPREADSHEET.sheets.service.taskListCoordinates.col, endLineOffset, SPREADSHEET.sheets.service.taskListCoordinates.nbCols);
}

function getNbTasks(){
    return SPREADSHEET.sheets.serviceTaskList.sheet.getRange(SPREADSHEET.sheets.serviceTaskList.rowsInTaskListCell).getValue();
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
    return SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.taskTypeCell).getValue();
}

function getEquipmentNumber(){
    return SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.equipmentNumberCell).getValue();
}
function getMachineHours(){
    return SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.machineHoursCell).getValue();
}
function getTaskType(){
    var type = SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.typeCell).getValue();
    // todo
    var hrSuffix = " hr"; /* Je pense que ca doit peut etre etre change ici puisque j'ai enleve le hr aux service types, c'est juste le nombre maintenant*/
    if(type.substring(type.length - hrSuffix.length) === hrSuffix)
        type = type.substring(0, type.length - hrSuffix.length);
    return type;
}
function getTaskDate(){
    return SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.taskDateCell).getValue();
}
