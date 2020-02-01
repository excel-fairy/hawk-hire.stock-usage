var SPREADSHEET = {
    spreadSheet: SpreadsheetApp.getActiveSpreadsheet(),
    sheets: {
        service:{
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Service sheet"),
            equipmentOwnerCell: 'C5',
            equipmentTypeCell: 'E5',
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
        serviceTaskList: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("service task list"),
            rowsInTaskListCell: 'AH6',
            commentCellRowCell: 'AH6',
            sourceDataRange: 'AC7:AF80',
            startRange1Cell: "AK7",
            endRange1Cell: "AK8",
            startRange2Cell: "AK9",
            endRange2Cell: "AK10"
        },
        dataValidation: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data validation"),
            equipmentsRange: 'B3:B36',
            yesCell: "A2",
            noCell: "A3"
        },
        emailAutomation: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("email automation"),
            recipientCell: 'B8',
            copyRecipientCell: 'B9',
            subjectCell: 'B10',
            bodyCell: 'B11'
        },
        references: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("references"),
            stockUsageSpreadsheetIdCell: 'AE6',
            equipmentOwnerColStart0: ColumnNames.letterToColumnStart0('A'),
            equipmentTypeColStart0: ColumnNames.letterToColumnStart0('B'),
            exportFolder1ColStart0: ColumnNames.letterToColumnStart0('D'),
            exportFolder2ColStart0: ColumnNames.letterToColumnStart0('E'),
            isExportSubfoldersColStart0: ColumnNames.letterToColumnStart0('F'),
            serviceRegisterUrlColStart0: ColumnNames.letterToColumnStart0('G'),
            serviceregisterSheetNameColStart0: ColumnNames.letterToColumnStart0('H'),
            serviceRegisterCols: {
                unitNoStart0: ColumnNames.letterToColumnStart0('I'),
                engineHoursStart0: ColumnNames.letterToColumnStart0('P'),
                serviceTypeStart0: ColumnNames.letterToColumnStart0('Q'),
                serviceDateStart0: ColumnNames.letterToColumnStart0('R'),
                serviceDueStartColStart0: ColumnNames.letterToColumnStart0('S'),
                commentsStart0: ColumnNames.letterToColumnStart0('T')
            },
            referencesFirstCol: ColumnNames.letterToColumn('A'),
            referencesLastFirstCol: ColumnNames.letterToColumn('S'),
            referencesFirstRow: 3
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
    return SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.typeCell).getValue();
}
function getTaskDate(){
    return SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.taskDateCell).getValue();
}

/**
 *
 * @param equipmentOwner
 * @param equipmentType
 * @returns {{isExportSubfolders: *, ServiceRegisterCols: {serviceType: *, unitNo: *, engineHours: *, comments: *, serviceDate: *}, exportFolder1: *, serviceregisterSheetNamecol: *, equipmentOwner: *, serviceRegisterUrl: *, equipmentType: *, exportFolder2: *}}
 */
function getReferences(equipmentOwner, equipmentType) {
    var allReferences = SPREADSHEET.sheets.references.sheet.getRange(
        SPREADSHEET.sheets.references.referencesFirstRow,
        SPREADSHEET.sheets.references.referencesFirstCol,
        SPREADSHEET.sheets.references.sheet.getLastRow(),
        SPREADSHEET.sheets.references.referencesLastCol - SPREADSHEET.sheets.references.referencesFirstCol);

    var equipmentOwnerColOffset = SPREADSHEET.sheets.references.equipmentOwnerColStart0;
    var equipmentTypeColOffset = SPREADSHEET.sheets.references.equipmentTypeColStart0;
    // We know this array has exactly one element
    var referenceArray = allReferences.filter(function (reference) {
        return equipmentOwner === reference[equipmentOwnerColOffset]
            && equipmentType === reference[equipmentType];
    });
    var referenceObj = referenceArray[0];
    return {
        equipmentOwner: referenceObj[equipmentOwnerColStart0],
        equipmentType: referenceObj[equipmentTypeColStart0],
        exportFolder1: folderUrlToId(referenceObj[exportFolder1ColStart0]),
        exportFolder2: referenceObj[exportFolder2ColStart0] !== 'N/A'
            ? folderUrlToId(referenceObj[exportFolder2ColStart0])
            : null,
        isExportSubfolders: referenceObj[isExportSubfoldersColStart0] === 'Y',
        serviceRegisterId: spreadsheetUrlToId(referenceObj[serviceRegisterUrlColStart0]),
        serviceregisterSheetName: referenceObj[serviceregisterSheetNameColStart0],
        serviceRegisterCols: {
            unitNo: referenceObj[unitNoStartColStart0],
            engineHours: referenceObj[engineHoursStartColStart0],
            serviceType: referenceObj[serviceTypeStartColStart0],
            serviceDate: referenceObj[serviceDateStartColStart0],
            serviceDue: referenceObj[serviceDueStartColStart0],
            comments: referenceObj[commentsStartColStart0],
        }
    };
}
