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
            },
            specialPartsCellsContents: {
                clientSuppliedParts: 'Client Supplied Parts',
                totalNumberHoursOfJob: 'Total number of hours of the job:'
            }
        },
        serviceTaskList: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("service task list"),
            rowsInTaskListCell: 'AH6',
            commentCellRowCell: 'AH6',
            sourceDataRange: 'AC7:AF80',
            numberValidationStartRange: "AK7",
            numberValidationEndRange: "AK8",
            yesNoValidationStartRange: "AK9",
            yesNoValidationEndRange: "AK10",
            clientPartValidationRange1: "AK12",
            clientPartValidationRange2: "AK13",
        },
        dataValidation: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data validation"),
            equipmentsRange: 'J3:J60',
            yesCell: "A2",
            noCell: "A3",
            numbersrange: "L1:L500"
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
            stockUsageSpreadsheetIdCell: 'B15',
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
                serviceDueStart0: ColumnNames.letterToColumnStart0('S'),
                commentsStart0: ColumnNames.letterToColumnStart0('T')
            },
            referencesFirstCol: ColumnNames.letterToColumn('A'),
            referencesLastCol: ColumnNames.letterToColumn('T'),
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

function getEquipmentType(){
    return SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.equipmentTypeCell).getValue();
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
 * @returns {{isExportSubfolders: boolean, serviceRegisterId: *, exportFolder1: *, serviceregisterSheetName: *, equipmentOwner: *, serviceRegisterCols: {serviceType: *, unitNo: *, engineHours: *, comments: *, serviceDate: *, serviceDue: *}, equipmentType: *, exportFolder2: (*|null)}}
 */
function getReferences(equipmentOwner, equipmentType) {
    var allReferences = SPREADSHEET.sheets.references.sheet.getRange(
        SPREADSHEET.sheets.references.referencesFirstRow,
        SPREADSHEET.sheets.references.referencesFirstCol,
        SPREADSHEET.sheets.references.sheet.getLastRow(),
        SPREADSHEET.sheets.references.referencesLastCol - SPREADSHEET.sheets.references.referencesFirstCol + 1)
        .getValues();

    var equipmentOwnerColOffset = SPREADSHEET.sheets.references.equipmentOwnerColStart0;
    var equipmentTypeColOffset = SPREADSHEET.sheets.references.equipmentTypeColStart0;
    // We know this array has exactly one element
    var referenceArray = allReferences.filter(function (reference) {
        return equipmentOwner === reference[equipmentOwnerColOffset]
            && equipmentType === reference[equipmentTypeColOffset];
    });
    var referenceObj = referenceArray[0];
    return {
        equipmentOwner: referenceObj[equipmentOwnerColOffset],
        equipmentType: referenceObj[equipmentTypeColOffset],
        exportFolder1: folderUrlToId(referenceObj[SPREADSHEET.sheets.references.exportFolder1ColStart0]),
        exportFolder2: referenceObj[SPREADSHEET.sheets.references.exportFolder2ColStart0] !== 'N/A'
            ? folderUrlToId(referenceObj[SPREADSHEET.sheets.references.exportFolder2ColStart0])
            : null,
        isExportSubfolders: referenceObj[SPREADSHEET.sheets.references.isExportSubfoldersColStart0] === 'Y',
        serviceRegisterId: spreadsheetUrlToId(referenceObj[SPREADSHEET.sheets.references.serviceRegisterUrlColStart0]),
        serviceregisterSheetName: referenceObj[SPREADSHEET.sheets.references.serviceregisterSheetNameColStart0],
        serviceRegisterCols: {
            unitNo: referenceObj[SPREADSHEET.sheets.references.serviceRegisterCols.unitNoStart0],
            engineHours: referenceObj[SPREADSHEET.sheets.references.serviceRegisterCols.engineHoursStart0],
            serviceType: referenceObj[SPREADSHEET.sheets.references.serviceRegisterCols.serviceTypeStart0],
            serviceDate: referenceObj[SPREADSHEET.sheets.references.serviceRegisterCols.serviceDateStart0],
            serviceDue: referenceObj[SPREADSHEET.sheets.references.serviceRegisterCols.serviceDueStart0],
            comments: referenceObj[SPREADSHEET.sheets.references.serviceRegisterCols.commentsStart0],
        }
    };
}
