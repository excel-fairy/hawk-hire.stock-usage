var SERVICE_REGISTER_SPREADSHEET = {
    servicesFirstRow: 4,
    nextServiceDueIncrement: 250
};

/**
 * Copy the service data to the service registry spreadsheet
 * @param equipmentReferences
 */
function copyDataToServiceRegistry(equipmentReferences){
    var serviceregisterSpreadSheetId = equipmentReferences.serviceRegisterId;
    var serviceRegisterSpreadsheet = SpreadsheetApp.openById(serviceregisterSpreadSheetId);
    var serviceRegisterSheet = serviceRegisterSpreadsheet.getSheetByName(equipmentReferences.serviceregisterSheetName);

    var unitNoCol = ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.unitNo) !== null
        ? ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.unitNo) : null;
    var engineHoursCol = ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.engineHours) !== null
        ? ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.engineHours) : null;
    var serviceTypeCol = ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.serviceType) !== null
        ? ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.serviceType) : null;
    var serviceDateCol = ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.serviceDate) !== null
        ? ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.serviceDate) : null;
    var nextServiceDueCol = ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.serviceDue) !== null
        ? ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.serviceDue) : null;
    var commentsCol = ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.comments) !== null
        ? ColumnNames.letterToColumn(equipmentReferences.serviceRegisterCols.comments) : null;

    var firstCol = unitNoCol;
    var lasttCol = commentsCol;

    var equipmentsNumbersRange = serviceRegisterSheet.getRange(
        SERVICE_REGISTER_SPREADSHEET.servicesFirstRow,
        firstCol,
        serviceRegisterSheet.getLastRow() - SERVICE_REGISTER_SPREADSHEET.servicesFirstRow,
        lasttCol - firstCol + 1
    );
    var equipmentsNumbersValues = equipmentsNumbersRange.getValues();
    var equipmentNumber = getEquipmentNumber();

    var equipmentRow = null;
    // Iterate through equipments in the service register and stop when the row of the right equipment is found
    for(var i=0; i < equipmentsNumbersValues.length; i++){
        if(equipmentsNumbersValues[i][unitNoCol - 1] === equipmentNumber)
            equipmentRow = equipmentsNumbersRange.offset(i, 0, 1);
    }
    if(equipmentRow !== null) {
        // The current equipment has been found in the service register
        var values = equipmentRow.getValues();
        if(serviceSheetIsServiceMode()){
            if(engineHoursCol !== null) {
                values[0][engineHoursCol - 1] = getMachineHours();
            }
            if(serviceTypeCol !== null) {
                values[0][serviceTypeCol - 1] = getTaskType();
            }
            if(serviceDateCol !== null) {
                values[0][serviceDateCol - 1] = getTaskDate();
            }
            if(nextServiceDueCol !== null) {
                values[0][nextServiceDueCol - 1] = parseInt(getTaskType())
                    + SERVICE_REGISTER_SPREADSHEET.nextServiceDueIncrement;
            }
        }
        if(commentsCol !== null) {
            values[0][commentsCol - 1] = getComments();
        }

        equipmentRow.setValues(values);
    }
}

/**
 * Get the comments of the service
 * @returns {string} The comments
 */
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
