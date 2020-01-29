var SERVICE_REGISTER_SPREADSHEET = {
    valuesRange: 'A3:AA47',
    hoursColumnOffset: 2,
    lastServiceCompletedColumnOffset: 3,
    dateLastServiceCompletedColumnOffset: 4,
    nextServiceDueColumnOffset: 5,
    partsRequiredForNextService: 6
};

function getServiceRegisterSpreadsheetId() {
    // TODO: dynamically get service register spreadsheet ID
}

function copyDataToServiceRegistry(){
    var serviceRegisterSpreadsheet = SpreadsheetApp.openById(getServiceRegisterSpreadsheetId());
    var serviceRegisterSheet = serviceRegisterSpreadsheet.getActiveSheet();
    var equipmentsNumbersRange = serviceRegisterSheet.getRange(SERVICE_REGISTER_SPREADSHEET.valuesRange);
    var equipmentsNumbersValues = equipmentsNumbersRange.getValues();
    var equipmentNumber = getEquipmentNumber();

    var equipmentRow = null;
    // Iterate through equipments in the service register and stop when
    for(var i=0; i < equipmentsNumbersValues.length; i++){
        if(equipmentsNumbersValues[i][0] === equipmentNumber)
            equipmentRow = equipmentsNumbersRange.offset(i, 0, 1);
    }
    if(equipmentRow !== null) {
        // The current equipment has been found in the service register
        var values = equipmentRow.getValues();
        if(serviceSheetIsServiceMode()){
            values[0][SERVICE_REGISTER_SPREADSHEET.hoursColumnOffset] = getMachineHours();
            values[0][SERVICE_REGISTER_SPREADSHEET.lastServiceCompletedColumnOffset] = getTaskType();
            values[0][SERVICE_REGISTER_SPREADSHEET.dateLastServiceCompletedColumnOffset] = getTaskDate();
            values[0][SERVICE_REGISTER_SPREADSHEET.nextServiceDueColumnOffset] = parseInt(getTaskType()) + 250;
        }
        values[0][SERVICE_REGISTER_SPREADSHEET.partsRequiredForNextService] = getComments();

        equipmentRow.setValues(values);
    }
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
