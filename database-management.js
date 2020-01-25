var DATA_TYPE = 'Generator';

var DATABASE_SPREADSHEET = {
        getSpreadsheet: () => { return SpreadsheetApp.openById(DATABASE_SPREADSHEET_ID); },
        getSheet: () => { return DATABASE_SPREADSHEET.getSpreadsheet().getSheetByName("database"); },
        partsFirstCol: ColumnNames.letterToColumn('A'),
        partsLastCol: ColumnNames.letterToColumn('F'),
        partsFirstRow: '3'
    };

    /**
 * Export all parts in the database spreadsheet
 */
function exportPartsToDatabase() {
    var parts;
    if(serviceSheetIsServiceMode()) {
        parts = getPartsInServiceMode();
    }
    else if(serviceSheetIsRepairMode()) {
        parts = getPartsInRepairMode();
    }
    else  {
        // Do not export data to the database in repair mode
        return;
    }
    sendPartsToDatabase(parts);
}

function getPartsInServiceMode() {
    var beforeAdditionalPartsKey = 'Part no';
    var stopKey = 'Comments';
    var firstPartsRow = SPREADSHEET.sheets.serviceSheet.serviceMode.firstEntryRow;
    var nonFilteredParts = getPartsWithQuantityNonFiltered(firstPartsRow);
    var filteredParts = nonFilteredParts.filter(function (e) { return e[0] !== '' });

    var beforeAdditionalPartsIndex = nonFilteredParts.indexOf(beforeAdditionalPartsKey);
    var stopIndex = nonFilteredParts.indexOf(stopKey);
    var replaceParts =  nonFilteredParts.slice(0, beforeAdditionalPartsIndex);
    var additionalParts =  nonFilteredParts.slice(beforeAdditionalPartsIndex, stopIndex);

    var date = serviceSheet.getRange(SPREADSHEET.sheets.serviceSheet.taskDateCell);
    var equipmentNo = serviceSheet.getRange(SPREADSHEET.sheets.serviceSheet.equipmentNumberCell);
    var type = serviceSheet.getRange(SPREADSHEET.sheets.serviceSheet.typeCell);
    var task = serviceSheet.getRange(SPREADSHEET.sheets.serviceSheet.taskTypeCell);

    var transformedReplaceParts = replaceParts.map(function (e) {
        return [date, equipmentNo, type, task, e[0], 1];
    });

    var transformeAdditionalParts = additionalParts.map(function (e) {
        return [date, equipmentNo, type, task, e[0], e[1]];
    });

    return transformedReplaceParts.concat(transformeAdditionalParts);
}

function getPartsInRepairMode() {
    var firstPartsRow = SPREADSHEET.sheets.serviceSheet.serviceMode.firstEntryRow;
    var nonFilteredParts = getPartsWithQuantityNonFiltered(firstPartsRow);
    var filteredParts = nonFilteredParts.filter(function (e) { return e[0] !== '' });

    var serviceSheet = SPREADSHEET.sheets.serviceSheet.sheet;
    var date = serviceSheet.getRange(SPREADSHEET.sheets.serviceSheet.taskDateCell);
    var equipmentNo = serviceSheet.getRange(SPREADSHEET.sheets.serviceSheet.equipmentNumberCell);
    var type = serviceSheet.getRange(SPREADSHEET.sheets.serviceSheet.typeCell);
    var task = serviceSheet.getRange(SPREADSHEET.sheets.serviceSheet.taskTypeCell);

    var retVal = filteredParts.map(function (e) {
        return [date, equipmentNo, type, task, e[0], e[1]];
    });
    return retVal;
}

function sendPartsToDatabase(parts) {
    var dbSheet = DATABASE_SPREADSHEET.getSheet();
    var firstEmptyRow = getDatabaseFirstEmptyRow();
    var  insertRange = dbSheet.getRange(
        firstEmptyRow,
        DATABASE_SPREADSHEET.partsFirstCol,
        parts.length,
        DATABASE_SPREADSHEET.partsLastCol - DATABASE_SPREADSHEET.partsFirstCol + 1);
    insertRange.setValues(parts);
}

function getDatabaseFirstEmptyRow() {
    var dbSheet = DATABASE_SPREADSHEET.getSheet();
    return Math.min(dbSheet.getLastRow(), DATABASE_SPREADSHEET.partsFirstRow) + 1;
}

/**
 * Get a two dimensional array of values representing the parts and their quantities. It is not filtered (there should
 * be many empty lines at the bottom of the matrix)
 * @param firstrow
 * @returns {*}
 */
function getPartsWithQuantityNonFiltered(firstrow) {
    var range = SPREADSHEET.sheets.serviceSheet.getRange(
        firstrow,
        SPREADSHEET.sheets.serviceSheet.partsCol,
        SPREADSHEET.sheets.serviceSheet.getLastRow() + firstrow + 1,
        SPREADSHEET.sheets.serviceSheet.quantityCol - SPREADSHEET.sheets.serviceSheet.partsCol + 1);
    return range.getValues();
}