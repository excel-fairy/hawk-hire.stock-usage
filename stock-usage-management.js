var DATA_TYPE = 'Generator';

var DATABASE_SPREADSHEET = {
        partsFirstCol: ColumnNames.letterToColumn('A'),
        partsLastCol: ColumnNames.letterToColumn('F'),
        partsFirstRow: '3'
};

function getStockUsageSheet() {
    // TODO get from universal service sheet:references:B15
    // return SpreadsheetApp.openById(DATABASE_SPREADSHEET_ID).getSheetByName("Sheet1");
}

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
    var firstPartsRow = SPREADSHEET.sheets.service.serviceMode.firstEntryRow;
    var nonFilteredParts = getPartsWithQuantityNonFiltered(firstPartsRow);
    var filteredParts = nonFilteredParts.filter(function (e) { return e[0] !== '' });

    var beforeAdditionalPartsIndex = getIndexOfKeyIn2DArray(filteredParts, beforeAdditionalPartsKey);
    var stopIndex = getIndexOfKeyIn2DArray(filteredParts, stopKey);
    var replaceParts = filteredParts.slice(0, beforeAdditionalPartsIndex);
    var additionalParts = filteredParts.slice(beforeAdditionalPartsIndex + 1, stopIndex);

    var serviceSheet = SPREADSHEET.sheets.service.sheet;
    var date = serviceSheet.getRange(SPREADSHEET.sheets.service.taskDateCell).getValue();
    var equipmentNo = serviceSheet.getRange(SPREADSHEET.sheets.service.equipmentNumberCell).getValue();
    var type = DATA_TYPE;
    var task = serviceSheet.getRange(SPREADSHEET.sheets.service.taskTypeCell).getValue();

    var transformedReplaceParts = replaceParts.map(function (e) {
        return [date, equipmentNo, type, task, e[0], 1];
    });

    var transformeAdditionalParts = additionalParts.map(function (e) {
        return [date, equipmentNo, type, task, e[0], e[1]];
    });

    return transformedReplaceParts.concat(transformeAdditionalParts);
}

function getPartsInRepairMode() {
    var firstPartsRow = SPREADSHEET.sheets.service.repairMode.firstEntryRow;
    var nonFilteredParts = getPartsWithQuantityNonFiltered(firstPartsRow);
    var filteredParts = nonFilteredParts.filter(function (e) { return e[0] !== '' });

    var serviceSheet = SPREADSHEET.sheets.service.sheet;
    var date = serviceSheet.getRange(SPREADSHEET.sheets.service.taskDateCell).getValue();
    var equipmentNo = serviceSheet.getRange(SPREADSHEET.sheets.service.equipmentNumberCell).getValue();
    var type = DATA_TYPE;
    var task = serviceSheet.getRange(SPREADSHEET.sheets.service.taskTypeCell).getValue();

    var retVal = filteredParts.map(function (e) {
        return [date, equipmentNo, type, task, e[0], e[1]];
    });
    return retVal;
}

function sendPartsToDatabase(parts) {
    var dbSheet = getStockUsageSheet();
    var firstEmptyRow = getDatabaseFirstEmptyRow();
    var  insertRange = dbSheet.getRange(
        firstEmptyRow,
        DATABASE_SPREADSHEET.partsFirstCol,
        parts.length,
        DATABASE_SPREADSHEET.partsLastCol - DATABASE_SPREADSHEET.partsFirstCol + 1);
    insertRange.setValues(parts);
}

function getDatabaseFirstEmptyRow() {
    var dbSheet = getStockUsageSheet();
    return Math.max(dbSheet.getLastRow(), DATABASE_SPREADSHEET.partsFirstRow) + 1;
}

/**
 * Get a two dimensional array of values representing the parts and their quantities. It is not filtered (there should
 * be many empty lines at the bottom of the matrix)
 * @param firstrow
 * @returns {*}
 */
function getPartsWithQuantityNonFiltered(firstrow) {
    var range = SPREADSHEET.sheets.service.sheet.getRange(
        firstrow,
        SPREADSHEET.sheets.service.partsCol,
        SPREADSHEET.sheets.service.sheet.getLastRow() + firstrow + 1,
        SPREADSHEET.sheets.service.quantityCol - SPREADSHEET.sheets.service.partsCol + 1);
    return range.getValues();
}

/**
 * Get the index of the key (content of first column) of a 2D array
 */
function getIndexOfKeyIn2DArray(array, key) {
    for (var i = 0; i < array.length; i++) {
        if (array[i][0] === key) {
            return i;
        }
    }
    return -1;
}