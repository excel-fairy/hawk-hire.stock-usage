var DATA_TYPE = 'Generator';

var STOCK_USAGE_SPREADSHEET = {
    partsFirstCol: ColumnNames.letterToColumn('A'),
    partsLastCol: ColumnNames.letterToColumn('F'),
    sheetName: 'Sheet1',
    partsFirstRow: '3'
};

function getStockUsageSheet() {
    var spreadsheetUrl = SPREADSHEET.sheets.references.sheet.getRange(
        SPREADSHEET.sheets.references.stockUsageSpreadsheetIdCell).getValue();
    var spreadSheetId = spreadsheetUrlToId(spreadsheetUrl);
    return SpreadsheetApp.openById(spreadSheetId).getSheetByName(STOCK_USAGE_SPREADSHEET.sheetName);
}

/**
 * Export all parts in the 'stock usage' spreadsheet
 */
function exportPartsToStockUsageSheet() {
    var parts;
    if(serviceSheetIsServiceMode()) {
        parts = getPartsInServiceMode();
    }
    else if(serviceSheetIsRepairMode()) {
        parts = getPartsInRepairMode();
    }
    else  {
        // Do not export data to the 'stock usage' sheet in repair mode
        return;
    }
    sendPartsToStockUsageSheet(parts);
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
        var partNoAndPartAmount = getReplacePartNo(e[0]);
        var partNo = partNoAndPartAmount.partNo;
        var partAmount = partNoAndPartAmount.partAmount;
        return [date, equipmentNo, type, task, partNo, partAmount];
    });

    var transformeAdditionalParts = additionalParts.map(function (e) {
        return [date, equipmentNo, type, task, e[0], e[1]];
    });

    return transformedReplaceParts.concat(transformeAdditionalParts);
}


/**
 * Extract the part number and amount of parts used from the replace "Part no" section
 * @param replacePartNo The combined part number and amount. It has the followinf formatting:
 * "[part no] x [amount of parts]"
 * @returns {{partAmount: *, partNo: *}} An object containing the part number and the amount of parts used
 */
function getReplacePartNo(replacePartNo) {
    var partNoRegex = /^(.*?) +x +(.*)$/
    var match = partNoRegex.exec(replacePartNo);
    // No idea why this is required. If not set, variable "match" will sometimes be null in the return line ...
    partNoRegex.exec(replacePartNo);

    var partNo = null;
    var partAmount = null;
    if(match !== null && match.length === 3) {
        // Both part no and part amount found
        partNo = match[1];
        partAmount = match[2]
    } else if(match == null) {
        // Separator not found. Assuming whole string is the part no
        partNo = 1;
        partAmount = replacePartNo
    } else {
        return null;
    }
    return {
        partNo: partNo,
        partAmount: partAmount
    }
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

function sendPartsToStockUsageSheet(parts) {
    var dbSheet = getStockUsageSheet();
    var firstEmptyRow = getStockUsageSheetFirstEmptyRow();
    var  insertRange = dbSheet.getRange(
        firstEmptyRow,
        STOCK_USAGE_SPREADSHEET.partsFirstCol,
        parts.length,
        STOCK_USAGE_SPREADSHEET.partsLastCol - STOCK_USAGE_SPREADSHEET.partsFirstCol + 1);
    insertRange.setValues(parts);
}

function getStockUsageSheetFirstEmptyRow() {
    var dbSheet = getStockUsageSheet();
    return Math.max(dbSheet.getLastRow(), STOCK_USAGE_SPREADSHEET.partsFirstRow) + 1;
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
