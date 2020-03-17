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
    if(spreadSheetId != null) {
        return SpreadsheetApp.openById(spreadSheetId).getSheetByName(STOCK_USAGE_SPREADSHEET.sheetName);
    } else {
        return null;
    }
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
    var parts = getPartsWithQuantityNonFiltered(firstPartsRow);

    var beforeAdditionalPartsIndex = getIndexOfKeyIn2DArray(parts, beforeAdditionalPartsKey);
    var stopIndex = getIndexOfKeyIn2DArray(parts, stopKey);

    var retVal = [];

    var replaceParts = parts.slice(0, beforeAdditionalPartsIndex);
    var clientSuppliedPartsReplace = getClientSuppliedPart(replaceParts);
    if(!clientSuppliedPartsReplace) {
        retVal = retVal.concat(replaceParts)
    }

    var additionalParts = parts.slice(beforeAdditionalPartsIndex + 1, stopIndex);
    var clientSuppliedPartsAdditional = getClientSuppliedPart(additionalParts);
    if(!clientSuppliedPartsAdditional) {
        retVal = retVal.concat(additionalParts)
    }

    return buildServiceSheetRowToStockUsageRow(retVal);
}

function getPartsInRepairMode() {
    var firstPartsRow = SPREADSHEET.sheets.service.repairMode.firstEntryRow;
    var nonFilteredParts = getPartsWithQuantityNonFiltered(firstPartsRow);

    return buildServiceSheetRowToStockUsageRow(nonFilteredParts);
}

function buildServiceSheetRowToStockUsageRow(parts) {
    var filteredParts = filterPartsForExport(parts);

    var serviceSheet = SPREADSHEET.sheets.service.sheet;
    var date = serviceSheet.getRange(SPREADSHEET.sheets.service.taskDateCell).getValue();
    var equipmentNo = serviceSheet.getRange(SPREADSHEET.sheets.service.equipmentNumberCell).getValue();
    var type = getEquipmentType();
    var task = serviceSheet.getRange(SPREADSHEET.sheets.service.taskTypeCell).getValue();

    var retVal = filteredParts.map(function (e) {
        return [date, equipmentNo, type, task, e[2], e[3]];
    });

    var nbHours = getTotalNumberOfHoursOfTheJob(parts);
    if(nbHours) {
        // A total number of hours has been found in the parts list
        var totalNumberOfHours = [date, equipmentNo, type, task, 'Labour', nbHours];
        retVal.push(totalNumberOfHours);
    }

    return retVal;
}

function filterPartsForExport(parts) {
    return parts.filter(function (e) {
        // Row is an actual parts row, not a special row
        return e[0] !== SPREADSHEET.sheets.service.specialPartsCellsContents.clientSuppliedParts
            && e[0] !== SPREADSHEET.sheets.service.specialPartsCellsContents.totalNumberHoursOfJob
            // Part has a name
            && e[2] !== '';
    });
}

function getTotalNumberOfHoursOfTheJob(parts) {
    var totalNumberOfHoursRow = parts.filter(function (e) {
        return e[0] === SPREADSHEET.sheets.service.specialPartsCellsContents.totalNumberHoursOfJob;
    })[0]; // We know there is exactly one row that matches the filter
    if(!totalNumberOfHoursRow) {
        return undefined;
    }
    return totalNumberOfHoursRow[3];
}

function getClientSuppliedPart(parts) {
    var tickChar = "\u2714";

    var clientSuppliedPartsRow = parts.filter(function (e) {
        return e[0] === SPREADSHEET.sheets.service.specialPartsCellsContents.clientSuppliedParts;
    })[0]; // We know there is exactly one row that matches the filter
    if(!clientSuppliedPartsRow) {
        return false;
    }
    return clientSuppliedPartsRow[3] === tickChar;
}

function sendPartsToStockUsageSheet(parts) {
    if(parts.length > 0) {
        var dbSheet = getStockUsageSheet();
        if(dbSheet != null) {
            var firstEmptyRow = getStockUsageSheetFirstEmptyRow();
            var  insertRange = dbSheet.getRange(
                firstEmptyRow,
                STOCK_USAGE_SPREADSHEET.partsFirstCol,
                parts.length,
                STOCK_USAGE_SPREADSHEET.partsLastCol - STOCK_USAGE_SPREADSHEET.partsFirstCol + 1);
            insertRange.setValues(parts);
        }
    }
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
        SPREADSHEET.sheets.service.typeCol,
        SPREADSHEET.sheets.service.sheet.getLastRow() + firstrow + 1,
        SPREADSHEET.sheets.service.quantityCol - SPREADSHEET.sheets.service.typeCol + 1);
    return range.getValues();
}

/**
 * Get the index of the key (content of first column) of a 2D array
 */
function getIndexOfKeyIn2DArray(array, key) {
    for (var i = 0; i < array.length; i++) {
        if (array[i][2] === key) {
            return i;
        }
    }
    return -1;
}
