var SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Valid");
var FOLDER_NAMES_RANGE = 'B3:B36';

function createExportFolders(){
    var range = SHEET.getRange(FOLDER_NAMES_RANGE);
    var values = range.getDisplayValues();
    var parentFolder = DriveApp.getFolderById(EXPORT_FOLDER_ID);
    for(var i=0; i < values.length; i++){
        var folderName = values[i][0];
        if(folderName !== '' && !parentFolder.getFoldersByName(folderName).hasNext())
            parentFolder.createFolder(folderName);
    }
    parentFolder.createFolder("Other");
}