// todo
function exportToPdf() { /* ici aussi j'imagine qu'il va falloir changer puisque maintenant il y a plusieurs folder export en fonction de ce qui est selectionne dans la service sheet*/
    var equipmentNumber = getEquipmentNumber();
    var exportFolderId = getFolderToExportPdfTo(equipmentNumber).getId();
    var fileName = SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.taskListNameCell).getValue();

    var exportOptions = {
        exportFolderId: exportFolderId,
        exportFileName: fileName,
        range: {
            r1: SPREADSHEET.sheets.service.taskListCoordinates.fullDocumentBeginningRow - 1,
            r2: SPREADSHEET.sheets.service.taskListCoordinates.row + getNbTasks(),
            c1: SPREADSHEET.sheets.service.taskListCoordinates.col - 1,
            c2: SPREADSHEET.sheets.service.taskListCoordinates.col + SPREADSHEET.sheets.service.taskListCoordinates.nbCols - 1
        },
        repeatHeader: true,
        fileFormat: 'pdf'
    };
    var pdfFile = ExportSpreadsheet.export(exportOptions);
    sendEmail(pdfFile);
    exportPartsToDatabase();
    copyDataToServiceRegistry();
}

// todo
function sendEmail(attachment) { /* Ici est ce qu'on pourrait rajouter une email addresse en copie? C'est l'addresse qui est dans l'onglet email automation B9*/
    var recipient = SPREADSHEET.sheets.emailAutomation.getRange("B8").getValue();
    var subject = SPREADSHEET.sheets.emailAutomation.getRange("B10").getValue();
    var message = SPREADSHEET.sheets.emailAutomation.getRange("B11").getValue();
    var emailOptions = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: 'Automatic service sheet form mail sender'
    };
    MailApp.sendEmail(recipient, subject, message, emailOptions);
}

function getFolderToExportPdfTo(folderName){
    createExportFoldersIfNotExist();
    var parentFolder = DriveApp.getFolderById(getExportFolderId());
    var folders = parentFolder.getFolders();
    while (folders.hasNext()){
        var folder = folders.next();
        if(folder.getName() === folderName)
            return folder;
    }
    var otherFolder = parentFolder.getFoldersByName("Other");
    if(otherFolder.hasNext())
        return otherFolder.next();
    else
        return null;
}

function createExportFoldersIfNotExist(){
    var range = SPREADSHEET.sheets.dataValidation.sheet.getRange(SPREADSHEET.sheets.dataValidation.equipmentsRange);
    var values = range.getDisplayValues();
    var parentFolder = DriveApp.getFolderById(getExportFolderId());
    for(var i=0; i < values.length; i++){
        var folderName = values[i][0];
        if(folderName !== '' && !parentFolder.getFoldersByName(folderName).hasNext())
            parentFolder.createFolder(folderName);
    }
    parentFolder.createFolder("Other");
}

function getExportFolderId(){
    // TODO: dynamically get export folder ID
}