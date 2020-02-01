/**
 * Export the serivce sheet to:
 * - GDrive folder 1 as PDF
 * - GDrive folder 2 as PDF
 * - Email as attachement in PDF format
 * - Servide register spreadsheet
 * - Stock usage spreadsheet
 */
function exportServiceSheet() {
    var equipmentOwner = SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.equipmentOwnerCell);
    var equipmentType = SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.equipmentTypeCell);
    var equipmentReferences = getReferences(equipmentOwner, equipmentType);

    var equipmentNumber = getEquipmentNumber();
    var exportFolder1Id = getFolderToExportPdfTo(equipmentReferences.exportFolder1,
        equipmentReferences.isExportSubfolders, equipmentNumber).getId();
    var exportFolder2Id = getFolderToExportPdfTo(equipmentReferences.exportFolder2,
        equipmentReferences.isExportSubfolders, equipmentNumber).getId();

    var pdfFile = savePdfToDrive(exportFolder1Id);

    if(exportFolder2Id != null) {
        savePdfToDrive(exportFolder1Id);
    }

    sendEmail(pdfFile);
    exportPartsToStockUsageSheet();
    copyDataToServiceRegistry(equipmentReferences);
}

/**
 * Save the service sheet as PDF to the given GDrive folder
 * @param folderId The GDrive folder
 */
function savePdfToDrive(folderId) {
    var fileName = SPREADSHEET.sheets.service.sheet.getRange(SPREADSHEET.sheets.service.taskListNameCell).getValue();
    var exportOptions = {
        exportFolderId: folderId,
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
    return ExportSpreadsheet.export(exportOptions);
}

/**
 * Send an email with the exported PDF as attachment
 * @param attachment The exported PDF
 */
function sendEmail(attachment) {
    var copyRecipient = SPREADSHEET.sheets.emailAutomation
        .getRange(SPREADSHEET.sheets.emailAutomation.copyRecipientCell).getValue();
    var recipient = SPREADSHEET.sheets.emailAutomation
        .getRange(SPREADSHEET.sheets.emailAutomation.recipientCell).getValue();
    var subject = SPREADSHEET.sheets.emailAutomation
        .getRange(SPREADSHEET.sheets.emailAutomation.subjectCell).getValue();
    var message = SPREADSHEET.sheets.emailAutomation.getRange(SPREADSHEET.sheets.emailAutomation.bodyCell).getValue();
    var emailOptions = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: 'Automatic service sheet form mail sender',
        cc: copyRecipient
    };
    MailApp.sendEmail(recipient, subject, message, emailOptions);
}

/**
 * Get the ID of the folder to export the PDf file to
 * @param baseFolderId The base export folder
 * @param isExportSubfolders Should the PDF file be savec in a subfolder which name is the equipment number
 * @param equipmentNumber The equipment number
 * @returns {}
 */
function getFolderToExportPdfTo(baseFolderId, isExportSubfolders, equipmentNumber){
    var baseFolder = DriveApp.getFolderById(baseFolderId);
    if(!isExportSubfolders) {
        // PDF file should be exported straight in the base folder
        return baseFolder;
    } else {
        createExportFoldersIfNotExist(baseFolderId);
        var folders = baseFolder.getFolders();
        while (folders.hasNext()){
            var folder = folders.next();
            if(folder.getName() === equipmentNumber)
                return folder;
        }
        var otherFolder = baseFolder.getFoldersByName("Other");
        if(otherFolder.hasNext())
            return otherFolder.next();
        else
            return null;
    }
}

function createExportFoldersIfNotExist(baseFolderId){
    var range = SPREADSHEET.sheets.dataValidation.sheet.getRange(SPREADSHEET.sheets.dataValidation.equipmentsRange);
    var values = range.getDisplayValues();
    var baseFolder = DriveApp.getFolderById(baseFolderId);
    for(var i=0; i < values.length; i++){
        var folderName = values[i][0];
        if(folderName !== '' && !baseFolder.getFoldersByName(folderName).hasNext())
            baseFolder.createFolder(folderName);
    }
    baseFolder.createFolder("Other");
}
