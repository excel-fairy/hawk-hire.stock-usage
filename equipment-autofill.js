function onEdit() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetEquipAutofill = ss.getSheetByName("equipment autofill");
  var celltype = sheetEquipAutofill.getRange("B3").getValue();
  var cellenginemake = sheetEquipAutofill.getRange("D3").getValue();
  var cellenginemodel = sheetEquipAutofill.getRange("B4").getValue();
  var cellengineserialno = sheetEquipAutofill.getRange("D4").getValue();
  var cellengineyear = sheetEquipAutofill.getRange("B5").getValue();
  var cellunitserialno = sheetEquipAutofill.getRange("D5").getValue();

  var sheetservicesheet = ss.getSheetByName("Service sheet");

  sheetservicesheet.getRange("C7").setValue(celltype);
  sheetservicesheet.getRange("E7").setValue(cellenginemake);
  sheetservicesheet.getRange("C8").setValue(cellenginemodel);
  sheetservicesheet.getRange("E8").setValue(cellengineserialno);
  sheetservicesheet.getRange("C9").setValue(cellengineyear);
  sheetservicesheet.getRange("E9").setValue(cellunitserialno);


}