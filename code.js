 function onOpen(e) {
   SpreadsheetApp.getUi()
       .createMenu('CUSTOM OPTIONS')
       .addItem('Transfer Tables To Doc', 'createDoc')
       .addItem('Clear Tables', 'clearTables')
       .addToUi();
 }

function createDoc() {
  //get sheet information//
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];
  
  var facilityName = sheet.getRange("B1").getValue();
  var weekEnding1 = sheet.getRange("B2").getValue();
  var weekDate = Utilities.formatDate(weekEnding1, 'GMT+1', 'MM/dd/yyyy');
  var weekEnding2 = sheet.getRange("B13").getValue();
  var weekDate2 = Utilities.formatDate(weekEnding2, 'GMT+1', 'MM/dd/yyyy');
  var titleDate = Utilities.formatDate(weekEnding1, 'GMT+1','MM/dd');
  var titleDate2 = Utilities.formatDate(weekEnding2, 'GMT+1','MM/dd');
  var table1 = sheet.getRange("A3:E11").getValues();
  var table2 = sheet.getRange("A14:E22").getValues();
  //var sun = sheet.getRange("A2:E9").getValues();
  //place data in array
  
  
  //create doc
  var doc = DocumentApp.create('Flash Sheet '+ titleDate + ' to ' + titleDate2 );
  var body = doc.getBody().setMarginTop(.75);
  body.appendParagraph('HEALTHCARE SERVICES GROUP, INC. \nFLASH SHEET')
      .setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(14).setFontFamily("Times New Roman");
  body.appendParagraph('FACILITY: '+facilityName).editAsText().setFontSize(10);
  body.appendParagraph('WEEK ENDING: Sat '+ weekDate).editAsText().setFontSize(10);
  table = body.appendTable(table1);
  table.getRow(0).editAsText().setBold(true);
  body.appendParagraph('WEEK ENDING: Sat '+ weekDate2);
  table3 = body.appendTable(table2);
  table3.getRow(0).editAsText().setBold(true);
  body.appendParagraph('FAX EVERY MON BEFORE 10AM @ 562-494-8039').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(10);
  body.appendParagraph('OR \n if you\'re running late \n CALL with "ACTUAL WEEKLY TOTAL" @ 562-494-7939 \n Please do not fax with a COVER SHEET').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function clearTables(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];
  
  var table1 = sheet.getRange("C4:C10").setValue(0);
  var table2 = sheet.getRange("C15:C21").setValue(0);
}
