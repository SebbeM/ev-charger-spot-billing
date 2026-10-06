function buildStartSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Start') || ss.insertSheet('Start', 0);
  sheet.clearContents(); sheet.clearFormats();

  sheet.getRange(1,1,1,2).merge()
    .setValue(ORG_NAME + ' — Laddboxrapport')
    .setFontSize(20).setFontWeight('bold').setBackground('#1E3A5F').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 60);

  sheet.getRange(3,1).setValue('📖 Instruktioner:').setFontSize(12).setFontWeight('bold');
  sheet.getRange(3,2)
    .setFormula('=HYPERLINK("' + INSTRUCTIONS_URL + '","Öppna instruktionsdokumentet")')
    .setFontSize(12).setFontColor('#1A56DB');

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 500);
  sheet.setHiddenGridlines(true);
}