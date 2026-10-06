// ── Settings sheet ────────────────────────────────────────────────────────────

var CONFIG_SHEET = 'Inställningar';

function getConfig() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) { sheet = ss.insertSheet(CONFIG_SHEET); writeConfigSheet(sheet); }

  var cfg = {};
  sheet.getDataRange().getValues().forEach(function(row) {
    if (row[0] && row[0] !== 'Inställning') cfg[row[0]] = row[1];
  });

  var transmissionFee = toNumber(cfg['transmission_fee'], 0);
  var energyTax       = toNumber(cfg['energy_tax'], 0);
  var markupOere      = toNumber(cfg['markup_oere'], 0);
  return {
    TRANSMISSION_FEE: transmissionFee,
    ENERGY_TAX:       energyTax,
    MARKUP_OERE:      markupOere,
    TOTAL_MARKUP:     transmissionFee + energyTax + markupOere, // öre/kWh excl. VAT
    VAT_FACTOR:       toNumber(cfg['vat_factor'], 1.25)
  };
}

// Accepts numbers as well as text such as "36,0". An empty cell gives the default,
// while an explicit 0 stays 0.
function toNumber(value, fallback) {
  if (typeof value === 'number') return value;
  var n = parseFloat(String(value == null ? '' : value).replace(',', '.'));
  return isNaN(n) ? fallback : n;
}

function writeConfigSheet(sheet) {
  sheet.clearContents(); sheet.clearFormats();
  sheet.getRange(1,1,1,3).merge()
    .setValue('Laddboxrapport — Inställningar')
    .setFontSize(13).setFontWeight('bold').setBackground('#1E3A5F').setFontColor('#FFFFFF');
  sheet.setRowHeight(1, 32);
  sheet.getRange(2,1,1,3).merge()
    .setValue('Redigera kolumnen Värde. Ändringarna gäller från nästa rapportkörning.')
    .setFontSize(9).setFontColor('#6B7280').setBackground('#F9FAFB');
  sheet.getRange(3,1,1,3).setValues([['Inställning', 'Värde', 'Beskrivning']])
    .setFontWeight('bold').setBackground('#EFF6FF').setFontColor('#1E3A8A');

  var rows = [
    ['transmission_fee', 0.0,  'Överföringsavgift i öre/kWh exkl. moms, enligt nätägarens faktura'],
    ['energy_tax',       36.0, 'Energiskatt i öre/kWh exkl. moms'],
    ['markup_oere',      0.0,  'Övrigt påslag i öre/kWh exkl. moms'],
    ['vat_factor',       1.25, '1,25 = 25 % moms  |  1,0 = ingen moms'],
  ];
  sheet.getRange(4,1,rows.length,3).setValues(rows);
  sheet.getRange(4,1,rows.length,1).setFontFamily('Courier New').setFontSize(10).setFontColor('#374151');
  sheet.getRange(4,2,rows.length,1)
    .setBackground('#EBF5FF').setFontSize(11).setFontWeight('bold').setFontColor('#1A56DB')
    .setBorder(true,true,true,true,false,false,'#93C5FD',SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(4,3,rows.length,1).setFontSize(9).setFontColor('#6B7280').setWrap(true);
  sheet.setColumnWidth(1,160); sheet.setColumnWidth(2,120); sheet.setColumnWidth(3,420);

  var linkRow = rows.length + 5;
  sheet.getRange(linkRow,1).setValue('ℹ️ Energiskatt — Skatteverket')
    .setFontSize(9).setFontColor('#6B7280');
  sheet.getRange(linkRow,2,1,2).merge()
    .setFormula('=HYPERLINK("https://www.skatteverket.se/foretag/skatterochavdrag/punktskatter/energiskatter/skattpael.4.15532c7b1442f256bae5e4c.html","Skatteverket — Skatt på elektrisk kraft")')
    .setFontColor('#1A56DB').setFontSize(9);

  var p = sheet.protect().setDescription('Inställningsstruktur');
  p.setUnprotectedRanges([sheet.getRange(4,2,rows.length,1)]);
  p.setWarningOnly(true);
}
