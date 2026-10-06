// ── Shared sheet formatting helpers ───────────────────────────────────────────

var COLOR_HEADER  = '#1E3A5F';
var COLOR_SUBTLE  = '#EFF6FF';
var COLOR_NOTE    = '#F9FAFB';
var COLOR_ROW_ALT = '#F8FAFC';
var COLOR_FLAG    = '#FEF3C7';
var COLOR_COST    = '#065F46';
var COLOR_ENERGY  = '#1E3A8A';

// Returns the named sheet emptied of content, formatting and merges.
function resetSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.setFrozenRows(0);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  return sheet;
}

function writeTitle(sheet, text, numCols) {
  sheet.getRange(1, 1, 1, numCols).merge()
    .setValue(text)
    .setFontSize(14).setFontWeight('bold').setBackground(COLOR_HEADER).setFontColor('#FFFFFF');
  sheet.setRowHeight(1, 36);
}

function writeNote(sheet, row, text, numCols) {
  sheet.getRange(row, 1, 1, numCols).merge()
    .setValue(text)
    .setFontSize(9).setFontColor('#6B7280').setBackground(COLOR_NOTE).setWrap(true);
}

function writeHeader(sheet, row, headers) {
  sheet.getRange(row, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground(COLOR_HEADER).setFontColor('#FFFFFF').setFontSize(10)
    .setWrap(true).setVerticalAlignment('middle');
  sheet.setFrozenRows(row);
}

// Writes rows with alternating backgrounds in one call. Rows whose index is in
// flaggedRows get a highlight colour instead.
function writeRows(sheet, startRow, rows, fontSize, flaggedRows) {
  if (!rows.length) return;
  var numCols = rows[0].length;
  var flagged = flaggedRows || {};
  var backgrounds = rows.map(function(_, i) {
    var color = flagged[i] ? COLOR_FLAG : (i % 2 === 0 ? '#FFFFFF' : COLOR_ROW_ALT);
    var line = [];
    for (var c = 0; c < numCols; c++) line.push(color);
    return line;
  });
  sheet.getRange(startRow, 1, rows.length, numCols)
    .setValues(rows).setBackgrounds(backgrounds).setFontSize(fontSize || 10);
}

function writeTotals(sheet, row, numCols, sumColumns, firstDataRow) {
  sheet.getRange(row, 1).setValue('TOTALT');
  sumColumns.forEach(function(col) {
    var l = String.fromCharCode(64 + col);
    sheet.getRange(row, col).setFormula('=SUM(' + l + firstDataRow + ':' + l + (row - 1) + ')');
  });
  sheet.getRange(row, 1, 1, numCols).setFontWeight('bold').setBackground(COLOR_SUBTLE);
}

function setColumnWidths(sheet, widths) {
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
}

function vatPercent(cfg) {
  return ((cfg.VAT_FACTOR - 1) * 100).toFixed(0) + ' %';
}

function generatedAt() {
  return formatLocal(new Date(), 'yyyy-MM-dd HH:mm');
}
