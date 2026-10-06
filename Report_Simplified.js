// ── Simplified report for the accountant ──────────────────────────────────────

function buildSimplifiedReport(agg, label, cfg) {
  var sheet = resetSheet(SIMPLIFIED_SHEET);
  var NCOL  = 3;

  writeTitle(sheet, ORG_NAME + ' Laddboxar, förbrukning ' + label, NCOL);

  var note = 'Belopp exkl. moms. Genererad: ' + generatedAt();
  if (agg.unpricedKwh > 0.0005) {
    note += '\nOBS: ' + agg.unpricedKwh.toFixed(3) + ' kWh saknar spotpris och ingår inte i beloppen.';
  }
  writeNote(sheet, 2, note, NCOL);

  writeHeader(sheet, 3, ['Lägenhetsnummer', 'Energi (kWh)', 'Kostnad exkl. moms (' + CURRENCY + ')']);

  var rows = agg.rows.map(function(r) {
    return [r.name, round(r.energy, 3), round(r.cost / cfg.VAT_FACTOR, 2)];
  });
  writeRows(sheet, 4, rows, 10);
  if (rows.length) sheet.getRange(4, 3, rows.length, 1).setFontWeight('bold').setFontColor(COLOR_COST);

  writeTotals(sheet, rows.length + 4, NCOL, [2, 3], 4);
  setColumnWidths(sheet, [200, 130, 200]);
}
