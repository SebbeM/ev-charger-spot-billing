// ── Detailed report: one row per apartment ────────────────────────────────────

// periodAvgSpot: plain average of all spot prices in the period, öre/kWh excl. VAT.
function buildDetailedReport(agg, label, cfg, periodAvgSpot) {
  var sheet = resetSheet(DETAILED_SHEET);
  var NCOL  = 11;

  writeTitle(sheet, ORG_NAME + ' Laddboxar, förbrukning ' + label, NCOL);

  sheet.getRange(2, 1, 1, NCOL).setValues([[
    'Sessioner: ' + agg.sessionCount, '',
    'Total kWh: ' + agg.totalEnergy.toFixed(2), '',
    'Total kostnad: ' + agg.totalCost.toFixed(2) + ' ' + CURRENCY, '', '',
    'Med notering: ' + agg.flaggedCount, '',
    'Genererad: ' + generatedAt(), ''
  ]]).setBackground(COLOR_SUBTLE).setFontSize(9).setFontColor(COLOR_ENERGY);

  writeNote(sheet, 3,
    'Spotpris per 15 min + överföringsavgift ' + cfg.TRANSMISSION_FEE + ' + energiskatt ' + cfg.ENERGY_TAX
    + ' + påslag från leverantör ' + cfg.SUPPLIER_MARKUP + ' + påslag från förening ' + cfg.ASSOCIATION_MARKUP
    + ' öre/kWh, plus moms ' + vatPercent(cfg)
    + '  |  Område: ' + NORDPOOL_AREA
    + '  |  Perioden omfattar sessioner som avslutades under perioden  |  Redigera avgifter i fliken Inställningar',
    NCOL);

  writeHeader(sheet, 4, [
    'Lägenhetsnummer', 'Laddare-ID', 'Sessioner', 'Energi (kWh)', 'Snitt kWh/session', 'Timmar',
    'Kostnad (' + CURRENCY + ')', 'Snittpris (öre/kWh)', 'Ej prissatt (kWh)', 'Första session', 'Senaste session'
  ]);

  var flagged = {};
  var rows = agg.rows.map(function(r, i) {
    var starts = r.sessions.map(function(s) { return s.start; }).filter(function(t) { return !isNaN(t); });
    if (r.unpricedKwh > 0.0005) flagged[i] = true;
    return [
      r.name,
      r.id,
      r.sessions.length,
      round(r.energy, 3),
      round(r.sessions.length ? r.energy / r.sessions.length : 0, 3),
      round(r.duration / HOUR_MS, 1),
      round(r.cost, 2),
      r.pricedKwh > 0 ? round(r.cost / r.pricedKwh * 100, 1) : '',
      round(r.unpricedKwh, 3),
      starts.length ? formatLocalDateTime(Math.min.apply(null, starts)) : '',
      starts.length ? formatLocalDateTime(Math.max.apply(null, starts)) : ''
    ];
  });

  // Keep first/last session as plain text so Sheets does not parse some of
  // them into dates and leave others as strings.
  if (rows.length) sheet.getRange(5, 10, rows.length, 2).setNumberFormat('@');
  writeRows(sheet, 5, rows, 10, flagged);
  if (rows.length) {
    sheet.getRange(5, 4, rows.length, 1).setFontWeight('bold').setFontColor(COLOR_ENERGY);
    sheet.getRange(5, 7, rows.length, 1).setFontWeight('bold').setFontColor(COLOR_COST);
  }

  writeTotals(sheet, rows.length + 5, NCOL, [3, 4, 6, 7, 9], 5);

  // Explicit number formats for data and totals rows; otherwise Sheets may
  // infer a date format (0 shows as 1899-12-30).
  var formats = { 3: '0', 4: '0.000', 5: '0.000', 6: '0.0', 7: '0.00', 8: '0.0', 9: '0.000' };
  Object.keys(formats).forEach(function(col) {
    sheet.getRange(5, Number(col), rows.length + 1, 1).setNumberFormat(formats[col]);
  });
  writeBillCheck(sheet, rows.length + 7, NCOL, agg, cfg, periodAvgSpot);
  setColumnWidths(sheet, [150, 260, 80, 100, 120, 70, 110, 120, 110, 140, 140]);
}

// Figures to compare with the electricity bill, which covers the whole building
// and states consumption and an average price for spot + supplier markup.
function writeBillCheck(sheet, row, numCols, agg, cfg, periodAvgSpot) {
  var withMarkup = function(spot) { return spot == null ? '' : round(spot + cfg.SUPPLIER_MARKUP, 1); };
  // The bill shows the supplier's price, so the association's own markup is left out.
  var markup = ' + påslag från leverantör ' + cfg.SUPPLIER_MARKUP + ' öre/kWh';
  var lines = [
    ['Laddad energi (kWh)', round(agg.totalEnergy, 3), '0.000',
     'Laddboxarnas del av fastighetens förbrukning. Ska vara mindre än förbrukningen på elräkningen.'],
    ['Snittpris laddning (öre/kWh exkl. moms)', withMarkup(agg.avgSpot), '0.0',
     'Spotpris viktat efter när laddboxarna laddade' + markup + '. Lägre än periodens snitt om laddningen sker på billiga timmar.'],
    ['Snittpris perioden (öre/kWh exkl. moms)', withMarkup(periodAvgSpot), '0.0',
     'Medel av alla kvartspriser i perioden' + markup + '. Snittpriset på elräkningen bör ligga nära detta.']
  ];

  sheet.getRange(row, 1, 1, numCols).merge().setValue('Kontroll mot elräkning')
    .setFontWeight('bold').setBackground(COLOR_HEADER).setFontColor('#FFFFFF').setFontSize(10);
  lines.forEach(function(line, i) {
    var r = row + 1 + i;
    sheet.getRange(r, 1, 1, 3).merge().setValue(line[0]).setFontWeight('bold');
    sheet.getRange(r, 4).setValue(line[1]).setNumberFormat(line[2]).setFontWeight('bold').setFontColor(COLOR_ENERGY);
    sheet.getRange(r, 5, 1, numCols - 4).merge().setValue(line[3]).setFontColor('#6B7280').setWrap(true);
    sheet.getRange(r, 1, 1, numCols).setFontSize(9).setBackground(i % 2 === 0 ? '#FFFFFF' : COLOR_ROW_ALT);
  });
  writeNote(sheet, row + 1 + lines.length,
    'Elräkningen gäller kalendermånad, medan rapporten tar med laddningar som avslutades under perioden. Små avvikelser vid månadsskiftet är därför normala.',
    numCols);
}

function round(n, decimals) {
  var f = Math.pow(10, decimals);
  return Math.round(n * f) / f;
}
