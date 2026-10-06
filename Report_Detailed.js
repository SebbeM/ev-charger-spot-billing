// ── Detailed report: one row per apartment ────────────────────────────────────

function buildDetailedReport(agg, label, cfg) {
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
    + ' + påslag ' + cfg.MARKUP_OERE + ' öre/kWh, plus moms ' + vatPercent(cfg)
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

  writeRows(sheet, 5, rows, 10, flagged);
  if (rows.length) {
    sheet.getRange(5, 4, rows.length, 1).setFontWeight('bold').setFontColor(COLOR_ENERGY);
    sheet.getRange(5, 7, rows.length, 1).setFontWeight('bold').setFontColor(COLOR_COST);
  }

  writeTotals(sheet, rows.length + 5, NCOL, [3, 4, 6, 7, 9], 5);
  setColumnWidths(sheet, [150, 260, 80, 100, 120, 70, 110, 120, 110, 140, 140]);
}

function round(n, decimals) {
  var f = Math.pow(10, decimals);
  return Math.round(n * f) / f;
}
