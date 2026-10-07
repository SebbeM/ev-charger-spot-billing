// ── Detailed report: one row per apartment ────────────────────────────────────

function buildDetailedReport(agg, label, cfg, spotMap, fromMs, toMs) {
  var sheet = resetSheet(DETAILED_SHEET);
  var NCOL  = 12;

  writeTitle(sheet, ORG_NAME + ' Laddboxar, förbrukning ' + label, NCOL);

  sheet.getRange(2, 1, 1, NCOL).setValues([[
    'Sessioner: ' + agg.sessionCount, '',
    'Total kWh: ' + agg.totalEnergy.toFixed(2), '',
    'Total kostnad: ' + agg.totalCost.toFixed(2) + ' ' + CURRENCY, '', '',
    'Med notering: ' + agg.flaggedCount, '',
    'Genererad: ' + generatedAt(), '', ''
  ]]).setBackground(COLOR_SUBTLE).setFontSize(9).setFontColor(COLOR_ENERGY);

  writeNote(sheet, 3,
    'Spotpris per 15 min + överföringsavgift ' + cfg.TRANSMISSION_FEE + ' + energiskatt ' + cfg.ENERGY_TAX
    + ' + påslag från leverantör ' + cfg.SUPPLIER_MARKUP + ' + påslag från förening ' + cfg.ASSOCIATION_MARKUP
    + ' öre/kWh, plus moms ' + vatPercent(cfg)
    + '  |  Område: ' + NORDPOOL_AREA
    + '  |  Perioden omfattar sessioner som avslutades under perioden  |  Redigera avgifter i fliken Inställningar',
    NCOL);

  writeHeader(sheet, 4, [
    'Lägenhetsnummer', 'Mätare', 'Laddare-ID', 'Sessioner', 'Energi (kWh)', 'Snitt kWh/session', 'Timmar',
    'Kostnad (' + CURRENCY + ')', 'Snittpris (öre/kWh)', 'Ej prissatt (kWh)', 'Första session', 'Senaste session'
  ]);

  var flagged = {};
  var rows = agg.rows.map(function(r, i) {
    var starts = r.sessions.map(function(s) { return s.start; }).filter(function(t) { return !isNaN(t); });
    if (r.unpricedKwh > 0.0005) flagged[i] = true;
    return [
      r.name,
      r.meter,
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
  if (rows.length) sheet.getRange(5, 11, rows.length, 2).setNumberFormat('@');
  writeRows(sheet, 5, rows, 10, flagged);
  if (rows.length) {
    sheet.getRange(5, 5, rows.length, 1).setFontWeight('bold').setFontColor(COLOR_ENERGY);
    sheet.getRange(5, 8, rows.length, 1).setFontWeight('bold').setFontColor(COLOR_COST);
  }

  writeTotals(sheet, rows.length + 5, NCOL, [4, 5, 7, 8, 10], 5);

  // Explicit number formats for data and totals rows; otherwise Sheets may
  // infer a date format (0 shows as 1899-12-30).
  var formats = { 4: '0', 5: '0.000', 6: '0.000', 7: '0.0', 8: '0.00', 9: '0.0', 10: '0.000' };
  Object.keys(formats).forEach(function(col) {
    sheet.getRange(5, Number(col), rows.length + 1, 1).setNumberFormat(formats[col]);
  });
  writeBillCheck(sheet, rows.length + 7, NCOL, agg, cfg, spotMap, fromMs, toMs);
  setColumnWidths(sheet, [150, 90, 260, 80, 100, 120, 70, 110, 120, 110, 140, 140]);
}

// Figures to compare with the monthly electricity bills: one row per calendar
// month and meter. Each bill states the meter's consumption and an average
// price for spot + supplier markup.
function writeBillCheck(sheet, row, numCols, agg, cfg, spotMap, fromMs, toMs) {
  var months = stockholmMonths(fromMs, toMs);
  var meters = agg.meters.map(function(m) { return m.name; });

  // kWh per month and meter, counted by the quarter the energy was used in.
  // Sessions are attributed to the month they ended, but the bill is not.
  var cells = {};
  agg.sessions.forEach(function(s) {
    var meter = meterName(s);
    s.cost.quarters.forEach(function(q) {
      for (var i = 0; i < months.length; i++) {
        if (q.key < months[i].from || q.key >= months[i].to) continue;
        var key = i + '|' + meter;
        var c = cells[key] || (cells[key] = { kwh: 0, pricedKwh: 0, spotKwh: 0 });
        c.kwh += q.kwh;
        if (q.spot != null) { c.pricedKwh += q.kwh; c.spotKwh += q.kwh * q.spot; }
        return;
      }
    });
  });

  // The bill shows the supplier's price, so the association's own markup is left out.
  var withMarkup = function(spot) { return spot == null ? '' : round(spot + cfg.SUPPLIER_MARKUP, 1); };
  var rows = [], backgrounds = [];
  months.forEach(function(month, i) {
    var monthAvg = withMarkup(averageSpotPrice(spotMap, month.from, month.to));
    meters.forEach(function(meter) {
      var c = cells[i + '|' + meter] || { kwh: 0, pricedKwh: 0, spotKwh: 0 };
      rows.push([
        month.label + (month.partial ? ' (del av månad)' : ''),
        meter,
        round(c.kwh, 3),
        c.pricedKwh > 0 ? withMarkup(c.spotKwh / c.pricedKwh) : '',
        monthAvg
      ]);
      backgrounds.push(i % 2 === 0 ? '#FFFFFF' : COLOR_ROW_ALT);
    });
  });

  sheet.getRange(row, 1, 1, numCols).merge().setValue('Månadsstatistik per elmätare')
    .setFontWeight('bold').setBackground(COLOR_HEADER).setFontColor('#FFFFFF').setFontSize(10);
  sheet.getRange(row + 1, 1, 1, 5).setValues([[
    'Månad', 'Mätare', 'Laddad energi (kWh)', 'Snittpris laddning (öre/kWh)', 'Snittpris månaden (öre/kWh)'
  ]]).setFontWeight('bold').setBackground(COLOR_SUBTLE).setFontColor(COLOR_ENERGY).setFontSize(9).setWrap(true);

  var first = row + 2;
  if (rows.length) {
    // Plain text, or Sheets parses "september 2026" (an English month name) as a date
    sheet.getRange(first, 1, rows.length, 1).setNumberFormat('@');
    var range = sheet.getRange(first, 1, rows.length, 5).setValues(rows).setFontSize(9);
    range.setBackgrounds(backgrounds.map(function(b) { return [b, b, b, b, b]; }));
    sheet.getRange(first, 3, rows.length, 1).setNumberFormat('0.000');
    sheet.getRange(first, 4, rows.length, 2).setNumberFormat('0.0');
    sheet.getRange(first, 3, rows.length, 3).setFontWeight('bold').setFontColor(COLOR_ENERGY);
  }

  writeNote(sheet, first + rows.length,
    'Jämför varje rad med elräkningen för samma månad och mätare. Priserna är spotpris + påslag från leverantör '
    + cfg.SUPPLIER_MARKUP + ' öre/kWh, exkl. moms.\n'
    + 'Laddad energi: laddboxarnas del av mätarens förbrukning, räknad per kvart till den månad då den förbrukades. '
    + 'Ska vara mindre än förbrukningen på elräkningen. Laddningar som pågick vid periodens slut saknas i den sista månaden.\n'
    + 'Snittpris laddning: spotpris viktat efter när laddboxarna laddade. Lägre än månadens snitt om laddningen sker på billiga timmar.\n'
    + 'Snittpris månaden: medel av alla kvartspriser i månaden. Snittpriset på elräkningen bör ligga nära detta.',
    numCols);
  sheet.setRowHeight(first + rows.length, 70);
}

function round(n, decimals) {
  var f = Math.pow(10, decimals);
  return Math.round(n * f) / f;
}
