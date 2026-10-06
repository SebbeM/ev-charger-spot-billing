// ── Debug: show the quarter-by-quarter pricing of the longest session ────────
// Uses exactly the same priceSession() as the reports, so what you see here is
// what the reports bill.

var DEBUG_SHEET = 'Debug';

function runDebug(token, fromMs, toMs) {
  var cfg      = getConfig();
  var sessions = fetchZaptecSessions(token, fromMs, toMs).filter(function(s) { return s.energy > 1; });
  if (!sessions.length) throw new Error('Hittade ingen session över 1 kWh under perioden.');

  var session = sessions.reduce(function(a, b) { return durationMs(b) > durationMs(a) ? b : a; });
  var range   = sessionTimeRange([session], session.start, session.end);
  var prices  = fetchSpotPrices(range.from, range.to);
  var p       = priceSession(session, prices.spotMap, cfg);

  var sheet = resetSheet(DEBUG_SHEET);
  var NCOL  = 6;
  writeTitle(sheet, 'Debug — prissättning per kvart', NCOL);

  var info = [
    ['Session ID',            session.id],
    ['Lägenhet',              apartmentName(session)],
    ['Start (svensk tid)',    formatLocalDateTime(session.start)],
    ['Slut (svensk tid)',     formatLocalDateTime(session.end)],
    ['Total kWh (mätare)',    session.energy],
    ['Antal mätpunkter',      session.points.length],
    ['Påslag exkl. moms',     cfg.TOTAL_MARKUP + ' öre/kWh'],
    ['Momsfaktor',            cfg.VAT_FACTOR],
    ['Kostnad',               round(p.cost, 2) + ' ' + CURRENCY],
    ['Noteringar',            p.notes.join('; ') || '—']
  ];
  sheet.getRange(2, 1, info.length, 2).setValues(info);
  sheet.getRange(2, 1, info.length, 1).setFontWeight('bold').setBackground(COLOR_SUBTLE);

  var hdrRow = info.length + 3;
  writeHeader(sheet, hdrRow, [
    'Kvart (svensk tid)', 'kWh i kvarten', 'Spotpris (öre/kWh)',
    'Totalpris inkl. avgifter+moms (öre/kWh)', 'Kostnad (' + CURRENCY + ')', 'Notering'
  ]);

  var flagged = {};
  var rows = p.quarters.map(function(q, i) {
    if (q.spot == null) flagged[i] = true;
    return [
      formatLocal(q.key, 'yyyy-MM-dd HH:mm') + '–' + formatLocal(q.key + QUARTER_MS, 'HH:mm'),
      round(q.kwh, 4),
      q.spot != null ? round(q.spot, 2) : '—',
      q.price != null ? round(q.price, 2) : '—',
      round(q.cost, 4),
      q.spot == null ? 'Inget spotpris' : ''
    ];
  });
  writeRows(sheet, hdrRow + 1, rows, 9, flagged);
  writeTotals(sheet, hdrRow + 1 + rows.length, NCOL, [2, 5], hdrRow + 1);
  setColumnWidths(sheet, [190, 120, 130, 220, 120, 160]);

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  return 'Debug klar för ' + apartmentName(session) + ', ' + formatLocalDateTime(session.start)
       + '.\n\n' + p.quarters.length + ' kvartar, ' + session.energy.toFixed(3) + ' kWh, '
       + p.cost.toFixed(2) + ' ' + CURRENCY + '.\n\nSe fliken Debug.';
}
