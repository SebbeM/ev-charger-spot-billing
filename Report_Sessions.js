// ── Sessions sheet: one row per charging session ──────────────────────────────

function buildSessionsSheet(agg, label, cfg) {
  var sheet = resetSheet(SESSIONS_SHEET);
  var NCOL  = 9;

  writeTitle(sheet, 'Alla sessioner — ' + label, NCOL);
  writeNote(sheet, 2, 'Tider i svensk tid. Gulmarkerade rader har en notering som bör kontrolleras.', NCOL);
  writeHeader(sheet, 3, [
    'Start', 'Slut', 'Lägenhetsnummer', 'Laddare-ID', 'kWh', 'Varaktighet (h)',
    'Snittpris (öre/kWh inkl. avgifter+moms)', 'Kostnad (' + CURRENCY + ')', 'Notering'
  ]);

  var flagged = {};
  var rows = agg.sessions.map(function(s, i) {
    var p = s.cost;
    if (p.notes.length) flagged[i] = true;
    return [
      formatLocalDateTime(s.start),
      formatLocalDateTime(s.end),
      apartmentName(s),
      s.chargerId,
      round(s.energy, 3),
      round(durationMs(s) / HOUR_MS, 2),
      p.avgPrice != null ? round(p.avgPrice, 2) : '',
      round(p.cost, 3),
      p.notes.join('; ')
    ];
  });

  writeRows(sheet, 4, rows, 9, flagged);
  if (rows.length) {
    sheet.getRange(4, 8, rows.length, 1).setFontColor(COLOR_COST);
    sheet.getRange(4, 9, rows.length, 1).setWrap(true);
  }
  setColumnWidths(sheet, [130, 130, 130, 260, 70, 90, 170, 100, 340]);
}
