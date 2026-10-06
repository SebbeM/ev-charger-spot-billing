// ── Orchestration: fetch, price, and build all report sheets ─────────────────

var DETAILED_SHEET   = 'Detaljerad rapport';
var SIMPLIFIED_SHEET = 'Förenklad rapport';
var SESSIONS_SHEET   = 'Sessioner';
var SPOT_SHEET       = 'Spotpriser';

// Returns a summary message, shown in the login dialog when finished.
function buildSheets(token, fromMs, toMs, label) {
  var cfg      = getConfig();
  var sessions = fetchZaptecSessions(token, fromMs, toMs);

  // Prices are needed for the full time span of every session, including
  // sessions that started before the period or run past its end.
  var range  = sessionTimeRange(sessions, fromMs, toMs);
  var prices = fetchNordpoolPrices(range.from, range.to);

  var agg = aggregateSessions(sessions, prices.spotMap, cfg);

  buildDetailedReport(agg, label, cfg);
  buildSimplifiedReport(agg, label, cfg);
  buildSessionsSheet(agg, label, cfg);
  buildSpotPricesSheet(prices.spotMap, fromMs, toMs, label, cfg);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName(DETAILED_SHEET));

  return summaryMessage(agg, prices, label);
}

function summaryMessage(agg, prices, label) {
  var lines = [
    'Rapport för ' + label + ' är klar.',
    '',
    agg.sessionCount + ' sessioner, ' + prices.count + ' spotpriser hämtade.',
    agg.totalEnergy.toFixed(2) + ' kWh  →  ' + agg.totalCost.toFixed(2) + ' ' + CURRENCY + ' inkl. moms.'
  ];

  var warnings = [];
  if (prices.missingDays.length) {
    warnings.push('Spotpriser saknas för: ' + prices.missingDays.join(', ') + '.');
  }
  if (agg.unpricedKwh > 0.0005) {
    warnings.push(agg.unpricedKwh.toFixed(3) + ' kWh saknar spotpris och ingår inte i kostnaden. Kör rapporten igen senare.');
  }
  if (agg.flaggedCount) {
    warnings.push(agg.flaggedCount + ' sessioner har en notering. Se kolumnen Notering i fliken Sessioner.');
  }
  if (warnings.length) lines.push('', '⚠️ Kontrollera:', warnings.join('\n'));
  return lines.join('\n');
}
