// ── Spot prices sheet: every 15-minute price in the report period ────────────

function buildSpotPricesSheet(spotMap, fromMs, toMs, label, cfg) {
  var sheet = resetSheet(SPOT_SHEET);
  var NCOL  = 3;

  writeTitle(sheet, 'Nord Pool 15-min priser — ' + NORDPOOL_AREA + ' — ' + label, NCOL);
  writeHeader(sheet, 2, ['Tid (svensk tid)', 'Spotpris (öre/kWh)', 'Totalt inkl. avgifter+moms (öre/kWh)']);

  var rows = Object.keys(spotMap).map(Number)
    .filter(function(k) { return k >= fromMs && k < toMs; })
    .sort(function(a, b) { return a - b; })
    .map(function(k) {
      var raw = spotMap[k];
      return [formatLocalDateTime(k), round(raw, 2), round((raw + cfg.TOTAL_MARKUP) * cfg.VAT_FACTOR, 2)];
    });

  writeRows(sheet, 3, rows, 9);
  setColumnWidths(sheet, [160, 160, 240]);
}
