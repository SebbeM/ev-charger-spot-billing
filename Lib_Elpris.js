// ── elprisetjustnu.se: fallback source with free price history ───────────────
// Free JSON API with 15-minute day-ahead prices per Swedish price area, in SEK
// and EUR. Used for days Nord Pool no longer serves for free. Its SEK prices
// are converted from EUR with the exchange rate given in each entry (EXR),
// so they can differ from Nord Pool's own SEK figures by a fraction of an öre.

var ELPRIS_BATCH_SIZE = 20;

// Returns { day: { quarterKeyMs: öre/kWh, ... } } for days that returned data.
function fetchElprisDays(days) {
  return fetchDaysInBatches(days, ELPRIS_BATCH_SIZE, function(day) {
    var p = day.split('-'); // yyyy-MM-dd → /yyyy/MM-dd_SE3.json
    return 'https://www.elprisetjustnu.se/api/v1/prices/' + p[0] + '/' + p[1] + '-' + p[2]
         + '_' + NORDPOOL_AREA + '.json';
  }, function(body) {
    var prices = {};
    (body || []).forEach(function(e) {
      var start   = parseUtc(e.time_start);
      var perKwh  = CURRENCY === 'EUR' ? e.EUR_per_kWh : e.SEK_per_kWh;
      if (perKwh == null || isNaN(start)) return;
      prices[quarterKey(start)] = perKwh * 100; // SEK/kWh → öre/kWh
    });
    return prices;
  }, 'elprisetjustnu.se');
}
