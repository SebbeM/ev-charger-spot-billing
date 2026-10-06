// ── Nord Pool day-ahead prices (15-minute resolution) ─────────────────────────
// Note: dataportal-api.nordpoolgroup.com is Nord Pool's public data portal API.
// It is not formally documented and may change, so missing days are detected
// and reported instead of being ignored.

var NORDPOOL_BATCH_SIZE = 40; // parallel requests per batch

// Fetches all prices for the Stockholm days overlapping [fromMs, toMs).
// Returns { spotMap: {quarterKeyMs: öre/kWh}, missingDays: ['yyyy-MM-dd', ...], count }.
function fetchNordpoolPrices(fromMs, toMs) {
  var days        = stockholmDays(fromMs, toMs);
  var spotMap     = {};
  var missingDays = [];

  for (var b = 0; b < days.length; b += NORDPOOL_BATCH_SIZE) {
    var batch    = days.slice(b, b + NORDPOOL_BATCH_SIZE);
    var requests = batch.map(function(day) {
      return {
        url: 'https://dataportal-api.nordpoolgroup.com/api/DayAheadPrices'
           + '?date=' + day + '&market=DayAhead'
           + '&deliveryArea=' + NORDPOOL_AREA + '&currency=' + CURRENCY,
        muteHttpExceptions: true
      };
    });

    UrlFetchApp.fetchAll(requests).forEach(function(res, i) {
      var found = 0;
      if (res.getResponseCode() === 200) {
        try {
          (JSON.parse(res.getContentText()).multiAreaEntries || []).forEach(function(e) {
            var perMWh = (e.entryPerArea && e.entryPerArea[NORDPOOL_AREA] != null)
                       ? e.entryPerArea[NORDPOOL_AREA] : null;
            var start  = parseUtc(e.deliveryStart);
            if (perMWh == null || isNaN(start)) return;
            spotMap[quarterKey(start)] = perMWh / 10; // SEK/MWh → öre/kWh
            found++;
          });
        } catch (err) {
          Logger.log('Nord Pool parse error ' + batch[i] + ': ' + err);
        }
      } else {
        Logger.log('Nord Pool HTTP ' + res.getResponseCode() + ' for ' + batch[i]);
      }
      if (found === 0) missingDays.push(batch[i]);
    });
  }

  return { spotMap: spotMap, missingDays: missingDays, count: Object.keys(spotMap).length };
}

// Raw spot price (öre/kWh excl. VAT) for the quarter containing ms, or null.
function lookupSpotPrice(ms, spotMap) {
  var v = spotMap[quarterKey(ms)];
  return v != null ? v : null;
}
