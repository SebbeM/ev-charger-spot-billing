// ── Nord Pool day-ahead prices (15-minute resolution) ─────────────────────────
// dataportal-api.nordpoolgroup.com is Nord Pool's public data portal API. It is
// not formally documented, and days older than roughly two months return
// HTTP 401 (history requires a paid subscription). Lib_Prices.gs handles that
// by keeping an archive and falling back to elprisetjustnu.se.

var NORDPOOL_BATCH_SIZE = 40; // parallel requests per batch

// Fetches the given Stockholm days ('yyyy-MM-dd').
// Returns { day: { quarterKeyMs: öre/kWh, ... } } for days that returned data.
function fetchNordpoolDays(days) {
  return fetchDaysInBatches(days, NORDPOOL_BATCH_SIZE, function(day) {
    return 'https://dataportal-api.nordpoolgroup.com/api/DayAheadPrices'
         + '?date=' + day + '&market=DayAhead'
         + '&deliveryArea=' + NORDPOOL_AREA + '&currency=' + CURRENCY;
  }, function(body) {
    var prices = {};
    (body.multiAreaEntries || []).forEach(function(e) {
      var perMWh = (e.entryPerArea && e.entryPerArea[NORDPOOL_AREA] != null)
                 ? e.entryPerArea[NORDPOOL_AREA] : null;
      var start  = parseUtc(e.deliveryStart);
      if (perMWh == null || isNaN(start)) return;
      prices[quarterKey(start)] = perMWh / 10; // SEK/MWh → öre/kWh
    });
    return prices;
  }, 'Nord Pool');
}

// Shared by the price sources: fetches one URL per day in parallel batches,
// parses each JSON body, and keeps the days that produced at least one price.
function fetchDaysInBatches(days, batchSize, urlForDay, parseBody, sourceName) {
  var result = {};
  for (var b = 0; b < days.length; b += batchSize) {
    var batch    = days.slice(b, b + batchSize);
    var requests = batch.map(function(day) {
      return { url: urlForDay(day), muteHttpExceptions: true };
    });
    UrlFetchApp.fetchAll(requests).forEach(function(res, i) {
      var day = batch[i];
      if (res.getResponseCode() !== 200) {
        Logger.log(sourceName + ' HTTP ' + res.getResponseCode() + ' for ' + day);
        return;
      }
      try {
        var prices = parseBody(JSON.parse(res.getContentText()));
        if (Object.keys(prices).length) result[day] = prices;
      } catch (err) {
        Logger.log(sourceName + ' parse error ' + day + ': ' + err);
      }
    });
  }
  return result;
}

// Raw spot price (öre/kWh excl. VAT) for the quarter containing ms, or null.
function lookupSpotPrice(ms, spotMap) {
  var v = spotMap[quarterKey(ms)];
  return v != null ? v : null;
}
