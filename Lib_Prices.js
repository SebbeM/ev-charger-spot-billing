// ── Spot prices: archive first, then Nord Pool, then elprisetjustnu.se ──────
// Every complete day that is fetched is saved in the hidden sheet Prisarkiv.
// Monthly reports keep the archive filled, so quarterly reports (whose first
// month is older than Nord Pool's free window) find their prices there.

var PRICE_ARCHIVE_SHEET = 'Prisarkiv';
var SOURCE_NORDPOOL     = 'Nord Pool';
var SOURCE_ELPRIS       = 'elprisetjustnu.se';

// Returns {
//   spotMap:      { quarterKeyMs: öre/kWh },
//   sourceMap:    { quarterKeyMs: source name },
//   missingDays:  days without a complete set of prices,
//   dayCounts:    { archive: n, 'Nord Pool': n, 'elprisetjustnu.se': n },
//   count:        number of quarters with a price
// }
function fetchSpotPrices(fromMs, toMs) {
  var days    = stockholmDays(fromMs, toMs);
  var archive = readPriceArchive();

  var needed   = days.filter(function(d) { return !archive[d]; });
  var nordpool = needed.length ? fetchNordpoolDays(needed) : {};
  var stillNeeded = needed.filter(function(d) { return !isCompleteDay(d, nordpool[d]); });
  var elpris   = stillNeeded.length ? fetchElprisDays(stillNeeded) : {};

  var result = { spotMap: {}, sourceMap: {}, missingDays: [], dayCounts: {}, count: 0 };
  var newArchiveRows = [];
  var count = function(key) { result.dayCounts[key] = (result.dayCounts[key] || 0) + 1; };

  days.forEach(function(day) {
    var prices, source;
    if (archive[day]) {
      prices = archive[day].prices; source = archive[day].source; count('archive');
    } else if (isCompleteDay(day, nordpool[day])) {
      prices = nordpool[day]; source = SOURCE_NORDPOOL; count(source);
    } else if (isCompleteDay(day, elpris[day])) {
      prices = elpris[day]; source = SOURCE_ELPRIS; count(source);
    } else {
      // Use whatever partial data exists; missing quarters show up as unpriced kWh.
      prices = nordpool[day] || elpris[day] || {};
      source = nordpool[day] ? SOURCE_NORDPOOL : SOURCE_ELPRIS;
      result.missingDays.push(day);
    }

    Object.keys(prices).forEach(function(k) {
      result.spotMap[k]   = prices[k];
      result.sourceMap[k] = source;
    });

    if (!archive[day] && isCompleteDay(day, prices)) {
      newArchiveRows.push([day, source, generatedAt(), JSON.stringify(prices)]);
    }
  });

  if (newArchiveRows.length) appendPriceArchive(newArchiveRows);
  result.count = Object.keys(result.spotMap).length;
  return result;
}

// A day is complete when it has a price for every quarter. Daylight saving days
// have 92 or 100 quarters instead of 96.
function isCompleteDay(day, prices) {
  if (!prices) return false;
  var expected = (stockholmMidnight(nextDate(day)) - stockholmMidnight(day)) / QUARTER_MS;
  return Object.keys(prices).length >= expected;
}

// ── Archive sheet: one row per day, prices stored as JSON ──────────────────────

function getPriceArchiveSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PRICE_ARCHIVE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PRICE_ARCHIVE_SHEET);
    sheet.getRange(1, 1, 1, 4)
      .setValues([['Datum', 'Källa', 'Hämtad', 'Priser (öre/kWh per kvart, JSON)']])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@'); // keep dates as plain text
    sheet.protect().setDescription('Prisarkiv — fylls automatiskt').setWarningOnly(true);
    sheet.hideSheet();
  }
  return sheet;
}

// Returns { day: { source, prices: { quarterKeyMs: öre/kWh } } }.
function readPriceArchive() {
  var sheet   = getPriceArchiveSheet();
  var archive = {};
  if (sheet.getLastRow() < 2) return archive;

  sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().forEach(function(row) {
    var day = row[0] instanceof Date ? Utilities.formatDate(row[0], TZ, 'yyyy-MM-dd') : String(row[0]);
    try {
      var prices = JSON.parse(row[3]);
      if (isCompleteDay(day, prices)) archive[day] = { source: row[1], prices: prices };
    } catch (err) {
      Logger.log('Skipping unreadable archive row for ' + day + ': ' + err);
    }
  });
  return archive;
}

function appendPriceArchive(rows) {
  var sheet = getPriceArchiveSheet();
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
}

// Human-readable summary of where the prices came from.
function priceSourceSummary(prices) {
  var parts = [];
  if (prices.dayCounts.archive)        parts.push(prices.dayCounts.archive + ' från arkivet');
  if (prices.dayCounts[SOURCE_NORDPOOL]) parts.push(prices.dayCounts[SOURCE_NORDPOOL] + ' från Nord Pool');
  if (prices.dayCounts[SOURCE_ELPRIS])   parts.push(prices.dayCounts[SOURCE_ELPRIS] + ' från elprisetjustnu.se');
  return parts.length ? 'Spotpriser per dag: ' + parts.join(', ') + '.' : '';
}

// Plain average of all quarter prices in [fromMs, toMs), öre/kWh excl. VAT.
// Returns null when there are no prices in the period.
function averageSpotPrice(spotMap, fromMs, toMs) {
  var sum = 0, n = 0;
  Object.keys(spotMap).map(Number).forEach(function(k) {
    if (k >= fromMs && k < toMs) { sum += spotMap[k]; n++; }
  });
  return n ? sum / n : null;
}
