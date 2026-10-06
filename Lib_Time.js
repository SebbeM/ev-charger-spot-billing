// ── Time helpers ──────────────────────────────────────────────────────────────
// All internal timestamps are UTC milliseconds. Stockholm time is only used to
// define calendar periods and to display times. This keeps the code correct
// across daylight saving changes, where local clock times repeat or skip.

var TZ         = 'Europe/Stockholm';
var QUARTER_MS = 15 * 60 * 1000;
var HOUR_MS    = 60 * 60 * 1000;
var DAY_MS     = 24 * HOUR_MS;

var SWEDISH_MONTHS = ['januari', 'februari', 'mars', 'april', 'maj', 'juni',
                      'juli', 'augusti', 'september', 'oktober', 'november', 'december'];

// Zaptec sometimes returns timestamps without a zone ("2026-04-30T16:15:32.24").
// Those are UTC, but new Date() would read them as local time, so append 'Z'.
function parseUtc(s) {
  if (!s) return NaN;
  var str = String(s);
  var hasZone = /[zZ]$|[+-]\d\d:?\d\d$/.test(str);
  return new Date(hasZone ? str : str + 'Z').getTime();
}

// Start of the 15-minute block containing the given UTC timestamp.
function quarterKey(ms) {
  return Math.floor(ms / QUARTER_MS) * QUARTER_MS;
}

function formatLocal(dateOrMs, pattern) {
  return Utilities.formatDate(new Date(dateOrMs), TZ, pattern);
}

function formatLocalDateTime(ms) {
  return isNaN(ms) ? '' : formatLocal(ms, 'yyyy-MM-dd HH:mm');
}

// Current year and 0-based month in Stockholm.
function currentMonth() {
  var now = new Date();
  return { year: Number(formatLocal(now, 'yyyy')), month: Number(formatLocal(now, 'M')) - 1 };
}

// Stockholm midnight on the first day of a month. Month may be outside 0–11.
function monthStart(year, month) {
  var y = year + Math.floor(month / 12);
  var m = ((month % 12) + 12) % 12;
  return new Date(stockholmMidnight(y + '-' + pad2(m + 1) + '-01'));
}

function monthLabel(date) {
  return SWEDISH_MONTHS[Number(formatLocal(date, 'M')) - 1] + ' ' + formatLocal(date, 'yyyy');
}

// UTC milliseconds for 00:00 Stockholm time on a 'yyyy-MM-dd' date.
function stockholmMidnight(dateStr) {
  var p     = dateStr.split('-');
  var guess = Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  var ms    = guess - stockholmOffsetMs(guess);
  return guess - stockholmOffsetMs(ms); // re-check in case the offset differs at the result
}

// Stockholm's UTC offset at a given instant, in milliseconds (+1 h or +2 h).
function stockholmOffsetMs(ms) {
  var z    = formatLocal(ms, 'Z'); // e.g. "+0200"
  var sign = z.charAt(0) === '-' ? -1 : 1;
  return sign * (Number(z.substr(1, 2)) * HOUR_MS + Number(z.substr(3, 2)) * 60000);
}

function nextDate(dateStr) {
  var p = dateStr.split('-');
  var d = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]) + 1));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

// All Stockholm calendar days ('yyyy-MM-dd') that overlap [fromMs, toMs).
function stockholmDays(fromMs, toMs) {
  var days = [];
  var day  = formatLocal(fromMs, 'yyyy-MM-dd');
  while (stockholmMidnight(day) < toMs) {
    days.push(day);
    day = nextDate(day);
  }
  return days;
}

function pad2(n) {
  return (n < 10 ? '0' : '') + n;
}
