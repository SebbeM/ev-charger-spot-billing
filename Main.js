// ─── Static configuration ─────────────────────────────────────────────────────
var NORDPOOL_AREA   = 'SE3';
var CURRENCY        = 'SEK';
// ─────────────────────────────────────────────────────────────────────────────

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Instruktioner')) buildInstructionsSheet();
  SpreadsheetApp.getUi()
    .createMenu('Beräkna')
    .addItem('Hämta förra månadens rapport', 'fetchLastMonth')
    .addItem('Hämta senaste kvartalets rapport', 'fetchLastQuarter')
    .addItem('Hämta rapport för valfri period', 'fetchCustomPeriod')
    .addSeparator()
    .addItem('Debug — analysera längsta session', 'debugLongSession')
    .addToUi();
}

// ── Menu handlers: compute the period, then open the login dialog ────────────
// All periods are calendar boundaries in Stockholm time, independent of the
// script's time zone setting.

function fetchLastMonth() {
  var cur  = currentMonth();
  var from = monthStart(cur.year, cur.month - 1);
  var to   = monthStart(cur.year, cur.month);
  openLoginDialog({ action: 'report', from: from.getTime(), to: to.getTime(), label: monthLabel(from) });
}

function fetchLastQuarter() {
  var cur        = currentMonth();
  var startMonth = Math.floor(cur.month / 3) * 3 - 3; // may be negative; monthStart rolls back the year
  var from       = monthStart(cur.year, startMonth);
  var to         = monthStart(cur.year, startMonth + 3);
  var quarter    = Math.floor((Number(formatLocal(from, 'M')) - 1) / 3) + 1;
  openLoginDialog({
    action: 'report',
    from:   from.getTime(),
    to:     to.getTime(),
    label:  'Q' + quarter + ' ' + formatLocal(from, 'yyyy')
  });
}

function fetchCustomPeriod() {
  openLoginDialog({ action: 'report', custom: true });
}

function debugLongSession() {
  var cur  = currentMonth();
  var from = monthStart(cur.year, cur.month - 1);
  var to   = monthStart(cur.year, cur.month);
  openLoginDialog({ action: 'debug', from: from.getTime(), to: to.getTime(), label: monthLabel(from) });
}

// ── Login dialog ──────────────────────────────────────────────────────────────

function openLoginDialog(ctx) {
  ctx.email = PropertiesService.getUserProperties().getProperty('zaptecEmail') || '';

  var template = HtmlService.createTemplateFromFile('LoginDialog');
  template.ctxJson = JSON.stringify(ctx).replace(/</g, '\\u003c');

  var title = ctx.custom ? 'Rapport för valfri period'
            : ctx.action === 'debug' ? 'Debug — ' + ctx.label
            : 'Rapport — ' + ctx.label;

  var html = template.evaluate().setWidth(400).setHeight(ctx.custom ? 400 : 320);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

// Called from LoginDialog via google.script.run.
// Returns a summary message for the dialog; throws an Error to show a failure.
function handleLogin(form) {
  var email = String(form.email || '').trim();
  if (!email || !form.password) throw new Error('Ange e-postadress och lösenord.');

  var from  = form.from;
  var to    = form.to;
  var label = form.label;
  if (form.custom) {
    var period = parseCustomPeriod(form.fromDate, form.toDate);
    from  = period.from;
    to    = period.to;
    label = period.label;
  }

  var token = authenticateZaptec(email, form.password);
  // Only the e-mail address is remembered, never the password.
  PropertiesService.getUserProperties().setProperty('zaptecEmail', email);

  return form.action === 'debug'
    ? runDebug(token, from, to)
    : buildSheets(token, from, to, label);
}

function parseCustomPeriod(fromStr, toStr) {
  var re = /^\d{4}-\d{2}-\d{2}$/;
  if (!re.test(fromStr || '') || !re.test(toStr || '')) {
    throw new Error('Ange datum i formatet ÅÅÅÅ-MM-DD.');
  }
  if (toStr < fromStr) {
    throw new Error('Slutdatum måste vara samma dag som eller efter startdatum.');
  }
  return {
    from:  stockholmMidnight(fromStr),
    to:    stockholmMidnight(nextDate(toStr)), // end date is inclusive
    label: fromStr + ' → ' + toStr
  };
}
