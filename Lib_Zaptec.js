// ── Zaptec API: authentication, sessions and session pricing ─────────────────
// Uses GET /api/sessions/archived, which replaces the deprecated
// /api/chargehistory (removed by 1 January 2027).

var ZAPTEC_PAGE_SIZE   = 200;
var ZAPTEC_MAX_PAGES   = 100;
var SPREAD_NOTE_KWH    = 0.5;            // flag evenly spread intervals above this energy…
var SPREAD_NOTE_MS     = 30 * 60 * 1000; // …that span longer than this

function authenticateZaptec(email, password) {
  var res = UrlFetchApp.fetch('https://api.zaptec.com/oauth/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: 'grant_type=password&username=' + encodeURIComponent(email)
           + '&password=' + encodeURIComponent(password),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Inloggningen misslyckades (HTTP ' + res.getResponseCode()
      + '). Kontrollera e-postadress och lösenord till Zaptec Portal.');
  }
  return JSON.parse(res.getContentText()).access_token;
}

// All completed sessions for the installation that END within [fromMs, toMs).
// The archived endpoint filters on end time, so a session that started before
// the period but ended inside it is included in this period.
function fetchZaptecSessions(token, fromMs, toMs) {
  var sessions = [];
  var cursor   = null;

  for (var page = 0; page < ZAPTEC_MAX_PAGES; page++) {
    var url = 'https://api.zaptec.com/api/sessions/archived'
            + '?InstallationId=' + encodeURIComponent(INSTALLATION_ID)
            + '&From=' + encodeURIComponent(new Date(fromMs).toISOString())
            + '&To='   + encodeURIComponent(new Date(toMs).toISOString())
            + '&PageSize=' + ZAPTEC_PAGE_SIZE
            + (cursor ? '&Cursor=' + encodeURIComponent(cursor) : '');

    var res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      throw new Error('Fel vid hämtning av sessioner från Zaptec (HTTP '
        + res.getResponseCode() + '): ' + res.getContentText().substring(0, 300));
    }

    var body = JSON.parse(res.getContentText());
    (body.sessions || []).forEach(function(raw) {
      var s = normalizeSession(raw);
      if (s) sessions.push(s);
    });

    if (!body.hasMore || !body.cursor) return sessions;
    cursor = body.cursor;
  }
  throw new Error('För många sidor med sessioner. Välj en kortare period.');
}

// Converts an API session to an internal shape with UTC millisecond times.
// Accepts both camelCase (archived API) and PascalCase (legacy API) fields.
function normalizeSession(raw) {
  function f(name) {
    var pascal = name.charAt(0).toUpperCase() + name.slice(1);
    return raw[name] !== undefined ? raw[name] : raw[pascal];
  }
  if (f('voided')) return null; // superseded by a corrected session

  var points = (f('energyDetails') || []).map(function(d) {
    return {
      t: parseUtc(d.timestamp !== undefined ? d.timestamp : d.Timestamp),
      v: Number(d.energy !== undefined ? d.energy : d.Energy) || 0
    };
  }).filter(function(p) { return !isNaN(p.t); })
    .sort(function(a, b) { return a.t - b.t; });

  var start = parseUtc(f('startDateTime'));
  var end   = parseUtc(f('endDateTime'));
  if (isNaN(end) && points.length) end = points[points.length - 1].t;

  return {
    id:            f('id'),
    chargerId:     f('chargerId') || 'okänd',
    deviceName:    f('deviceName') || '',
    start:         start,
    end:           end,
    energy:        Number(f('energy')) || 0,
    points:        points,
    offline:       f('offline') === true,
    reliableClock: f('reliableClock') !== false
  };
}

function apartmentName(session) {
  return CHARGER_NAMES[session.chargerId] || session.deviceName || session.chargerId;
}

function durationMs(session) {
  return (!isNaN(session.start) && !isNaN(session.end) && session.end > session.start)
    ? session.end - session.start : 0;
}

// Splits a session into energy intervals [{from, to, kwh}].
// Energy points may be cumulative (archived API: "energy delivered up to
// timestamp") or per-interval deltas (legacy API). The interpretation whose
// total matches the session energy best is used.
function buildIntervals(session) {
  var pts = session.points;
  var E   = session.energy;

  if (pts.length === 0) {
    return [{ from: session.start, to: session.end, kwh: E }];
  }

  var sum = 0, monotonic = true;
  for (var i = 0; i < pts.length; i++) {
    sum += pts[i].v;
    if (i > 0 && pts[i].v < pts[i - 1].v - 1e-6) monotonic = false;
  }
  var last        = pts[pts.length - 1].v;
  var cumulative  = monotonic && Math.abs(last - E) <= Math.abs(sum - E);

  var intervals = [];
  var firstFrom = (!isNaN(session.start) && session.start < pts[0].t) ? session.start : pts[0].t;
  intervals.push({ from: firstFrom, to: pts[0].t, kwh: pts[0].v }); // energy before first point

  for (var j = 1; j < pts.length; j++) {
    intervals.push({
      from: pts[j - 1].t,
      to:   pts[j].t,
      kwh:  cumulative ? pts[j].v - pts[j - 1].v : pts[j].v
    });
  }
  return intervals;
}

// Prices a session per 15-minute quarter.
// Each interval's energy is distributed over the quarters it overlaps, in
// proportion to the overlap. An interval's energy was consumed BEFORE its end
// timestamp, so it is never assigned to the quarter after the reading.
// Returns {cost, pricedKwh, unpricedKwh, spotKwhOere, avgPrice, quarters:[...], notes:[...]}.
// spotKwhOere is Σ kWh × spot over priced quarters, for consumption-weighted averages.
function priceSession(session, spotMap, cfg) {
  var notes     = [];
  var alloc     = {};   // quarterKey → kWh
  var spreadKwh = 0, spreadMs = 0;

  buildIntervals(session).forEach(function(iv) {
    if (!(iv.kwh > 0)) return;
    var span = iv.to - iv.from;
    if (isNaN(span) || span <= 0) {
      var k = quarterKey(isNaN(iv.to) ? iv.from : iv.to - 1);
      alloc[k] = (alloc[k] || 0) + iv.kwh;
      return;
    }
    if (iv.kwh >= SPREAD_NOTE_KWH && span > SPREAD_NOTE_MS) {
      spreadKwh += iv.kwh;
      spreadMs   = Math.max(spreadMs, span);
    }
    for (var q = quarterKey(iv.from); q < iv.to; q += QUARTER_MS) {
      var overlap = Math.min(iv.to, q + QUARTER_MS) - Math.max(iv.from, q);
      if (overlap > 0) alloc[q] = (alloc[q] || 0) + iv.kwh * overlap / span;
    }
  });

  // Bill the session's meter total; readings may differ slightly due to rounding.
  var allocated = 0;
  Object.keys(alloc).forEach(function(k) { allocated += alloc[k]; });
  var factor = allocated > 0 ? session.energy / allocated : 1;
  if (allocated > 0 && Math.abs(factor - 1) > 0.02) {
    notes.push('Mätpunkter summerar till ' + allocated.toFixed(3) + ' kWh, total ' + session.energy.toFixed(3) + ' kWh');
  }
  if (spreadKwh > 0) {
    notes.push(spreadKwh.toFixed(2) + ' kWh fördelad jämnt över upp till '
      + (spreadMs / HOUR_MS).toFixed(1) + ' h (mätpunkter saknas)');
  }
  if (session.offline) {
    notes.push(session.reliableClock ? 'Registrerad offline' : 'Registrerad offline, osäker klocka');
  }

  var cost = 0, pricedKwh = 0, unpricedKwh = 0, spotKwhOere = 0, quarters = [];
  Object.keys(alloc).map(Number).sort(function(a, b) { return a - b; }).forEach(function(q) {
    var kwh  = alloc[q] * factor;
    var spot = lookupSpotPrice(q, spotMap);
    var row  = { key: q, kwh: kwh, spot: spot, price: null, cost: 0 };
    if (spot == null) {
      unpricedKwh += kwh;
    } else {
      row.price  = (spot + cfg.TOTAL_MARKUP) * cfg.VAT_FACTOR; // öre/kWh incl. VAT
      row.cost   = kwh * row.price / 100;                       // SEK
      cost      += row.cost;
      pricedKwh += kwh;
      spotKwhOere += kwh * spot;
    }
    quarters.push(row);
  });

  if (unpricedKwh > 0.0005) {
    notes.push(unpricedKwh.toFixed(3) + ' kWh saknar spotpris och ingår inte i kostnaden');
  }

  return {
    cost:        cost,
    pricedKwh:   pricedKwh,
    unpricedKwh: unpricedKwh,
    spotKwhOere: spotKwhOere,
    avgPrice:    pricedKwh > 0 ? cost / pricedKwh * 100 : null,
    quarters:    quarters,
    notes:       notes
  };
}

// Prices all sessions and groups them per charger.
function aggregateSessions(sessions, spotMap, cfg) {
  var byCharger = {};
  var totals = { energy: 0, cost: 0, pricedKwh: 0, unpricedKwh: 0, spotKwhOere: 0, flagged: 0 };

  sessions.forEach(function(s) {
    var p  = priceSession(s, spotMap, cfg);
    s.cost = p;

    var c = byCharger[s.chargerId];
    if (!c) {
      c = byCharger[s.chargerId] = {
        id: s.chargerId, name: apartmentName(s), sessions: [],
        energy: 0, duration: 0, cost: 0, pricedKwh: 0, unpricedKwh: 0
      };
    }
    c.sessions.push(s);
    c.energy      += s.energy;
    c.duration    += durationMs(s);
    c.cost        += p.cost;
    c.pricedKwh   += p.pricedKwh;
    c.unpricedKwh += p.unpricedKwh;

    totals.energy      += s.energy;
    totals.cost        += p.cost;
    totals.pricedKwh   += p.pricedKwh;
    totals.unpricedKwh += p.unpricedKwh;
    totals.spotKwhOere += p.spotKwhOere;
    if (p.notes.length) totals.flagged++;
  });

  return {
    rows:         Object.keys(byCharger).map(function(k) { return byCharger[k]; })
                    .sort(function(a, b) { return b.energy - a.energy; }),
    sessions:     sessions.slice().sort(function(a, b) { return a.start - b.start; }),
    sessionCount: sessions.length,
    totalEnergy:  totals.energy,
    totalCost:    totals.cost,
    unpricedKwh:  totals.unpricedKwh,
    // Spot price weighted by charged kWh, öre/kWh excl. VAT and fees
    avgSpot:      totals.pricedKwh > 0 ? totals.spotKwhOere / totals.pricedKwh : null,
    flaggedCount: totals.flagged
  };
}

// Earliest start and latest end over all sessions, for fetching prices.
function sessionTimeRange(sessions, fromMs, toMs) {
  var lo = fromMs, hi = toMs;
  sessions.forEach(function(s) {
    [s.start, s.end].concat(s.points.length ? [s.points[0].t] : []).forEach(function(t) {
      if (!isNaN(t)) { lo = Math.min(lo, t); hi = Math.max(hi, t); }
    });
  });
  return { from: lo, to: hi + 1 };
}
