// Quiz sezonu 2026: serwer w Apps Script (arkusz "Quiz 2026 – wyniki").
// Karty arkusza: Sterowanie, Pytania, Kategorie, Odpowiedzi, Odpowiedzi live, Ranking.

var LIVE = 'Odpowiedzi live';
var NAMES = 'Pseudonimy';   // rezerwacje pseudonimów: Gra, Klucz, Pseudonim, Kod urządzenia, Czas

// Klucz pseudonimu: bez rozróżniania wielkości liter i nadmiarowych spacji.
function nkey_(n) { return String(n).trim().toLowerCase().replace(/\s+/g, ' '); }

// Rezerwacja pseudonimu w danej grze ('free' w trybie dowolnym, id gry na żywo).
// Ten sam kod urządzenia może wracać; inny kod dostaje taken=1.
function claim_(scope, name, tok) {
  var key = nkey_(name);
  if (!key) return { ok: 0 };
  var lock = LockService.getScriptLock(); lock.waitLock(8000);
  try {
    var s = sh_(NAMES);
    if (!s) { s = SpreadsheetApp.getActive().insertSheet(NAMES); s.appendRow(['Gra', 'Klucz', 'Pseudonim', 'Kod urządzenia', 'Czas']); s.setFrozenRows(1); }
    var last = s.getLastRow();
    if (last >= 2) {
      var rows = s.getRange(2, 1, last - 1, 4).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]) === scope && String(rows[i][1]) === key) {
          return String(rows[i][3]) === String(tok) ? { ok: 1 } : { ok: 0, taken: 1 };
        }
      }
    }
    s.appendRow([scope, key, String(name).trim().slice(0, 20), String(tok), new Date()]);
    return { ok: 1 };
  } finally { lock.releaseLock(); }
}

function sh_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function cfg_() {
  var rows = sh_('Sterowanie').getRange('A2:B12').getValues(), c = {};
  rows.forEach(function (r) { if (r[0] !== '') c[String(r[0]).trim()] = r[1]; });
  var open = c['Odblokowane pytania (tryb dowolny)'];
  return {
    mode: c['Tryb gry'] === 'Na żywo' ? 'live' : 'free',
    category: String(c['Kategoria do gry'] || 'Wszystkie'),
    time: Number(c['Czas na pytanie (s)']) || 20,
    phone: String(c['Telefon w trybie na żywo pokazuje'] || '').indexOf('Tylko') === 0 ? 'buttons' : 'full',
    open: (open === '' || String(open).toLowerCase() === 'wszystkie') ? 999 : Number(open) || 0,
    pin: String(c['PIN prowadzącego'] === undefined ? '' : c['PIN prowadzącego'])
  };
}

// Pytania aktywne, z wybranej kategorii. id = numer wiersza w arkuszu.
function questions_(category) {
  var s = sh_('Pytania'), last = s.getLastRow();
  if (last < 2) return [];
  var rows = s.getRange(2, 1, last - 1, 8).getValues(), out = [];
  rows.forEach(function (r, k) {
    if (!r[1] || String(r[7]).toUpperCase() !== 'TAK') return;
    if (category && category !== 'Wszystkie' && String(r[0]) !== category) return;
    var c = 'ABCD'.indexOf(String(r[6]).trim().toUpperCase());
    if (c < 0) return;
    out.push({ id: k + 2, cat: String(r[0]), q: String(r[1]).trim(), o: [r[2], r[3], r[4], r[5]].map(String), c: c });
  });
  return out;
}

function questionById_(id) {
  var all = questions_('Wszystkie'), i;
  for (i = 0; i < all.length; i++) if (all[i].id === id) return all[i];
  return null;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function getState_() {
  var p = PropertiesService.getScriptProperties().getProperty('state');
  return p ? JSON.parse(p) : { game: 'g0', phase: 'lobby', i: -1, ids: [], startedAt: 0 };
}
function setState_(s) {
  PropertiesService.getScriptProperties().setProperty('state', JSON.stringify(s));
  CacheService.getScriptCache().removeAll(['rk', 'st', 'stfull']);
}

// Ranking gry na żywo, z cache 1,5 s
function liveRows_(game) {
  var s = sh_(LIVE), last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 8).getValues().filter(function (r) { return r[1] === game; });
}
function ranking_(game) {
  var cache = CacheService.getScriptCache(), hit = cache.get('rk');
  if (hit) return JSON.parse(hit);
  var by = {};
  liveRows_(game).forEach(function (r) {
    var n = String(r[2]);
    if (!by[n]) by[n] = { name: n, pts: 0, good: 0 };
    if (Number(r[3]) >= 0) { by[n].pts += Number(r[6]) || 0; by[n].good += Number(r[5]) || 0; }
  });
  var arr = Object.keys(by).map(function (k) { return by[k]; });
  arr.sort(function (a, b) { return b.pts - a.pts; });
  cache.put('rk', JSON.stringify(arr), 2);
  return arr;
}

function freeRanking_() {
  var s = sh_('Odpowiedzi'), last = s.getLastRow();
  if (last < 2) return [];
  var by = {};
  s.getRange(2, 1, last - 1, 4).getValues().forEach(function (r) {
    var n = String(r[1]); if (!n) return;
    var pts = Number(r[2]) || 0;
    if (!by[n] || pts > by[n].pts) by[n] = { name: n, pts: pts, good: Number(r[3]) || 0 };
  });
  return Object.keys(by).map(function (k) { return by[k]; }).sort(function (a, b) { return b.pts - a.pts; });
}

function stateOut_(p) {
  var c = cfg_(), st = getState_(), out = {
    mode: c.mode, phone: c.phone, time: c.time, now: Date.now(),
    game: st.game, phase: st.phase, i: st.i, n: st.ids.length, startedAt: st.startedAt
  };
  if (st.i >= 0 && st.i < st.ids.length && (st.phase === 'question' || st.phase === 'reveal' || st.phase === 'ranking')) {
    var q = questionById_(st.ids[st.i]);
    if (q) {
      out.q = { q: q.q, o: q.o, cat: q.cat };
      if (st.phase !== 'question') out.q.c = q.c;
    }
  }
  var rk = ranking_(st.game);
  if (p.name) {
    var pos = -1, me = null;
    rk.forEach(function (r, k) { if (r.name === p.name) { pos = k + 1; me = r; } });
    out.me = me ? { pts: me.pts, good: me.good, place: pos, of: rk.length } : null;
    if (me && st.phase === 'reveal') {
      var mine = liveRows_(st.game).filter(function (r) { return String(r[2]) === p.name && Number(r[3]) === st.i; })[0];
      out.last = mine ? { choice: Number(mine[4]), ok: Number(mine[5]), pts: Number(mine[6]) } : null;
    }
  }
  if (p.full) {
    var rows = liveRows_(st.game);
    out.joined = rk.length;
    out.answered = rows.filter(function (r) { return Number(r[3]) === st.i; }).length;
    if (st.phase === 'lobby') out.names = rk.map(function (r) { return r.name; });
    if (st.phase === 'reveal') {
      var d = [0, 0, 0, 0];
      rows.forEach(function (r) { if (Number(r[3]) === st.i && Number(r[4]) >= 0) d[Number(r[4])]++; });
      out.dist = d;
    }
    if (st.phase === 'reveal' || st.phase === 'ranking' || st.phase === 'final') out.rank = rk.slice(0, 10);
    if (c.mode === 'free') out.freerank = freeRanking_().slice(0, 10);
  }
  return out;
}

function doGet(e) {
  var p = e.parameter || {}, action = p.action || '', c = cfg_();
  if (action === '') return json_({ open: c.open });                        // zgodność ze starą wersją quizu
  if (action === 'config') return json_({ mode: c.mode, open: c.open, time: c.time, phone: c.phone, category: c.category });
  if (action === 'questions') {
    return json_({ time: c.time, open: c.open, questions: questions_(c.category).map(function (q) { return { q: q.q, o: q.o, c: q.c, cat: q.cat }; }) });
  }
  if (action === 'state') return json_(stateOut_(p));
  if (action === 'claim') {
    var scope = p.scope === 'live' ? getState_().game : 'free';
    if (!p.tok) return json_({ ok: 0 });
    return json_(claim_(scope, String(p.name || '').slice(0, 20), p.tok));
  }
  return json_({ error: 'unknown action' });
}

function doPost(e) {
  var d = JSON.parse(e.postData.contents), type = d.type || 'result', c = cfg_(), cache = CacheService.getScriptCache();
  if (type === 'result') {                                                   // wynik w trybie dowolnym
    if (d.tok && !claim_('free', String(d.name).slice(0, 20), d.tok).ok) return json_({ ok: 0, taken: 1 });
    sh_('Odpowiedzi').appendRow([new Date(), String(d.name).slice(0, 20), Number(d.pts), Number(d.good)]);
    return json_({ ok: 1 });
  }
  if (type === 'host') return json_(host_(d, c));
  var st = getState_(), name = String(d.name || '').slice(0, 20);
  if (!name) return json_({ ok: 0 });
  if (type === 'join') {
    if (!d.tok || !claim_(st.game, name, d.tok).ok) return json_({ ok: 0, taken: 1 });
    if (!cache.get('j|' + st.game + '|' + name)) {
      cache.put('j|' + st.game + '|' + name, '1', 21600);
      sh_(LIVE).appendRow([new Date(), st.game, name, -1, -1, 0, 0, 0]);
      cache.remove('rk');
    }
    return json_({ ok: 1 });
  }
  if (type === 'answer') {
    var i = Number(d.i), choice = Number(d.choice), now = Date.now();
    if (st.phase !== 'question' || i !== st.i) return json_({ ok: 0, late: 1 });
    if (now - st.startedAt > (c.time * 1000) + 2000) return json_({ ok: 0, late: 1 });
    var key = 'a|' + st.game + '|' + name + '|' + i;
    if (cache.get(key)) return json_({ ok: 0, dup: 1 });
    cache.put(key, '1', 21600);
    var q = questionById_(st.ids[i]), ok = q && q.c === choice ? 1 : 0;
    var ms = Math.min(c.time * 1000, Math.max(Number(d.ms) || 0, now - st.startedAt - 2000, 0));
    var pts = ok ? 500 + Math.round(500 * Math.max(0, 1 - ms / (c.time * 1000))) : 0;
    sh_(LIVE).appendRow([new Date(), st.game, name, i, choice, ok, pts, ms]);
    cache.remove('rk');
    return json_({ ok: 1 });
  }
  return json_({ ok: 0 });
}

function host_(d, c) {
  if (c.pin && String(d.pin) !== c.pin) return { ok: 0, error: 'pin' };
  var st = getState_(), now = Date.now(), cmd = d.cmd;
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    if (cmd === 'reset') {
      var s = sh_(LIVE); if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 8).clearContent();
      st = { game: 'g' + now, phase: 'lobby', i: -1, ids: [], startedAt: 0 };
    } else if (cmd === 'start') {
      var ids = questions_(c.category).map(function (q) { return q.id; });
      if (!ids.length) return { ok: 0, error: 'brak pytań' };
      st.ids = ids; st.i = 0; st.phase = 'question'; st.startedAt = now;
    } else if (cmd === 'reveal') {
      if (st.phase === 'question') st.phase = 'reveal';
    } else if (cmd === 'ranking') {
      if (st.phase === 'reveal') st.phase = 'ranking';
    } else if (cmd === 'next') {
      if (st.i + 1 < st.ids.length) { st.i++; st.phase = 'question'; st.startedAt = now; }
      else st.phase = 'final';
    } else if (cmd === 'final') {
      st.phase = 'final';
    }
    setState_(st);
  } finally { lock.releaseLock(); }
  return { ok: 1, phase: st.phase, i: st.i };
}
