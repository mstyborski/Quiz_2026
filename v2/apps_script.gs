// Quiz sezonu 2026: serwer w Apps Script (arkusz "Quiz 2026 – wyniki").
// Karty arkusza: Sterowanie, Pytania, Kategorie, Odpowiedzi, Odpowiedzi live, Ranking.

var LIVE = 'Odpowiedzi live';
var NAMES = 'Pseudonimy';   // rezerwacje pseudonimów: Gra, Klucz, Pseudonim, Kod urządzenia, Czas

// Klucz pseudonimu: bez rozróżniania wielkości liter i nadmiarowych spacji.
function nkey_(n) { return String(n).trim().toLowerCase().replace(/\s+/g, ' '); }

// Rezerwacja pseudonimu w danej grze ('free' w trybie dowolnym, id gry na żywo).
// Ten sam kod urządzenia może wracać; inny kod dostaje taken=1.
function claim_(scope, name, tok) {
  var lock = LockService.getScriptLock(); lock.waitLock(8000);
  try { return claimCore_(scope, name, tok); } finally { lock.releaseLock(); }
}
function claimCore_(scope, name, tok) {
  var key = nkey_(name);
  if (!key) return { ok: 0 };
  {
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
  }
}

function sh_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function cfg_() {
  var rows = sh_('Sterowanie').getRange('A2:B40').getValues(), c = {};
  rows.forEach(function (r) { if (r[0] !== '') c[String(r[0]).trim()] = r[1]; });
  var open = c['Odblokowane pytania (tryb dowolny)'];
  return {
    mode: c['Tryb gry'] === 'Na żywo' ? 'live' : 'free',
    category: String(c['Kategoria do gry'] || 'Wszystkie'),
    time: Number(c['Czas na pytanie (s)']) || 20,
    phone: String(c['Telefon w trybie na żywo pokazuje'] || '').indexOf('Tylko') === 0 ? 'buttons' : 'full',
    open: (open === '' || String(open).toLowerCase() === 'wszystkie') ? 999 : Number(open) || 0,
    pin: String(c['PIN prowadzącego'] === undefined ? '' : c['PIN prowadzącego']),
    change: c['Zmiana odpowiedzi w czasie pytania'] === 'Tak',
    folder: (function (v) { v = String(v || '').trim(); var m = v.match(/folders\/([\w-]+)/); return m ? m[1] : v; })(c['Folder ze zdjęciami i nagraniami (link lub ID)']),
    title: String(c['Nazwa quizu'] || 'Quiz żeglarski').trim(),
    sub: c.hasOwnProperty('Podtytuł (np. nazwa wydarzenia)') ? String(c['Podtytuł (np. nazwa wydarzenia)']).trim() : 'Podsumowanie Sezonu 2026',
    gp: c['Grand Prix'] === 'Włączone',
    gpName: String(c['Nazwa sezonu Grand Prix'] || 'Grand Prix'),
    gpN: Number(c['Liczba pytań w rundzie']) || 5,
    gpPool: String(c['Pula pytań (kategoria)'] || 'Wszystkie'),
    gpLimit: (function (v) { return (v === '' || v === undefined || String(v).toLowerCase().indexOf('bez') === 0) ? 999 : (Number(v) || 999); })(c['Limit prób w rundzie']),
    gpBest: String(c['Zasada punktów'] || '').indexOf('Tylko') !== 0
  };
}

// Pytania aktywne, z wybranej kategorii. id = numer wiersza w arkuszu.
function questions_(category) {
  var s = sh_('Pytania'), last = s.getLastRow();
  if (last < 2) return [];
  var rows = s.getRange(2, 1, last - 1, 11).getValues(), out = [];
  rows.forEach(function (r, k) {
    if (!r[1] || String(r[7]).toUpperCase() !== 'TAK') return;
    if (category && category !== 'Wszystkie' && String(r[0]) !== category) return;
    var c = 'ABCD'.indexOf(String(r[6]).trim().toUpperCase());
    if (c < 0) return;
    out.push({ id: k + 2, cat: String(r[0]), q: String(r[1]).trim(), o: [r[2], r[3], r[4], r[5]].map(String), c: c,
      typ: String(r[8] || 'Tekst').trim(), src: String(r[9] || '').trim(), cap: String(r[10] || '').trim() });
  });
  return out;
}

function questionById_(id) {
  var all = questions_('Wszystkie'), i;
  for (i = 0; i < all.length; i++) if (all[i].id === id) return all[i];
  return null;
}

// ======================= ZDJĘCIA I MUZYKA =======================
// Nazwa pliku (z rozszerzeniem lub bez) -> id, z folderu na Dysku Google. Cache 5 min.
function folderMap_(c) {
  if (!c.folder) return {};
  var cache = CacheService.getScriptCache(), hit = cache.get('fm');
  if (hit) return JSON.parse(hit);
  var m = {};
  try {
    var f = DriveApp.getFolderById(c.folder);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    var it = f.getFiles();
    while (it.hasNext()) {
      var x = it.next(), n = x.getName().toLowerCase(), b = n.replace(/\.[^.]+$/, '');
      m[n] = x.getId(); if (!m[b]) m[b] = x.getId();
    }
  } catch (e) {}
  cache.put('fm', JSON.stringify(m), 300);
  return m;
}

// Adres zdjęcia / dźwięku dla pytania; fragmenty muzyki szuka w iTunes. Wynik w cache 6 h.
function mediaFor_(q, c) {
  if (!q.src || q.typ === 'Tekst') return {};
  var cache = CacheService.getScriptCache(), ck = 'm|' + q.typ + '|' + q.src, hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  var out = {}, src = q.src, id = folderMap_(c)[src.toLowerCase()] || null;
  if (!id) { var mm = src.match(/\/d\/([\w-]{10,})|[?&]id=([\w-]{10,})/); if (mm) id = mm[1] || mm[2]; }
  if (q.typ === 'Zdjęcie') {
    if (id) out.img = 'https://lh3.googleusercontent.com/d/' + id + '=w1600';
    else if (/^https?:/i.test(src)) out.img = src;
  } else if (q.typ === 'Muzyka') {
    if (id) out.audio = 'https://drive.google.com/uc?export=download&id=' + id;
    else if (/^https?:/i.test(src)) out.audio = src;
    else {
      try {
        var r = JSON.parse(UrlFetchApp.fetch('https://itunes.apple.com/search?media=music&entity=song&limit=5&country=PL&term=' + encodeURIComponent(src), { muteHttpExceptions: true }).getContentText());
        var t = (r.results || []).filter(function (x) { return x.previewUrl; })[0];
        if (t) { out.audio = t.previewUrl; out.auto = t.artistName + ' – ' + t.trackName; }
      } catch (e) {}
    }
  }
  if (out.img || out.audio) cache.put(ck, JSON.stringify(out), 21600);
  return out;
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
  var rows = s.getRange(2, 1, last - 1, 8).getValues().filter(function (r) { return r[1] === game; });
  // przy zmianie odpowiedzi liczy się ostatni wpis gracza dla danego pytania
  var seen = {}, out = [];
  rows.forEach(function (r) {
    if (Number(r[3]) < 0) { out.push(r); return; }
    var k = String(r[2]) + '|' + r[3];
    if (seen[k] === undefined) { seen[k] = out.length; out.push(r); } else out[seen[k]] = r;
  });
  return out;
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
    mode: c.mode, phone: c.phone, time: c.time, change: c.change, title: c.title, sub: c.sub, now: Date.now(),
    game: st.game, phase: st.phase, i: st.i, n: st.ids.length, startedAt: st.startedAt
  };
  if (st.i >= 0 && st.i < st.ids.length && (st.phase === 'question' || st.phase === 'reveal' || st.phase === 'ranking')) {
    var q = questionById_(st.ids[st.i]);
    if (q) {
      var md = mediaFor_(q, c);
      out.q = { q: q.q, o: q.o, cat: q.cat, typ: q.typ, img: md.img, audio: md.audio };
      if (st.phase !== 'question') { out.q.c = q.c; out.q.cap = q.cap || md.auto || ''; }
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


// ======================= GRAND PRIX =======================
var GPA = 'Odpowiedzi GP', GPK = 'Klasyfikacja GP', GPT = 'Terminarz GP';

function toDate_(v) { if (v === '' || v === null || v === undefined) return null; var d = v instanceof Date ? v : new Date(v); return isNaN(d.getTime()) ? null : d; }

// Rundy z terminarza; status wg dat albo wymuszony ('Otwarta'/'Zamknięta')
function rounds_() {
  var s = sh_(GPT); if (!s || s.getLastRow() < 2) return [];
  var now = Date.now(), out = [];
  s.getRange(2, 1, s.getLastRow() - 1, 6).getValues().forEach(function (r) {
    var n = Number(r[0]); if (!n) return;
    var a = toDate_(r[1]), b = toDate_(r[2]), f = String(r[5] || '').trim(), st;
    if (f === 'Otwarta' || f === 'Zamknięta') st = f;
    else if (a && now < a.getTime()) st = 'Zaplanowana';
    else if (b && now > b.getTime()) st = 'Zamknięta';
    else st = 'Otwarta';
    out.push({ n: n, start: a ? a.getTime() : 0, end: b ? b.getTime() : 0, pool: String(r[3] || ''), cnt: Number(r[4]) || 0, status: st });
  });
  return out;
}
function openRound_() {
  var rs = rounds_().filter(function (r) { return r.status === 'Otwarta'; });
  return rs.length ? rs[rs.length - 1] : null;
}
function nextRound_() {
  var rs = rounds_().filter(function (r) { return r.status === 'Zaplanowana'; });
  rs.sort(function (a, b) { return a.start - b.start; });
  return rs[0] || null;
}

// Te same pytania w tej samej kolejności dla wszystkich graczy danej rundy
function roundQuestions_(round, c) {
  var pool = round.pool && round.pool !== 'Wszystkie' ? round.pool : c.gpPool;
  var qs = questions_(pool === 'Wszystkie' ? '' : pool).sort(function (a, b) { return a.id - b.id; });
  var seed = round.n * 7919 + 13;
  function rnd() { seed = (seed * 1664525 + 1013904223) % 4294967296; return seed / 4294967296; }
  for (var i = qs.length - 1; i > 0; i--) { var j = Math.floor(rnd() * (i + 1)), t = qs[i]; qs[i] = qs[j]; qs[j] = t; }
  return qs.slice(0, round.cnt || c.gpN);
}

function gpRows_() {
  var s = sh_(GPA); if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 7).getValues();   // czas, runda, klucz, pseudonim, próba, punkty, poprawne
}

function gpRanking_(c) {
  var by = {}, maxRound = 0;
  gpRows_().forEach(function (r) {
    var k = String(r[2]), n = Number(r[1]), pts = Number(r[5]) || 0; if (!k) return;
    if (n > maxRound) maxRound = n;
    var p = by[k] || (by[k] = { name: String(r[3]), key: k, per: {}, total: 0, rounds: 0 });
    if (p.per[n] === undefined) p.per[n] = pts;
    else if (c.gpBest && pts > p.per[n]) p.per[n] = pts;
  });
  var arr = Object.keys(by).map(function (k) {
    var p = by[k]; Object.keys(p.per).forEach(function (n) { p.total += p.per[n]; p.rounds++; }); return p;
  });
  arr.sort(function (a, b) { return b.total - a.total || b.rounds - a.rounds; });
  rounds_().forEach(function (r) { if (r.n > maxRound) maxRound = r.n; });
  arr.maxRound = Math.min(Math.max(maxRound, 1), 20);
  return arr;
}

// Odświeża zakładkę 'Klasyfikacja GP' (czytelną dla prowadzącego i do wykresu)
function writeGpSheet_(c) {
  var arr = gpRanking_(c), m = arr.maxRound, s = sh_(GPK);
  if (!s) s = SpreadsheetApp.getActive().insertSheet(GPK);
  var head = ['Miejsce', 'Pseudonim', 'Suma', 'Rozegrane rundy'], k;
  for (k = 1; k <= m; k++) head.push('Runda ' + k);
  var rows = arr.map(function (p, i) {
    var row = [i + 1, p.name, p.total, p.rounds];
    for (var n = 1; n <= m; n++) row.push(p.per[n] === undefined ? '' : p.per[n]);
    return row;
  });
  s.getRange(1, 1, 200, 24).clearContent();   // kolumny Z:AA (dane wykresu) zostają
  s.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length) s.getRange(2, 1, rows.length, head.length).setValues(rows);
}

function gpPublic_(c, name) {
  var cache = CacheService.getScriptCache(), hit = cache.get('gprank'), rk;
  if (hit) rk = JSON.parse(hit);
  else {
    rk = gpRanking_(c).slice(0, 30).map(function (p) { return { name: p.name, key: p.key, total: p.total, rounds: p.rounds }; });
    cache.put('gprank', JSON.stringify(rk), 5);
  }
  var out = { on: c.gp, name: c.gpName, title: c.title, sub: c.sub, change: c.change, best: c.gpBest, limit: c.gpLimit, time: c.time, now: Date.now(), rank: rk.slice(0, 20) };
  var r = openRound_(), nx = nextRound_();
  if (c.gp && r) {
    out.round = { n: r.n, end: r.end, start: r.start, count: roundQuestions_(r, c).length };
    if (name) {
      var key = nkey_(name);
      out.tries = gpRows_().filter(function (x) { return Number(x[1]) === r.n && String(x[2]) === key; }).length;
    }
  }
  if (c.gp && !r && nx) out.next = { n: nx.n, start: nx.start };
  if (name) { rk.forEach(function (p, i) { if (p.key === nkey_(name)) out.me = { total: p.total, place: i + 1, rounds: p.rounds, of: rk.length }; }); }
  return out;
}

function gpQuestions_(c, p) {
  var r = openRound_();
  if (!c.gp || !r) return { ok: 0, closed: 1 };
  var qs = roundQuestions_(r, c);
  return { ok: 1, round: r.n, end: r.end, time: c.time, change: c.change,
    questions: qs.map(function (q) { var md = mediaFor_(q, c); return { id: q.id, q: q.q, o: q.o, cat: q.cat, img: md.img, audio: md.audio }; }) };
}

function gpSubmit_(c, p) {
  var r = openRound_();
  if (!c.gp || !r) return { ok: 0, closed: 1 };
  if (Number(p.round) !== r.n) return { ok: 0, closed: 1 };
  var name = String(p.name || '').trim().slice(0, 20), key = nkey_(name);
  if (!key || !p.tok) return { ok: 0 };
  var lock = LockService.getScriptLock(); lock.waitLock(8000);
  try {
    if (!claimCore_('gp', name, p.tok).ok) return { ok: 0, taken: 1 };
    var rows = gpRows_().filter(function (x) { return Number(x[1]) === r.n && String(x[2]) === key; });
    if (rows.length >= c.gpLimit) return { ok: 0, limit: 1 };
    var qs = roundQuestions_(r, c), a = String(p.a || '').split(','), m = String(p.m || '').split(','), pts = 0, good = 0, det = [];
    qs.forEach(function (q, i) {
      var ch = Number(a[i]); if (isNaN(ch)) ch = -1;
      var ms = Math.min(c.time * 1000, Math.max(0, Number(m[i]) || 0));
      var ok = ch === q.c ? 1 : 0;
      if (ok) { good++; pts += 500 + Math.round(500 * Math.max(0, 1 - ms / (c.time * 1000))); }
      det.push({ ch: ch, c: q.c, ok: ok, cap: q.cap || mediaFor_(q, c).auto || '' });
    });
    var s = sh_(GPA);
    if (!s) { s = SpreadsheetApp.getActive().insertSheet(GPA); s.appendRow(['Czas', 'Runda', 'Klucz', 'Pseudonim', 'Próba', 'Punkty', 'Poprawne']); s.setFrozenRows(1); }
    s.appendRow([new Date(), r.n, key, name, rows.length + 1, pts, good]);
    CacheService.getScriptCache().remove('gprank');
    writeGpSheet_(c);
    var rk = gpRanking_(c), place = 0, tot = 0;
    rk.forEach(function (x, i) { if (x.key === key) { place = i + 1; tot = x.total; } });
    return { ok: 1, pts: pts, good: good, n: qs.length, attempt: rows.length + 1, details: det, total: tot, place: place, of: rk.length,
      left: Math.max(0, c.gpLimit - rows.length - 1) };
  } finally { lock.releaseLock(); }
}

// ----- Menu w arkuszu: archiwizacja sezonu i zwalnianie pseudonimów -----
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Quiz')
    .addItem('Zakończ sezon Grand Prix i zarchiwizuj', 'archiveGp')
    .addItem('Usuń gracza / zwolnij pseudonim…', 'releaseName')
    .addToUi();
}
function archiveGp() {
  var ui = SpreadsheetApp.getUi(), c = cfg_();
  if (ui.alert('Zarchiwizować klasyfikację "' + c.gpName + '" i zacząć nowy sezon od zera?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  archiveCore_(c);
  ui.alert('Gotowe. Klasyfikacja została skopiowana do nowej zakładki, a sezon wyzerowany.');
}
function archiveCore_(c) {
  writeGpSheet_(c);
  var ss = SpreadsheetApp.getActive(), src = sh_(GPK);
  var stamp = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'yyyy-MM-dd HH-mm');
  var dst = src.copyTo(ss); dst.setName(('Archiwum GP ' + stamp).slice(0, 90));
  var a = sh_(GPA); if (a && a.getLastRow() > 1) a.getRange(2, 1, a.getLastRow() - 1, 7).clearContent();
  var n = sh_(NAMES);
  if (n && n.getLastRow() > 1) {
    var rows = n.getRange(2, 1, n.getLastRow() - 1, 5).getValues();
    for (var i = rows.length - 1; i >= 0; i--) if (String(rows[i][0]) === 'gp') n.deleteRow(i + 2);
  }
  writeGpSheet_(c);
  CacheService.getScriptCache().remove('gprank');
}
function releaseName() {
  var ui = SpreadsheetApp.getUi(), r = ui.prompt('Jaki pseudonim usunąć (rezerwacja oraz gracz z bieżącej gry na żywo)?');
  if (r.getSelectedButton() !== ui.Button.OK) return;
  ui.alert('Usunięto wpisów: ' + releaseCore_(r.getResponseText()));
}
function releaseCore_(name) {
  var key = nkey_(name), n = sh_(NAMES), cnt = 0, i, rows;
  if (n && n.getLastRow() >= 2) {
    rows = n.getRange(2, 1, n.getLastRow() - 1, 5).getValues();
    for (i = rows.length - 1; i >= 0; i--) if (String(rows[i][1]) === key) { n.deleteRow(i + 2); cnt++; }
  }
  // usuń też gracza z bieżącej gry na żywo (lista w lobby, punkty, odpowiedzi)
  var game = getState_().game, s = sh_(LIVE), cache = CacheService.getScriptCache();
  if (s && s.getLastRow() >= 2) {
    rows = s.getRange(2, 1, s.getLastRow() - 1, 8).getValues();
    for (i = rows.length - 1; i >= 0; i--) {
      if (rows[i][1] === game && nkey_(rows[i][2]) === key) {
        cache.remove('j|' + game + '|' + String(rows[i][2]));
        s.deleteRow(i + 2); cnt++;
      }
    }
  }
  cache.removeAll(['rk', 'st', 'stfull', 'gprank']);
  return cnt;
}

function doGet(e) {
  var p = e.parameter || {}, action = p.action || '', c = cfg_();
  if (action === '') return json_({ open: c.open });                        // zgodność ze starą wersją quizu
  if (action === 'config') return json_({ mode: c.mode, open: c.open, time: c.time, phone: c.phone, category: c.category, change: c.change, title: c.title, sub: c.sub });
  if (action === 'questions') {
    return json_({ time: c.time, open: c.open, change: c.change, title: c.title, sub: c.sub, questions: questions_(c.category).map(function (q) { var md = mediaFor_(q, c); return { q: q.q, o: q.o, c: q.c, cat: q.cat, img: md.img, audio: md.audio, cap: q.cap || md.auto || '' }; }) });
  }
  if (action === 'state') return json_(stateOut_(p));
  if (action === 'gp') return json_(gpPublic_(c, String(p.name || '').slice(0, 20)));
  if (action === 'gpquestions') return json_(gpQuestions_(c, p));
  if (action === 'gpsubmit') return json_(gpSubmit_(c, p));
  if (action === 'claim') {
    var scope = p.scope === 'live' ? getState_().game : (p.scope === 'gp' ? 'gp' : 'free');
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
    var key = 'a|' + st.game + '|' + name + '|' + i, prev = cache.get(key);
    if (prev !== null && (!c.change || Number(prev) === choice)) return json_({ ok: 0, dup: 1 });
    cache.put(key, String(choice), 21600);
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
