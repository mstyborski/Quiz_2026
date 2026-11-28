// Quiz sezonu 2026: serwer w Apps Script (arkusz "Quiz 2026 – wyniki").
// Karty arkusza: Sterowanie, Pytania, Kategorie, Odpowiedzi, Odpowiedzi live, Ranking, Historia wyników.

var LIVE = 'Odpowiedzi live';
var NAMES = 'Pseudonimy';   // rezerwacje pseudonimów: Gra, Klucz, Pseudonim, Kod urządzenia, Czas

// Klucz pseudonimu: bez rozróżniania wielkości liter i nadmiarowych spacji.
function nkey_(n) { return String(n).trim().toLowerCase().replace(/\s+/g, ' '); }

// Rezerwacja pseudonimu w danej grze ('free' w trybie dowolnym, id gry na żywo).
// Ten sam kod urządzenia może wracać; inny kod dostaje taken=1.
function claim_(scope, name, tok, uid) {
  var lock = LockService.getScriptLock(); lock.waitLock(8000);
  try { return claimCore_(scope, name, tok, uid); } finally { lock.releaseLock(); }
}
function claimCore_(scope, name, tok, uid) {
  var key = nkey_(name);
  if (!key) return { ok: 0 };
  {
    var s = sh_(NAMES);
    if (!s) { s = SpreadsheetApp.getActive().insertSheet(NAMES); s.appendRow(['Gra', 'Klucz', 'Pseudonim', 'Kod urządzenia', 'Czas', 'Dane uczestnika (niewidoczne na ekranie)']); s.setFrozenRows(1); }
    var last = s.getLastRow();
    if (last >= 2) {
      var rows = s.getRange(2, 1, last - 1, 4).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]) === scope && String(rows[i][1]) === key) {
          if (String(rows[i][3]) !== String(tok)) return { ok: 0, taken: 1 };
          if (uid) s.getRange(i + 2, 6).setValue(String(uid).slice(0, 60));
          return { ok: 1 };
        }
      }
    }
    s.appendRow([scope, key, String(name).trim().slice(0, 20), String(tok), new Date(), uid ? String(uid).slice(0, 60) : '']);
    return { ok: 1 };
  }
}

function sh_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function cfg_() {
  var rows = sh_('Sterowanie').getRange('A2:B80').getValues(), c = {};
  rows.forEach(function (r) { if (r[0] !== '') c[String(r[0]).trim()] = r[1]; });
  var open = c['Odblokowane pytania (tryb dowolny)'];
  return {
    mode: c['Tryb gry'] === 'Na żywo' ? 'live' : (String(c['Tryb gry'] || '').indexOf('Głosowanie') === 0 ? 'poll' : 'free'),
    pset: String(c['Zestaw głosowań i opinii'] || 'Wszystkie').trim() || 'Wszystkie',
    category: String(c['Kategoria do gry'] || 'Wszystkie'),
    time: Number(c['Czas na pytanie (s)']) || 20,
    phone: String(c['Telefon w trybie na żywo pokazuje'] || '').indexOf('Tylko') === 0 ? 'buttons' : 'full',
    open: (c['Gra z przerwami'] === 'Nie' || open === '' || String(open).toLowerCase() === 'wszystkie') ? 999 : Number(open) || 0,
    pin: String(c['PIN prowadzącego'] === undefined ? '' : c['PIN prowadzącego']),
    change: c['Zmiana odpowiedzi w czasie pytania'] === 'Tak',
    folder: (function (v) { v = String(v || '').trim(); var m = v.match(/folders\/([\w-]+)/); return m ? m[1] : v; })(c['Folder ze zdjęciami i nagraniami (link lub ID)']),
    nq: (function (v) { var n = Number(v); return (v === '' || v === undefined || isNaN(n) || n <= 0) ? 0 : n; })(c['Liczba pytań w quizie']),
    breaks: c['Gra z przerwami'] !== 'Nie',
    breakText: String(c['Tekst na planszy podczas przerwy'] || 'Za chwilę wracamy do gry…'),
    logo: String(c['Logo (nazwa pliku z folderu lub adres)'] || '').trim(),
    palette: String(c['Paleta kolorów'] || '').trim(),
    idOn: c['Dane uczestnika (opcjonalne)'] === 'Włączone',
    idLabel: String(c['Etykieta pola danych uczestnika'] || 'Imię i nazwisko lub nr indeksu'),
    idInfo: String(c['Informacja przy polu danych'] || 'Nie wyświetlamy tego na ekranie. Podaj, jeśli chcesz być uwzględniony w nagrodach.'),
    freeLimit: (function (v) { return (v === '' || v === undefined || String(v).toLowerCase().indexOf('bez') === 0) ? 999 : (Number(v) || 999); })(c['Tryb dowolny: limit prób']),
    freeBest: String(c['Tryb dowolny: zasada punktów'] || '').indexOf('Tylko') !== 0,
    random: c['Losowa kolejność pytań'] === 'Tak',
    streak: c['Bonus za serię poprawnych odpowiedzi'] === 'Tak',
    teams: String(c['Drużyny (nazwy po przecinku)'] || '').split(/[,;\n]/).map(function (x) { return x.trim().slice(0, 30); }).filter(function (x) { return x; }).slice(0, 12),
    teamAvg: String(c['Wynik drużyny'] || '').indexOf('Suma') !== 0,
    peer: c['Ponowne głosowanie (peer instruction)'] === 'Tak',
    round: (function (v) { var n = Number(v); return isNaN(n) || n <= 0 ? 0 : Math.floor(n); })(c['Runda quizu']),
    freeFrom: toDate_(c['Tryb dowolny: otwarty od']),
    freeTo: toDate_(c['Tryb dowolny: otwarty do']),
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
// Pytania do gry: cała kategoria albo jej losowy wybór (Losowa kolejność pytań = Tak),
// przycięte do liczby pytań w quizie. Ten sam seed daje ten sam zestaw i kolejność.
function pick_(c, seedStr) {
  var list = questions_(c.category);
  if (c.random) {
    var s = 0, str = String(seedStr) + '|' + c.category, i;
    for (i = 0; i < str.length; i++) s = (s * 31 + str.charCodeAt(i)) % 4294967296;
    function rnd() { s = (s * 1664525 + 1013904223) % 4294967296; return s / 4294967296; }
    for (i = list.length - 1; i > 0; i--) { var j = Math.floor(rnd() * (i + 1)), t = list[i]; list[i] = list[j]; list[j] = t; }
  }
  var off = (c.round && c.nq) ? (c.round - 1) * c.nq : 0;   // Runda quizu: kolejna paczka pytań
  return list.slice(off, off + (c.nq || 999));
}
// Seed trybu dowolnego: wszyscy w danej sesji (okno „otwarty od” albo dany dzień) dostają ten sam zestaw.
function freeSeed_(c) {
  return c.freeFrom ? c.freeFrom.getTime() : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
}

function questions_(category) {
  var s = sh_('Pytania'), last = s.getLastRow();
  if (last < 2) return [];
  var rows = s.getRange(2, 1, last - 1, QCOLS).getValues(), out = [];
  rows.forEach(function (r, k) {
    if (!r[1] || String(r[7]).toUpperCase() !== 'TAK') return;
    if (category && category !== 'Wszystkie' && String(r[0]) !== category) return;
    var kind = kind_(r[11]), q = { id: k + 2, cat: String(r[0]), q: String(r[1]).trim(), kind: kind,
      typ: String(r[8] || 'Tekst').trim(), src: String(r[9] || '').trim(), cap: String(r[10] || '').trim(), expl: String(r[12] || '').trim() };
    if (kind === 'num') {
      q.val = num_(r[13]); q.tol = Math.abs(num_(r[14])) || 0; q.o = []; q.c = -1;
      if (!isFinite(q.val)) return;
    } else {
      q.o = kind === 'tf' ? [String(r[2] || 'Prawda'), String(r[3] || 'Fałsz')] : [r[2], r[3], r[4], r[5]].map(String);
      q.c = 'ABCD'.indexOf(String(r[6]).trim().toUpperCase());
      if (q.c < 0 || q.c >= q.o.length) return;
    }
    out.push(q);
  });
  return out;
}

// Kolumny karty Pytania: A Kategoria, B Pytanie, C–F odpowiedzi A–D, G Poprawna, H Aktywne, I Typ (multimedia),
// J Źródło, K Podpis, L Rodzaj odpowiedzi, M Wyjaśnienie, N Wartość poprawna (Liczba), O Tolerancja (Liczba)
var QCOLS = 15;
var QHEAD = ['Rodzaj odpowiedzi (A–D / Prawda/Fałsz / Liczba)', 'Wyjaśnienie po odpowiedzi', 'Wartość poprawna (Liczba)', 'Tolerancja (Liczba)'];
function kind_(v) { v = String(v || '').toLowerCase(); return v.indexOf('prawda') >= 0 ? 'tf' : (v.indexOf('liczb') >= 0 ? 'num' : 'choice'); }
function num_(v) { if (typeof v === 'number') return v; var s = String(v === undefined || v === null ? '' : v).replace(/\s/g, '').replace(',', '.'); return s === '' ? NaN : Number(s); }

// Punkty za odpowiedź: 500 za poprawność + do 500 za szybkość. Liczba: poprawna, gdy błąd ≤ tolerancja;
// im bliżej wartości dokładnej, tym więcej punktów (na granicy tolerancji połowa).
function score_(q, ans, ms, time) {
  var speed = Math.max(0, 1 - ms / (time * 1000)), ok = 0, f = 0;
  if (q.kind === 'num') {
    var x = num_(ans);
    if (isFinite(x)) { var err = Math.abs(x - q.val); if (err <= q.tol + 1e-9) { ok = 1; f = q.tol > 0 ? 1 - 0.5 * err / q.tol : 1; } }
  } else if (Number(ans) === q.c) { ok = 1; f = 1; }
  return { ok: ok, pts: ok ? Math.round(f * (500 + 500 * speed)) : 0 };
}
// Pytanie wysyłane do przeglądarki; reveal = z poprawną odpowiedzią i wyjaśnieniem.
function pubQ_(q, c, reveal) {
  var md = mediaFor_(q, c), o = { id: q.id, q: q.q, o: q.o, kind: q.kind, cat: q.cat, typ: q.typ, img: md.img, audio: md.audio };
  if (reveal) { o.c = q.c; o.cap = q.cap || md.auto || ''; o.expl = q.expl; if (q.kind === 'num') { o.val = q.val; o.tol = q.tol; } }
  return o;
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

// Telefony nie są zalogowane na konto właściciela: plik musi być dostępny dla każdego z linkiem.
function ensurePublic_(id) {
  try {
    var f = DriveApp.getFileById(id), a = f.getSharingAccess();
    if (a !== DriveApp.Access.ANYONE_WITH_LINK && a !== DriveApp.Access.ANYONE) f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return '';
  } catch (e) { return 'Nie udało się udostępnić pliku publicznie: ' + e; }
}

// Adres zdjęcia / dźwięku dla pytania; fragmenty muzyki szuka w iTunes. Wynik w cache 6 h.
function mediaFor_(q, c) {
  if (!q.src || q.typ === 'Tekst') return {};
  var cache = CacheService.getScriptCache(), ck = 'm|' + q.typ + '|' + q.src, hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  var out = {}, src = q.src, id = folderMap_(c)[src.toLowerCase()] || null;
  if (!id) { var mm = src.match(/\/d\/([\w-]{10,})|[?&]id=([\w-]{10,})/); if (mm) id = mm[1] || mm[2]; }
  if (id) { var w = ensurePublic_(id); if (w) out.warn = w; }
  if (q.typ === 'Zdjęcie') {
    if (id) out.img = 'https://lh3.googleusercontent.com/d/' + id + '=w1600';
    else if (/^https?:/i.test(src)) out.img = src;
  } else if (q.typ === 'Muzyka') {
    if (id) out.audio = 'https://drive.google.com/uc?export=download&id=' + id;
    else if (/^https?:/i.test(src)) out.audio = src;
    else {
      var term = src.replace(/[–—\-]+/g, ' ').replace(/\s+/g, ' ').trim(), cs = ['PL', 'US'];
      for (var ci = 0; ci < cs.length && !out.audio; ci++) {
        try {
          var r = JSON.parse(UrlFetchApp.fetch('https://itunes.apple.com/search?media=music&entity=song&limit=10&country=' + cs[ci] + '&term=' + encodeURIComponent(term), { muteHttpExceptions: true }).getContentText());
          var t = (r.results || []).filter(function (x) { return x.previewUrl; })[0];
          if (t) { out.audio = t.previewUrl; out.auto = t.artistName + ' – ' + t.trackName; }
          else out.err = 'iTunes: brak wyników dla "' + term + '"';
        } catch (e) { out.err = String(e); }
      }
      if (out.audio) delete out.err;
    }
  }
  if (!out.img && !out.audio && !out.err) out.err = 'Nie znaleziono pliku "' + src + '" w folderze (sprawdź nazwę i pole Folder w Sterowaniu).';
  if ((out.img || out.audio) && !out.warn) cache.put(ck, JSON.stringify(out), 21600);
  return out;
}

// Okno czasowe trybu dowolnego: 'open' | 'before' | 'after' (puste daty = zawsze otwarte)
function freeWindow_(c, graceMs) {
  var now = Date.now(), a = c.freeFrom ? c.freeFrom.getTime() : 0, b = c.freeTo ? c.freeTo.getTime() : 0;
  var status = (a && now < a) ? 'before' : ((b && now > b + (graceMs || 0)) ? 'after' : 'open');
  return { status: status, from: a, to: b };
}

// Wspólne dane wyglądu i ustawień dołączane do odpowiedzi
function withMeta_(o, c) {
  var lg = c.logo ? (mediaFor_({ typ: 'Zdjęcie', src: c.logo }, c).img || '') : '';
  o.title = c.title; o.sub = c.sub; o.logo = lg; o.palette = c.palette; o.brk = c.breakText;
  o.idl = c.idOn ? { label: c.idLabel, info: c.idInfo } : null;
  return o;
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
  var rows = s.getRange(2, 1, last - 1, 10).getValues().filter(function (r) { return r[1] === game; });
  // przy zmianie odpowiedzi liczy się ostatni wpis gracza dla danego pytania
  var seen = {}, out = [];
  rows.forEach(function (r) {
    if (Number(r[3]) < 0) { out.push(r); return; }
    var k = String(r[2]) + '|' + r[3];
    if (seen[k] === undefined) { seen[k] = out.length; out.push(r); }
    else if (String(out[seen[k]][4]) !== String(r[4])) out[seen[k]] = r;   // ponowiony zapis tej samej odpowiedzi nie zmienia czasu
  });
  return out;
}
// Bonus za serię: druga z rzędu poprawna odpowiedź +100, trzecia +200 … maksymalnie +500 za pytanie.
function bonus_(n) { return n >= 2 ? Math.min(n - 1, 5) * 100 : 0; }
// Wyniki graczy gry na żywo: punkty, poprawne, drużyna, bieżąca seria i bonus za każde pytanie (lb: {i: bonus}).
function liveScores_(game, streakOn) {
  var by = {}, maxI = -1, rows = liveRows_(game);
  rows.forEach(function (r) {
    var n = String(r[2]), i = Number(r[3]);
    if (!by[n]) by[n] = { name: n, pts: 0, good: 0, team: '', streak: 0, bonus: 0, lb: {}, ans: {} };
    if (i < 0) { if (r[9]) by[n].team = String(r[9]); return; }
    by[n].pts += Number(r[6]) || 0; by[n].good += Number(r[5]) || 0; by[n].ans[i] = Number(r[5]) || 0;
    if (i > maxI) maxI = i;
  });
  var arr = Object.keys(by).map(function (k) {
    var p = by[k], st = 0;
    if (streakOn) for (var i = 0; i <= maxI; i++) {
      if (p.ans[i]) { st++; var b = bonus_(st); if (b) { p.lb[i] = b; p.bonus += b; } } else if (p.ans.hasOwnProperty(i) || i < maxI) st = 0;
    }
    p.streak = st; p.pts += p.bonus; delete p.ans;
    return p;
  });
  return arr.sort(function (a, b) { return b.pts - a.pts; });
}
function ranking_(game, streakOn) {
  var cache = CacheService.getScriptCache(), hit = cache.get('rk');
  if (hit) return JSON.parse(hit);
  var arr = liveScores_(game, streakOn);
  try { cache.put('rk', JSON.stringify(arr), 2); } catch (e) {}
  return arr;
}
// Ranking drużyn: średnia (domyślnie) albo suma punktów członków.
function teamRank_(arr, c) {
  if (!c.teams.length) return null;
  var by = {};
  c.teams.forEach(function (t) { by[t] = { team: t, sum: 0, n: 0 }; });
  arr.forEach(function (p) { if (p.team && by[p.team]) { by[p.team].sum += p.pts; by[p.team].n++; } });
  return Object.keys(by).map(function (k) { var t = by[k]; t.pts = c.teamAvg ? (t.n ? Math.round(t.sum / t.n) : 0) : t.sum; return t; })
    .sort(function (a, b) { return b.pts - a.pts || b.n - a.n; });
}

// ---------- Ranking bieżący (karta Ranking + wykres) i historia gier ----------
var HIST = 'Historia wyników', RANKS = 'Ranking';

function uidMap_(scope) {
  var n = sh_(NAMES), m = {};
  if (!n || n.getLastRow() < 2) return m;
  n.getRange(2, 1, n.getLastRow() - 1, 6).getValues().forEach(function (r) {
    if (String(r[0]) === String(scope) && r[5]) m[String(r[1])] = String(r[5]);
  });
  return m;
}
function liveRank_(game) { return liveScores_(game, cfg_().streak); }   // bez pamięci podręcznej
// Karta Ranking: wiersz 1 tytuł, wiersz 2 nagłówki, od wiersza 3 gracze.
// Kolumny Z:AA (formuły) zasilają wykres, więc ich nie czyścimy.
function writeRankSheet_(title, arr, uids) {
  var s = sh_(RANKS) || SpreadsheetApp.getActive().insertSheet(RANKS);
  s.getRange(1, 1, 300, 6).clearContent();
  s.getRange(1, 1, 2, 5).setValues([[title, '', '', '', ''], ['Miejsce', 'Pseudonim', 'Punkty', 'Poprawne', 'Dane uczestnika']]);
  var rows = arr.slice(0, 295).map(function (p, i) { return [i + 1, p.name, p.pts, p.good, uids[nkey_(p.name)] || '']; });
  if (rows.length) s.getRange(3, 1, rows.length, 5).setValues(rows);
}
function updateLiveRank_(c, st) {
  writeRankSheet_('Gra na żywo · ' + c.title + (c.sub ? ' · ' + c.sub : '') + ' · ' +
    Utilities.formatDate(new Date(Number(String(st.game).slice(1)) || Date.now()), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    liveRank_(st.game), uidMap_(st.game));
}
// Zapis zakończonej gry na żywo do karty Historia wyników (raz na grę).
function archiveLive_(game, c) {
  if (!game) return 0;
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('arch_' + game)) return 0;
  var arr = liveRank_(game);
  if (!arr.length) return 0;
  var uids = uidMap_(game), st = getState_(), nq = st.game === game ? st.ids.length : '';
  var s = sh_(HIST);
  if (!s) { s = SpreadsheetApp.getActive().insertSheet(HIST); s.appendRow(['Data gry', 'Tryb', 'Identyfikator gry', 'Quiz', 'Miejsce', 'Pseudonim', 'Dane uczestnika', 'Punkty', 'Poprawne', 'Liczba pytań']); s.setFrozenRows(1); }
  var when = new Date(Number(String(game).slice(1)) || Date.now()), quiz = c.title + (c.sub ? ' · ' + c.sub : '');
  var rows = arr.map(function (p, i) { return [when, 'Na żywo', game, quiz, i + 1, p.name, uids[nkey_(p.name)] || '', p.pts, p.good, nq]; });
  s.getRange(s.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
  props.setProperty('arch_' + game, '1');
  return rows.length;
}
function saveCurrentGame() {
  var n = archiveLive_(getState_().game, cfg_());
  SpreadsheetApp.getUi().alert(n ? 'Zapisano ' + n + ' wyników w karcie "' + HIST + '".' : 'Bieżąca gra jest już zapisana albo nie ma w niej odpowiedzi.');
}

// Wiersze trybu dowolnego bieżącej sesji (od początku okna, jeśli ustawione).
// Kolumny: Czas, Pseudonim, Punkty, Poprawne, Dane uczestnika, Quiz, Próba.
function freeRows_(c) {
  var s = sh_('Odpowiedzi'), last = s.getLastRow();
  if (last < 2) return [];
  var from = c.freeFrom ? c.freeFrom.getTime() : 0;
  return s.getRange(2, 1, last - 1, 7).getValues().map(function (r, i) {
    return { t: r[0] instanceof Date ? r[0].getTime() : 0, name: String(r[1]), key: nkey_(r[1]), pts: Number(r[2]) || 0, good: Number(r[3]) || 0, att: String(r[6] || ('w' + i)) };
  }).filter(function (r) { return r.name && (!from || r.t >= from); });
}
function freeAttempts_(c, key, rows) {
  var seen = {}, out = [];
  (rows || freeRows_(c)).forEach(function (r) { if (r.key === key && !seen[r.att]) { seen[r.att] = 1; out.push(r.att); } });
  return out;
}
// Ranking: w każdej próbie liczy się jej najwyższy (końcowy) wynik;
// potem pierwsza próba albo najlepsza, zależnie od zasady w Sterowaniu.
function freeRanking_(c) {
  c = c || cfg_();
  var by = {};
  freeRows_(c).forEach(function (r) {
    var p = by[r.key] || (by[r.key] = { name: r.name, order: [], att: {} });
    if (!p.att[r.att]) { p.att[r.att] = { pts: r.pts, good: r.good }; p.order.push(r.att); }
    else if (r.pts >= p.att[r.att].pts) p.att[r.att] = { pts: r.pts, good: r.good };
  });
  return Object.keys(by).map(function (k) {
    var p = by[k], pick = p.att[p.order[0]];
    if (c.freeBest) p.order.forEach(function (a) { if (p.att[a].pts > pick.pts) pick = p.att[a]; });
    return { name: p.name, pts: pick.pts, good: pick.good, tries: p.order.length };
  }).sort(function (a, b) { return b.pts - a.pts; });
}

function akey_(st, name, i) { return 'a|' + st.game + '|' + name + '|' + i + ((st.round || 1) > 1 ? '|r' + st.round : ''); }
// Rozkład odpowiedzi: liczby głosów (A–D, Prawda/Fałsz) albo lista wartości (Liczba).
function dist_(q, rows) {
  if (q.kind === 'num') return rows.map(function (r) { return Number(r[4]); }).filter(isFinite).slice(0, 400);
  var d = q.o.map(function () { return 0; });
  rows.forEach(function (r) { var k = Number(r[4]); if (k >= 0 && k < d.length) d[k]++; });
  return d;
}
// Ostatnia odpowiedź każdego gracza w danym głosowaniu (1 albo 2) danego pytania.
function roundRows_(game, i, rnd) {
  var s = sh_(LIVE), last = s.getLastRow(), by = {};
  if (last < 2) return [];
  s.getRange(2, 1, last - 1, 9).getValues().forEach(function (r) {
    if (r[1] === game && Number(r[3]) === i && (Number(r[8]) || 1) === rnd) by[String(r[2])] = r;
  });
  return Object.keys(by).map(function (k) { return by[k]; });
}
function round1Rows_(game, i) { return roundRows_(game, i, 1); }
function stateOut_(p) {
  var c = cfg_(), st = getState_(), out = {
    mode: c.mode, phone: c.phone, time: c.time, change: c.change, title: c.title, sub: c.sub, now: Date.now(),
    game: st.game, phase: st.phase, i: st.i, n: st.ids.length, startedAt: st.startedAt, round: st.round || 1, peer: c.peer
  };
  var q = null;
  if (st.i >= 0 && st.i < st.ids.length && ['question', 'discuss', 'reveal', 'ranking'].indexOf(st.phase) >= 0) {
    q = questionById_(st.ids[st.i]);
    if (q) out.q = pubQ_(q, c, st.phase === 'reveal' || st.phase === 'ranking');
  }
  var rk = ranking_(st.game, c.streak);
  out.streakOn = c.streak; out.teams = c.teams.length ? c.teams : null;
  if (p.name && st.phase === 'question') {   // potwierdzenie dla telefonu: odpowiedź zapisana na serwerze
    var mine = CacheService.getScriptCache().get(akey_(st, String(p.name).slice(0, 20), st.i));
    if (mine !== null) out.mine = Number(mine);
  }
  if (p.name) {
    var pos = -1, me = null;
    rk.forEach(function (r, k) { if (r.name === p.name) { pos = k + 1; me = r; } });
    out.me = me ? { pts: me.pts, good: me.good, place: pos, of: rk.length, team: me.team, streak: me.streak } : null;
    if (me && me.team) {
      var tr = teamRank_(rk, c) || [];
      tr.forEach(function (t, k) { if (t.team === me.team) out.me.tplace = { place: k + 1, of: tr.length, pts: t.pts }; });
    }
    if (me && st.phase === 'reveal') {
      var mine = liveRows_(st.game).filter(function (r) { return String(r[2]) === p.name && Number(r[3]) === st.i; })[0];
      out.last = mine ? { choice: Number(mine[4]), ok: Number(mine[5]), pts: Number(mine[6]), bonus: me.lb[st.i] || 0 } : null;
    }
    if (st.phase === 'discuss') {
      var m1 = liveRows_(st.game).filter(function (r) { return String(r[2]) === p.name && Number(r[3]) === st.i; })[0];
      out.first = m1 ? Number(m1[4]) : null;
    }
  }
  if (p.full) {
    var rows = liveRows_(st.game);
    out.joined = rk.length;
    out.answered = (st.round || 1) > 1 ? roundRows_(st.game, st.i, st.round).length : rows.filter(function (r) { return Number(r[3]) === st.i; }).length;
    if (st.phase === 'lobby') { out.names = rk.map(function (r) { return r.name; }); out.nteam = rk.map(function (r) { return r.team || ''; }); }
    if (c.teams.length && st.phase !== 'lobby') out.trank = teamRank_(rk, c);
    if ((st.phase === 'reveal' || st.phase === 'discuss') && q) {
      out.dist = dist_(q, rows.filter(function (r) { return Number(r[3]) === st.i; }));
      if ((st.round || 1) === 2 || st.phase === 'discuss') out.dist1 = dist_(q, round1Rows_(st.game, st.i));
      if (st.phase === 'discuss') delete out.dist;
    }
    if (st.phase === 'reveal' || st.phase === 'ranking' || st.phase === 'break' || st.phase === 'final') out.rank = rk.slice(0, 10).map(function (r) { return { name: r.name, pts: r.pts, good: r.good, team: r.team, streak: r.streak }; });
    if (c.mode === 'free') out.freerank = freeRanking_(c).slice(0, 10);
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
    questions: qs.map(function (q) { return pubQ_(q, c, false); }) };
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
    var qs = roundQuestions_(r, c), a = String(p.a || '').split(','), m = String(p.m || '').split(','), pts = 0, good = 0, det = [], gst = 0;
    qs.forEach(function (q, i) {
      var raw = a[i] === undefined ? '' : String(a[i]), ch = raw === '' ? NaN : num_(raw);
      if (!isFinite(ch)) ch = q.kind === 'num' ? null : -1;
      var ms = Math.min(c.time * 1000, Math.max(0, Number(m[i]) || 0));
      var sc = ch === null ? { ok: 0, pts: 0 } : score_(q, ch, ms, c.time), bon = 0;
      if (sc.ok) { good++; gst++; bon = c.streak ? bonus_(gst) : 0; } else gst = 0;
      pts += sc.pts + bon;
      det.push({ bonus: bon, ch: ch, c: q.c, ok: sc.ok, kind: q.kind, val: q.val, tol: q.tol, expl: q.expl, cap: q.cap || mediaFor_(q, c).auto || '' });
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
    .addItem('Sprawdź pytania ze zdjęciem i muzyką', 'checkMedia')
    .addItem('Zestawienie wyników z danymi uczestników', 'idReport')
    .addItem('Zapisz bieżącą grę na żywo w historii', 'saveCurrentGame')
    .addItem('Zakończ sezon Grand Prix i zarchiwizuj', 'archiveGp')
    .addItem('Usuń gracza / zwolnij pseudonim…', 'releaseName')
    .addToUi();
}
// Jednorazowo uruchom w edytorze (Uruchom), aby nadać zgody na Dysk i adresy zewnętrzne (iTunes).
function authorize() {
  DriveApp.getRootFolder();
  UrlFetchApp.fetch('https://itunes.apple.com/search?term=queen&media=music&limit=1');
  SpreadsheetApp.getActive().getName();
}
function checkMedia() {
  var c = cfg_(), lines = [];
  CacheService.getScriptCache().removeAll(['fm']);
  questions_('Wszystkie').forEach(function (q) {
    if (q.typ === 'Tekst' || !q.src) return;
    var m = mediaFor_(q, c);
    lines.push((m.img || m.audio ? 'OK   ' : 'BŁĄD ') + q.typ + ' · ' + q.src + (m.auto ? ' → ' + m.auto : '') + (m.err ? ' · ' + m.err : '') + (m.warn ? ' · UWAGA: ' + m.warn : ''));
  });
  SpreadsheetApp.getUi().alert(lines.length ? lines.join('\n') : 'Brak aktywnych pytań ze zdjęciem lub muzyką (kolumny Typ i Źródło w zakładce Pytania).');
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
    var rows = n.getRange(2, 1, n.getLastRow() - 1, 6).getValues();
    for (var i = rows.length - 1; i >= 0; i--) if (String(rows[i][0]) === 'gp') n.deleteRow(i + 2);
  }
  writeGpSheet_(c);
  CacheService.getScriptCache().remove('gprank');
}
// Zestawienie: kto (dane niewidoczne na ekranie) ile zdobył punktów
function idReport() {
  var c = cfg_(), n = sh_(NAMES), out = [];
  if (!n || n.getLastRow() < 2) { SpreadsheetApp.getUi().alert('Brak zarejestrowanych uczestników.'); return; }
  var free = {}, gp = {}, live = {}, game = getState_().game;
  freeRanking_(c).forEach(function (r) { free[nkey_(r.name)] = r; });
  gpRanking_(c).forEach(function (r) { gp[r.key] = { pts: r.total, good: '' }; });
  liveRank_(game).forEach(function (r) { live[nkey_(r.name)] = r; });
  n.getRange(2, 1, n.getLastRow() - 1, 6).getValues().forEach(function (r) {
    var scope = String(r[0]), key = String(r[1]), sc = scope === 'free' ? free[key] : (scope === 'gp' ? gp[key] : (scope === game ? live[key] : null));
    var mode = scope === 'free' ? 'Dowolny moment' : (scope === 'gp' ? 'Grand Prix' : (scope === game ? 'Na żywo (bieżąca gra)' : 'Na żywo (wcześniejsza gra)'));
    out.push([mode, String(r[2]), String(r[5] || ''), sc ? (sc.pts !== undefined ? sc.pts : sc.total) : '', sc && sc.good !== undefined ? sc.good : '', r[4]]);
  });
  out.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : (Number(b[3]) || 0) - (Number(a[3]) || 0); });
  var s = sh_('Wyniki imienne') || SpreadsheetApp.getActive().insertSheet('Wyniki imienne');
  s.getRange(1, 1, 400, 6).clearContent();
  s.getRange(1, 1, 1, 6).setValues([['Tryb', 'Pseudonim', 'Dane uczestnika', 'Punkty', 'Poprawne', 'Rejestracja']]);
  if (out.length) s.getRange(2, 1, out.length, 6).setValues(out);
  SpreadsheetApp.getUi().alert('Zestawienie zapisano w zakładce "Wyniki imienne" (' + out.length + ' wierszy).');
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
  if (action === 'config') return json_(withMeta_({ win: freeWindow_(c, 0), mode: c.mode, open: c.open, time: c.time, phone: c.phone, category: c.category, change: c.change, flimit: c.freeLimit, fbest: c.freeBest, streak: c.streak, teams: c.teams.length ? c.teams : null }, c));
  if (action === 'questions') {
    return json_(withMeta_({ win: freeWindow_(c, 0), time: c.time, open: c.open, change: c.change, questions: pick_(c, freeSeed_(c)).map(function (q) { return pubQ_(q, c, true); }) }, c));
  }
  if (action === 'state') return json_(withMeta_(stateOut_(p), c));
  if (action === 'pstate') return json_(withMeta_(pStateOut_(p, c), c));
  if (action === 'gp') return json_(withMeta_(gpPublic_(c, String(p.name || '').slice(0, 20)), c));
  if (action === 'gpquestions') return json_(withMeta_(gpQuestions_(c, p), c));
  if (action === 'gpsubmit') return json_(gpSubmit_(c, p));
  if (action === 'claim') {
    var scope = p.scope === 'live' ? getState_().game : (p.scope === 'gp' ? 'gp' : 'free');
    if (!p.tok) return json_({ ok: 0 });
    var cr = claim_(scope, String(p.name || '').slice(0, 20), p.tok, p.id);
    if (cr.ok && scope === 'free') { cr.tries = freeAttempts_(c, nkey_(p.name)).length; cr.limit = c.freeLimit; cr.best = c.freeBest; }
    return json_(cr);
  }
  return json_({ error: 'unknown action' });
}

function doPost(e) {
  var d = JSON.parse(e.postData.contents), type = d.type || 'result', c = cfg_(), cache = CacheService.getScriptCache();
  if (type === 'result') {                                                   // wynik w trybie dowolnym
    if (freeWindow_(c, 300000).status !== 'open') return json_({ ok: 0, closed: 1 });
    if (d.tok && !claim_('free', String(d.name).slice(0, 20), d.tok).ok) return json_({ ok: 0, taken: 1 });
    var od = sh_('Odpowiedzi'), att = String(d.att || ('x' + Date.now())).slice(0, 24);
    var lk = LockService.getScriptLock(); lk.waitLock(8000);
    try {
      var at = freeAttempts_(c, nkey_(d.name));
      if (at.indexOf(att) < 0 && at.length >= c.freeLimit) return json_({ ok: 0, limit: 1 });
      od.appendRow([new Date(), String(d.name).slice(0, 20), Number(d.pts), Number(d.good), String(d.id || '').slice(0, 60), c.title + (c.sub ? ' · ' + c.sub : ''), att]);
    } finally { lk.releaseLock(); }
    try { writeRankSheet_('Tryb dowolny · ' + c.title + (c.sub ? ' · ' + c.sub : '') + ' ' + (c.freeBest ? '(najlepszy wynik gracza)' : '(pierwsza próba gracza)'), freeRanking_(c), uidMap_('free')); } catch (e) {}
    return json_({ ok: 1 });
  }
  if (type === 'host' && String(d.cmd || '').charAt(0) === 'p') return json_(pHost_(d, c));
  if (type === 'host') return json_(host_(d, c));
  if (type === 'pvote' || type === 'popin') return json_(pAnswer_(d, c));
  if (type === 'admin') return json_(admin_(d, c));
  var st = getState_(), name = String(d.name || '').slice(0, 20);
  if (!name) return json_({ ok: 0 });
  if (type === 'join') {
    if (!d.tok || !claim_(st.game, name, d.tok).ok) return json_({ ok: 0, taken: 1 });
    if (c.teams.length && c.teams.indexOf(String(d.team || '')) < 0) return json_({ ok: 0, team: 1 });
    if (!cache.get('j|' + st.game + '|' + name)) {
      cache.put('j|' + st.game + '|' + name, '1', 21600);
      var team = c.teams.indexOf(String(d.team || '')) >= 0 ? String(d.team) : '';
      sh_(LIVE).appendRow([new Date(), st.game, name, -1, -1, 0, 0, 0, '', team]);
      cache.remove('rk');
    }
    return json_({ ok: 1 });
  }
  if (type === 'answer') {
    var i = Number(d.i), choice = num_(d.choice), now = Date.now(), rnd = Number(d.round) || 1;
    if (st.game !== (d.game || st.game) || st.phase !== 'question' || i !== st.i || rnd !== (st.round || 1)) return json_({ ok: 0, late: 1 });
    if (now - st.startedAt > (c.time * 1000) + 2000) return json_({ ok: 0, late: 1 });
    var q = questionById_(st.ids[i]);
    if (!q || !isFinite(choice) || (q.kind !== 'num' && !(choice >= 0 && choice < q.o.length && choice === Math.floor(choice)))) return json_({ ok: 0 });
    var key = akey_(st, name, i);
    // Zapis pod blokadą: równoczesne appendRow bez blokady potrafią nadpisać sobie wiersze.
    // Ponowione wysłanie tej samej odpowiedzi jest potwierdzane (ok), a nie odrzucane.
    var lk = LockService.getScriptLock(); lk.waitLock(10000);
    try {
      var prev = cache.get(key);
      if (prev !== null && Number(prev) === choice) return json_({ ok: 1, dup: 1, choice: choice });
      if (prev !== null && !c.change) return json_({ ok: 0, dup: 1, choice: Number(prev) });
      var ms = Math.min(c.time * 1000, Math.max(Number(d.ms) || 0, now - st.startedAt - 2000, 0));
      var sc = score_(q, choice, ms, c.time);
      sh_(LIVE).appendRow([new Date(), st.game, name, i, choice, sc.ok, sc.pts, ms, st.round || 1]);
      SpreadsheetApp.flush();
      cache.put(key, String(choice), 21600);   // dopiero po udanym zapisie
      cache.remove('rk');
    } finally { lk.releaseLock(); }
    return json_({ ok: 1, choice: choice });
  }
  return json_({ ok: 0 });
}

function host_(d, c) {
  if (c.pin && String(d.pin) !== c.pin) return { ok: 0, error: 'pin' };
  var st = getState_(), now = Date.now(), cmd = d.cmd;
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    if (cmd === 'reset') {
      archiveLive_(st.game, c);   // poprzednia gra trafia do historii; odpowiedzi nie są kasowane
      st = { game: 'g' + now, phase: 'lobby', i: -1, ids: [], startedAt: 0 };
    } else if (cmd === 'start') {
      var ids = pick_(c, 'g' + now).map(function (q) { return q.id; });
      if (!ids.length) return { ok: 0, error: 'brak pytań' };
      st.ids = ids; st.i = 0; st.phase = 'question'; st.startedAt = now; st.round = 1;
    } else if (cmd === 'reveal') {
      if (st.phase === 'question' || st.phase === 'discuss') st.phase = 'reveal';
    } else if (cmd === 'discuss') {           // peer instruction: rozkład głosów bez poprawnej odpowiedzi
      if (st.phase === 'question' && (st.round || 1) === 1) st.phase = 'discuss';
    } else if (cmd === 'revote') {            // drugie głosowanie po dyskusji
      if (st.phase === 'discuss') { st.phase = 'question'; st.round = 2; st.startedAt = now; }
    } else if (cmd === 'break') {
      if (st.phase === 'reveal' || st.phase === 'ranking') st.phase = 'break';
    } else if (cmd === 'ranking') {
      if (st.phase === 'reveal') st.phase = 'ranking';
    } else if (cmd === 'next') {
      if (st.i + 1 < st.ids.length) { st.i++; st.phase = 'question'; st.startedAt = now; st.round = 1; }
      else st.phase = 'final';
    } else if (cmd === 'final') {
      st.phase = 'final';
    }
    setState_(st);
  } finally { lock.releaseLock(); }
  try {
    if (cmd === 'reset') writeRankSheet_('Gra na żywo · czekamy na graczy', [], {});
    else if (cmd !== 'start' && cmd !== 'replay') updateLiveRank_(c, st);
    if (st.phase === 'final') archiveLive_(st.game, c);
  } catch (e) {}
  return { ok: 1, phase: st.phase, i: st.i };
}


// ======================= PANEL #admin =======================
// Ustawienia (karta Sterowanie) i pytania (karta Pytania) edytowane z przeglądarki.
// Dostęp: PIN prowadzącego; po 10 błędnych PIN-ach w ciągu 10 minut panel jest blokowany na 10 minut.
var ADM_FIELDS = 15;   // Pytania: kolumny A–O (zob. QCOLS)

function admin_(d, c) {
  var cache = CacheService.getScriptCache(), fails = Number(cache.get('adm_fail') || 0);
  if (fails >= 10) return { ok: 0, error: 'Zbyt wiele błędnych PIN-ów. Spróbuj ponownie za 10 minut.' };
  if (String(d.pin) !== c.pin || !c.pin) { cache.put('adm_fail', String(fails + 1), 600); return { ok: 0, error: 'pin' }; }
  var op = d.op;
  try {
    if (op === 'load') return admLoad_(c);
    var lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      var r;
      if (op === 'set') r = admSet_(d.changes || []);
      else if (op === 'qsave') r = admQSave_(d);
      else if (op === 'qactive') r = admQActive_(d);
      else if (op === 'qdelete') r = admQShift_(d, 'delete');
      else if (op === 'qmove') r = admQShift_(d, 'move');
      else if (op === 'psave') r = admPSave_(d);
      else if (op === 'pactive') r = admPActive_(d);
      else if (op === 'pdelete') r = admPShift_(d, 'delete');
      else if (op === 'pmove') r = admPShift_(d, 'move');
      else return { ok: 0, error: 'Nieznana operacja' };
      if (!r.ok) return r;
    } finally { lock.releaseLock(); }
    return admLoad_(cfg_());
  } catch (e) { return { ok: 0, error: String(e && e.message || e) }; }
}

function admFmt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  return v === null || v === undefined ? '' : String(v);
}
function admRuleOpts_(rule) {
  if (!rule) return null;
  try {
    var t = String(rule.getCriteriaType()), v = rule.getCriteriaValues(), list = [];
    if (t === 'VALUE_IN_LIST') list = v[0].map(String);
    else if (t === 'VALUE_IN_RANGE') v[0].getValues().forEach(function (r) { r.forEach(function (x) { if (x !== '') list.push(String(x)); }); });
    else return null;
    return { opts: list, strict: !rule.getAllowInvalid() };
  } catch (e) { return null; }
}
function admSettings_() {
  var s = sh_('Sterowanie'), rg = s.getRange('A2:B80'), vals = rg.getValues(), dv = rg.getDataValidations(), out = [], sec = '';
  vals.forEach(function (r, k) {
    var label = String(r[0]).trim();
    if (!label) return;
    if (label.indexOf('▸') === 0) { sec = label.replace(/^▸\s*/, ''); out.push({ section: sec }); return; }
    var o = { label: label, value: admFmt_(r[1]), row: k + 2 }, ru = admRuleOpts_(dv[k][1]);
    if (ru) { o.opts = ru.opts; o.strict = ru.strict; }
    if (/otwarty (od|do)/.test(label)) o.kind = 'datetime';
    if (/^ADRESY/.test(sec)) o.readonly = true;
    out.push(o);
  });
  return out;
}
function admCategories_() {
  var k = sh_('Kategorie'), out = [];
  if (k && k.getLastRow() >= 2) k.getRange(2, 1, k.getLastRow() - 1, 1).getValues().forEach(function (r) { var v = String(r[0]).trim(); if (v && v !== 'Wszystkie') out.push(v); });
  return out;
}
function admQuestions_() {
  var s = sh_('Pytania'), last = s.getLastRow(), out = [];
  if (last < 2) return out;
  s.getRange(2, 1, last - 1, ADM_FIELDS).getValues().forEach(function (r, k) {
    if (r.every(function (x) { return x === ''; })) return;
    out.push({ row: k + 2, v: r.map(admFmt_) });
  });
  return out;
}
function admGameActive_() { var st = getState_(); return ['question', 'reveal', 'ranking', 'break'].indexOf(st.phase) >= 0; }
function admHeaders_() {
  var s = sh_('Pytania'), h = s.getRange(1, 12, 1, QHEAD.length).getValues()[0];
  if (h.some(function (x) { return x === ''; })) s.getRange(1, 12, 1, QHEAD.length).setValues([QHEAD]);
}
function admLoad_(c) {
  try { admHeaders_(); } catch (e) {}
  var st = getState_(), media = [];
  try { media = Object.keys(folderMap_(c)).filter(function (n) { return /\.[a-z0-9]{2,5}$/.test(n); }).sort(); } catch (e) {}
  return { ok: 1, settings: admSettings_(), categories: admCategories_(), questions: admQuestions_(), media: media,
    topics: admPTopics_(), pgame: (function (ps) { return { phase: ps.phase, i: ps.i, n: ps.ids.length }; })(getP_()), puse: ptopics_(c.pset).map(function (t) { return t.id; }),
    game: { phase: st.phase, i: st.i, n: st.ids.length, active: admGameActive_() },
    use: pick_(c, 'g' + Date.now()).map(function (q) { return q.id; }), random: c.random, title: c.title, sub: c.sub };
}

function admSet_(changes) {
  var all = admSettings_(), by = {};
  all.forEach(function (o) { if (o.label) by[o.label] = o; });
  var s = sh_('Sterowanie');
  for (var i = 0; i < changes.length; i++) {
    var ch = changes[i], o = by[ch.label], v = String(ch.value === undefined ? '' : ch.value).trim();
    if (!o || o.readonly) return { ok: 0, error: 'Nie można zmienić ustawienia: ' + ch.label };
    if (o.strict && o.opts && o.opts.indexOf(v) < 0) return { ok: 0, error: 'Niedozwolona wartość „' + v + '” dla: ' + ch.label };
    if (o.label === 'PIN prowadzącego' && !/^\d{4,8}$/.test(v)) return { ok: 0, error: 'PIN musi mieć od 4 do 8 cyfr.' };
    if (o.kind === 'datetime' && v && !/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(v)) return { ok: 0, error: 'Data w formacie RRRR-MM-DD GG:MM: ' + ch.label };
  }
  changes.forEach(function (ch) { s.getRange(by[ch.label].row, 2).setValue(String(ch.value === undefined ? '' : ch.value).trim()); });
  CacheService.getScriptCache().removeAll(['rk', 'st', 'stfull', 'fm']);
  return { ok: 1 };
}

// Ochrona przed edycją niewłaściwego wiersza: klient przesyła treść pytania, którą widział.
function admCheckRow_(s, row, orig) {
  row = Number(row);
  if (!(row >= 2) || row > s.getLastRow()) return 'Pytanie nie istnieje (wiersz ' + row + ').';
  if (String(s.getRange(row, 2).getValue()) !== String(orig)) return 'Pytanie w wierszu ' + row + ' zmieniło się w międzyczasie. Odśwież listę.';
  return '';
}
function admQSave_(d) {
  var s = sh_('Pytania'), v = (d.v || []).slice(0, ADM_FIELDS).map(function (x) { return String(x === undefined ? '' : x).trim(); });
  while (v.length < ADM_FIELDS) v.push('');
  if (!v[0]) return { ok: 0, error: 'Podaj kategorię.' };
  if (!v[1]) return { ok: 0, error: 'Podaj treść pytania.' };
  var kind = kind_(v[11]);
  v[11] = kind === 'tf' ? 'Prawda/Fałsz' : (kind === 'num' ? 'Liczba' : 'A–D');
  if (kind === 'num') {
    if (!isFinite(num_(v[13]))) return { ok: 0, error: 'Podaj poprawną wartość liczbową.' };
    if (v[14] !== '' && !(num_(v[14]) >= 0)) return { ok: 0, error: 'Tolerancja musi być liczbą nieujemną.' };
    v[13] = String(num_(v[13])); v[14] = v[14] === '' ? '' : String(num_(v[14]));
    v[2] = v[3] = v[4] = v[5] = v[6] = '';
  } else {
    v[13] = v[14] = '';
    if (kind === 'tf') { v[2] = v[2] || 'Prawda'; v[3] = v[3] || 'Fałsz'; v[4] = v[5] = ''; }
    else for (var k = 2; k <= 5; k++) if (!v[k]) return { ok: 0, error: 'Uzupełnij odpowiedź ' + 'ABCD'[k - 2] + '.' };
    v[6] = v[6].toUpperCase();
    if ((kind === 'tf' ? 'AB' : 'ABCD').indexOf(v[6]) < 0 || v[6].length !== 1) return { ok: 0, error: 'Wskaż poprawną odpowiedź.' };
  }
  v[7] = v[7].toUpperCase() === 'NIE' ? 'NIE' : 'TAK';
  if (['Tekst', 'Zdjęcie', 'Muzyka'].indexOf(v[8]) < 0) v[8] = 'Tekst';
  if (v[8] !== 'Tekst' && !v[9]) return { ok: 0, error: 'Pytanie typu ' + v[8] + ' wymaga źródła (plik, adres albo „wykonawca – tytuł”).' };
  var row = Number(d.row) || 0;
  if (row) { var e = admCheckRow_(s, row, d.orig); if (e) return { ok: 0, error: e }; }
  else {
    var after = Number(d.after) || 0;
    if (after >= 2 && after < s.getLastRow()) {
      if (admGameActive_()) return { ok: 0, error: 'W trakcie gry na żywo nowe pytanie można dodać tylko na końcu listy.' };
      s.insertRowAfter(after); row = after + 1;
    } else row = s.getLastRow() + 1;
  }
  // tekst zaczynający się od = lub + nie może zostać potraktowany jak formuła
  s.getRange(row, 1, 1, ADM_FIELDS).setValues([v.map(function (x) { return /^[=+]/.test(x) ? "'" + x : x; })]);
  admAddCategory_(v[0]);
  return { ok: 1, row: row };
}
function admAddCategory_(cat) {
  if (admCategories_().indexOf(cat) >= 0) return;
  var k = sh_('Kategorie'); if (!k) return;
  if (k.getLastRow() < 20) k.getRange(k.getLastRow() + 1, 1).setValue(cat);   // walidacja w arkuszu sięga A20
}
function admQActive_(d) {
  var s = sh_('Pytania'), items = d.items || [], val = String(d.value).toUpperCase() === 'NIE' ? 'NIE' : 'TAK';
  for (var i = 0; i < items.length; i++) { var e = admCheckRow_(s, items[i].row, items[i].orig); if (e) return { ok: 0, error: e }; }
  items.forEach(function (it) { s.getRange(Number(it.row), 8).setValue(val); });
  return { ok: 1 };
}
// Usuwanie i przesuwanie zmienia numery wierszy (= id pytań), więc jest zablokowane w trakcie gry na żywo.
function admQShift_(d, kind) {
  if (admGameActive_()) return { ok: 0, error: 'Trwa gra na żywo. Usuwanie i zmiana kolejności pytań są możliwe w lobby albo po wyniku końcowym.' };
  var s = sh_('Pytania'), row = Number(d.row), e = admCheckRow_(s, row, d.orig);
  if (e) return { ok: 0, error: e };
  if (kind === 'delete') { s.deleteRow(row); return { ok: 1 }; }
  // przeniesienie pytania na miejsce wiersza „to” (pozostałe przesuwają się o jeden)
  var to = Number(d.to) || (row + (Number(d.dir) < 0 ? -1 : 1));
  if (to < 2 || to > s.getLastRow() || to === row) return { ok: 1 };
  var lo = Math.min(row, to), n = Math.abs(to - row) + 1, rg = s.getRange(lo, 1, n, ADM_FIELDS), b = rg.getValues();
  if (row < to) b.push(b.shift()); else b.unshift(b.pop());
  rg.setValues(b);
  return { ok: 1 };
}


// ======================= GŁOSOWANIE I OPINIE =======================
// Karta „Ankiety”: A Zestaw, B Rodzaj (Głosowanie / Opinia), C Temat, D–I Opcje 1–6 (głosowanie), J Aktywne.
// Karta „Odpowiedzi ankiet”: Czas, Sesja, Temat (nr strony), Kod urządzenia, Rodzaj (G/O), Wartość, Id opinii, Ukryta.
// Uczestnicy są anonimowi: na ekranie nie ma pseudonimów, urządzenie rozpoznaje losowy kod.
var PTOP = 'Ankiety', PANS = 'Odpowiedzi ankiet', PCOLS = 10, PMAXLEN = 280, PMAXOP = 20;
var PHEAD = ['Zestaw', 'Rodzaj (Głosowanie / Opinia)', 'Temat', 'Opcja 1', 'Opcja 2', 'Opcja 3', 'Opcja 4', 'Opcja 5', 'Opcja 6', 'Aktywne'];

function pSheet_() {
  var s = sh_(PTOP);
  if (!s) { s = SpreadsheetApp.getActive().insertSheet(PTOP); s.getRange(1, 1, 1, PCOLS).setValues([PHEAD]); s.setFrozenRows(1); }
  return s;
}
function pkind_(v) { return String(v || '').toLowerCase().indexOf('opin') >= 0 ? 'opin' : 'vote'; }
function ptopics_(set) {
  var s = sh_(PTOP), out = [];
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, PCOLS).getValues().forEach(function (r, k) {
    if (!r[2] || String(r[9]).toUpperCase() === 'NIE') return;
    if (set && set !== 'Wszystkie' && String(r[0]).trim() !== set) return;
    var t = { id: k + 2, set: String(r[0]).trim(), kind: pkind_(r[1]), t: String(r[2]).trim(),
      o: r.slice(3, 9).map(String).map(function (x) { return x.trim(); }).filter(function (x) { return x; }) };
    if (t.kind === 'vote' && t.o.length < 2) return;
    out.push(t);
  });
  return out;
}
function ptopicById_(id) { var a = ptopics_('Wszystkie'); for (var i = 0; i < a.length; i++) if (a[i].id === id) return a[i]; return null; }
function getP_() { var v = PropertiesService.getScriptProperties().getProperty('pstate'); return v ? JSON.parse(v) : { sid: 'p0', phase: 'lobby', i: -1, ids: [] }; }
function setP_(ps) { PropertiesService.getScriptProperties().setProperty('pstate', JSON.stringify(ps)); }

// Odpowiedzi bieżącej strony (z pamięci podręcznej na 2 s): [kod, rodzaj, wartość, id, ukryta]
function pRows_(sid, i) {
  var cache = CacheService.getScriptCache(), ck = 'pr|' + sid + '|' + i, hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  var s = sh_(PANS), out = [];
  if (s && s.getLastRow() >= 2) s.getRange(2, 1, s.getLastRow() - 1, 8).getValues().forEach(function (r) {
    if (String(r[1]) === sid && Number(r[2]) === i) out.push([String(r[3]), String(r[4]), String(r[5]), String(r[6]), String(r[7]) === 'TAK' ? 1 : 0]);
  });
  try { cache.put(ck, JSON.stringify(out), 2); } catch (e) {}
  return out;
}
function pAgg_(t, rows) {
  if (t.kind === 'vote') {
    var last = {}, d = t.o.map(function () { return 0; });
    rows.forEach(function (r) { if (r[1] === 'G') last[r[0]] = Number(r[2]); });
    Object.keys(last).forEach(function (k) { var v = last[k]; if (v >= 0 && v < d.length) d[v]++; });
    return { counts: d, voters: Object.keys(last).length };
  }
  var ops = rows.filter(function (r) { return r[1] === 'O' && !r[4]; }).map(function (r) { return { id: r[3], text: r[2] }; });
  return { opinions: ops.slice(-300), authors: Object.keys(rows.reduce(function (m, r) { if (r[1] === 'O') m[r[0]] = 1; return m; }, {})).length, total: ops.length };
}
function pStateOut_(p, c) {
  var ps = getP_(), out = { mode: c.mode, sid: ps.sid, phase: ps.phase, i: ps.i, n: ps.ids.length, now: Date.now() };
  if (ps.phase === 'topic' && ps.i >= 0 && ps.i < ps.ids.length) {
    var t = ptopicById_(ps.ids[ps.i]);
    if (t) {
      out.t = { kind: t.kind, t: t.t, o: t.o };
      var rows = pRows_(ps.sid, ps.i), agg = pAgg_(t, rows);
      if (p.full) out.agg = agg;
      else out.agg = t.kind === 'vote' ? { voters: agg.voters } : { authors: agg.authors };
      if (p.tok) {
        var mine = rows.filter(function (r) { return r[0] === String(p.tok); });
        if (t.kind === 'vote') { var g = mine.filter(function (r) { return r[1] === 'G'; }).pop(); out.mine = g ? Number(g[2]) : null; }
        else out.mine = mine.filter(function (r) { return r[1] === 'O'; }).length;
      }
    }
  }
  return out;
}
function pAnswer_(d, c) {
  var ps = getP_(), tok = String(d.tok || '').slice(0, 60);
  if (!tok) return { ok: 0 };
  if (ps.phase !== 'topic' || String(d.sid) !== ps.sid || Number(d.i) !== ps.i) return { ok: 0, late: 1 };
  var t = ptopicById_(ps.ids[ps.i]);
  if (!t) return { ok: 0 };
  var cache = CacheService.getScriptCache(), lk = LockService.getScriptLock(), row;
  if (d.type === 'pvote') {
    var ch = Number(d.choice);
    if (t.kind !== 'vote' || !(ch >= 0 && ch < t.o.length && ch === Math.floor(ch))) return { ok: 0 };
    row = [new Date(), ps.sid, ps.i, tok, 'G', ch, '', ''];
  } else {
    var text = String(d.text || '').replace(/\s+/g, ' ').trim().slice(0, PMAXLEN), cid = String(d.cid || '').slice(0, 40);
    if (t.kind !== 'opin' || !text || !cid) return { ok: 0 };
    if (cache.get('pc|' + cid)) return { ok: 1, dup: 1 };                     // ponowienie tej samej opinii
    if (pRows_(ps.sid, ps.i).filter(function (r) { return r[0] === tok && r[1] === 'O'; }).length >= PMAXOP) return { ok: 0, limit: 1 };
    row = [new Date(), ps.sid, ps.i, tok, 'O', /^[=+\-@]/.test(text) ? "'" + text : text, cid, ''];
  }
  lk.waitLock(10000);
  try {
    var s = sh_(PANS);
    if (!s) { s = SpreadsheetApp.getActive().insertSheet(PANS); s.appendRow(['Czas', 'Sesja', 'Temat (nr)', 'Kod urządzenia', 'Rodzaj', 'Wartość', 'Id opinii', 'Ukryta']); s.setFrozenRows(1); }
    s.appendRow(row);
    SpreadsheetApp.flush();
    if (d.type === 'popin') cache.put('pc|' + row[6], '1', 21600);
    cache.remove('pr|' + ps.sid + '|' + ps.i);
  } finally { lk.releaseLock(); }
  return { ok: 1 };
}
function pHost_(d, c) {
  if (c.pin && String(d.pin) !== c.pin) return { ok: 0, error: 'pin' };
  var ps = getP_(), cmd = d.cmd, lk = LockService.getScriptLock(); lk.waitLock(5000);
  try {
    if (cmd === 'pstart') {
      var ids = ptopics_(c.pset).map(function (t) { return t.id; });
      if (!ids.length) return { ok: 0, error: 'Brak aktywnych tematów w zestawie „' + c.pset + '”.' };
      ps = { sid: 'p' + Date.now(), phase: 'topic', i: 0, ids: ids };
    } else if (cmd === 'pnext') { if (ps.phase === 'topic') { if (ps.i + 1 < ps.ids.length) ps.i++; else ps.phase = 'end'; } }
    else if (cmd === 'pprev') { if (ps.phase === 'end') ps.phase = 'topic'; else if (ps.i > 0) ps.i--; }
    else if (cmd === 'pend') ps.phase = 'end';
    else if (cmd === 'preset') ps = { sid: 'p' + Date.now(), phase: 'lobby', i: -1, ids: [] };
    else if (cmd === 'phide') {   // moderacja: ukrycie opinii z ekranu (zostaje w arkuszu z oznaczeniem)
      var s = sh_(PANS), oid = String(d.oid || '');
      if (s && oid && s.getLastRow() >= 2) {
        var col = s.getRange(2, 7, s.getLastRow() - 1, 1).getValues();
        for (var k = 0; k < col.length; k++) if (String(col[k][0]) === oid) { s.getRange(k + 2, 8).setValue('TAK'); break; }
      }
      CacheService.getScriptCache().remove('pr|' + ps.sid + '|' + ps.i);
      return { ok: 1 };
    } else return { ok: 0 };
    setP_(ps);
  } finally { lk.releaseLock(); }
  return { ok: 1, phase: ps.phase, i: ps.i };
}

// --- panel: tematy głosowań i opinii ---
function admPTopics_() {
  var s = pSheet_(), out = [];
  if (s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, PCOLS).getValues().forEach(function (r, k) {
    if (r.every(function (x) { return x === ''; })) return;
    out.push({ row: k + 2, v: r.map(admFmt_) });
  });
  return out;
}
function admPCheck_(s, row, orig) {
  row = Number(row);
  if (!(row >= 2) || row > s.getLastRow()) return 'Temat nie istnieje (wiersz ' + row + ').';
  if (String(s.getRange(row, 3).getValue()) !== String(orig)) return 'Temat w wierszu ' + row + ' zmienił się w międzyczasie. Odśwież listę.';
  return '';
}
function pActive_() { return getP_().phase === 'topic'; }
function admPSave_(d) {
  var s = pSheet_(), v = (d.v || []).slice(0, PCOLS).map(function (x) { return String(x === undefined ? '' : x).trim(); });
  while (v.length < PCOLS) v.push('');
  v[1] = pkind_(v[1]) === 'opin' ? 'Opinia' : 'Głosowanie';
  v[9] = v[9].toUpperCase() === 'NIE' ? 'NIE' : 'TAK';
  if (!v[0]) v[0] = 'Ogólny';
  if (!v[2]) return { ok: 0, error: 'Podaj temat.' };
  if (v[1] === 'Głosowanie') {
    var opts = v.slice(3, 9).filter(function (x) { return x; });
    if (opts.length < 2) return { ok: 0, error: 'Głosowanie wymaga co najmniej dwóch opcji.' };
    for (var k = 3; k < 9; k++) v[k] = opts[k - 3] || '';
  } else for (var j = 3; j < 9; j++) v[j] = '';
  var row = Number(d.row) || 0;
  if (row) { var e = admPCheck_(s, row, d.orig); if (e) return { ok: 0, error: e }; }
  else {
    var after = Number(d.after) || 0;
    if (after >= 2 && after < s.getLastRow()) {
      if (pActive_()) return { ok: 0, error: 'Trwa sesja głosowania: nowy temat można dodać tylko na końcu listy.' };
      s.insertRowAfter(after); row = after + 1;
    } else row = s.getLastRow() + 1;
  }
  s.getRange(row, 1, 1, PCOLS).setValues([v.map(function (x) { return /^[=+]/.test(x) ? "'" + x : x; })]);
  return { ok: 1 };
}
function admPActive_(d) {
  var s = pSheet_(), e = admPCheck_(s, d.row, d.orig);
  if (e) return { ok: 0, error: e };
  s.getRange(Number(d.row), 10).setValue(String(d.value).toUpperCase() === 'NIE' ? 'NIE' : 'TAK');
  return { ok: 1 };
}
function admPShift_(d, kind) {
  if (pActive_()) return { ok: 0, error: 'Trwa sesja głosowania. Usuwanie i zmiana kolejności tematów są możliwe przed startem albo po zakończeniu.' };
  var s = pSheet_(), row = Number(d.row), e = admPCheck_(s, row, d.orig);
  if (e) return { ok: 0, error: e };
  if (kind === 'delete') { s.deleteRow(row); return { ok: 1 }; }
  var to = Number(d.to) || row;
  if (to < 2 || to > s.getLastRow() || to === row) return { ok: 1 };
  var lo = Math.min(row, to), rg = s.getRange(lo, 1, Math.abs(to - row) + 1, PCOLS), b = rg.getValues();
  if (row < to) b.push(b.shift()); else b.unshift(b.pop());
  rg.setValues(b);
  return { ok: 1 };
}
