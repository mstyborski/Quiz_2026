// Lokalny mock Google Apps Script: uruchamia prawdziwy v2/apps_script.gs w Node
// na arkuszu w pamięci i serwuje v2/index.html z podmienionym SCRIPT_URL.
// Użycie: node tests/server.js [port]   (dane startowe: tests/fixture.json)
'use strict';
const http = require('http'), fs = require('fs'), path = require('path'), vm = require('vm');
const ROOT = path.join(__dirname, '..'), PORT = Number(process.argv[2]) || 8787;

// ---------- arkusz w pamięci ----------
function colNum(s) { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
class Rule {
  constructor(r, ss) { this.r = r; this.ss = ss; }
  getCriteriaType() { return this.r.list ? 'VALUE_IN_LIST' : 'VALUE_IN_RANGE'; }
  getCriteriaValues() { return this.r.list ? [this.r.list, true] : [this.ss.range_(this.r.range), true]; }
  getAllowInvalid() { return !this.r.strict; }
}
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const v = (this.sh.d[this.r - 1 + i] || [])[this.c - 1 + j]; row.push(v === undefined || v === null ? '' : v); } o.push(row); } return o; }
  getValue() { return this.getValues()[0][0]; }
  setValues(v) { v.forEach((row, i) => row.forEach((x, j) => this.sh.put(this.r + i, this.c + j, x))); return this; }
  setValue(x) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.put(this.r + i, this.c + j, x); return this; }
  clearContent() { return this.setValue(''); }
  getDataValidations() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const k = (this.r + i) + ':' + (this.c + j), r = this.sh.dv[k]; row.push(r ? new Rule(r, this.sh.ss) : null); } o.push(row); } return o; }
  getRow() { return this.r; }
}
class Sheet {
  constructor(ss, name, d, dv) { this.ss = ss; this.name = name; this.d = d || []; this.dv = dv || {}; }
  getName() { return this.name; }
  put(r, c, x) { while (this.d.length < r) this.d.push([]); const row = this.d[r - 1]; while (row.length < c) row.push(''); row[c - 1] = x; }
  getLastRow() { for (let i = this.d.length; i > 0; i--) if ((this.d[i - 1] || []).some(v => v !== '' && v !== undefined && v !== null)) return i; return 0; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { const m = a.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/); const c1 = colNum(m[1]), r1 = +m[2], c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? +m[4] : r1; return new Range(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1); }
    return new Range(this, a, b, c || 1, d || 1);
  }
  appendRow(v) { this.put(this.getLastRow() + 1, 1, ''); this.getRange(this.getLastRow() + 1, 1, 1, v.length).setValues([v]); return this; }
  deleteRow(r) { this.d.splice(r - 1, 1); return this; }
  insertRowAfter(r) { this.d.splice(r, 0, []); return this; }
  setFrozenRows() { return this; }
}
class SS {
  constructor(fx) { this.sheets = {}; for (const [n, s] of Object.entries(fx.sheets)) this.sheets[n] = new Sheet(this, n, JSON.parse(JSON.stringify(s.data)), s.dv || {}); }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { return (this.sheets[n] = new Sheet(this, n)); }
  range_(a1) { const m = a1.replace(/\$/g, '').replace(/^=/, '').match(/^(.+)!(.+)$/); return this.sheets[m[1]].getRange(m[2]); }
  toast() {}
}

// ---------- usługi Apps Script ----------
function makeEnv(fx) {
  const ss = new SS(fx), props = {}, cache = {};
  const stats = { answerPosts: 0, failNext: 0 };
  const env = {
    SpreadsheetApp: { flush() {}, getActive: () => ss, getUi: () => ({ createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }), alert() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); }, deleteProperty: k => { delete props[k]; } }) },
    CacheService: { getScriptCache: () => ({
      get: k => { const e = cache[k]; if (!e || e.t < Date.now()) return null; return e.v; },
      put: (k, v, s) => { cache[k] = { v: String(v), t: Date.now() + (s || 600) * 1000 }; },
      remove: k => { delete cache[k]; }, removeAll: ks => ks.forEach(k => delete cache[k]) }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, tryLock() { return true; }, releaseLock() {} }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: t => ({ t, setMimeType() { return this; } }) },
    Utilities: { formatDate: (d, tz, f) => { const p = n => String(n).padStart(2, '0'); return f.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate())).replace('HH', p(d.getHours())).replace('mm', p(d.getMinutes())); } },
    Session: { getScriptTimeZone: () => 'Europe/Warsaw' },
    DriveApp: { Access: {}, Permission: {}, getFolderById: () => { throw new Error('brak Dysku w mocku'); }, getFileById: () => { throw new Error('brak'); } },
    UrlFetchApp: { fetch: () => ({ getContentText: () => '{"results":[]}' }) },
    console
  };
  env.__ss = ss; env.__props = props; env.__cache = cache; env.__stats = stats;
  return env;
}

let env;
function load() {
  const fx = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixture.json'), 'utf8'));
  env = makeEnv(fx);
  vm.createContext(env);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'v2', 'apps_script.gs'), 'utf8'), env, { filename: 'apps_script.gs' });
}
load();

function sendJson(res, obj, code) { res.writeHead(code || 200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }); res.end(JSON.stringify(obj)); }

http.createServer((req, res) => {
  const u = new URL(req.url, 'http://x');
  if (u.pathname === '/exec') {
    if (req.method === 'OPTIONS') { res.writeHead(405); res.end(); return; } // jak Apps Script: brak preflight
    if (req.method === 'GET') {
      const p = Object.fromEntries(u.searchParams);
      try { const o = env.doGet({ parameter: p }); sendJson(res, JSON.parse(o.t)); } catch (e) { console.error(e); sendJson(res, { error: String(e) }, 500); }
      return;
    }
    let body = ''; req.on('data', c => body += c); req.on('end', () => {
      try {
        const d = JSON.parse(body);
        if (d.type === 'answer') { env.__stats.answerPosts++; if (env.__stats.failNext > 0) { env.__stats.failNext--; res.writeHead(500); res.end('symulowana awaria'); return; } }
        const o = env.doPost({ postData: { contents: body } }); sendJson(res, JSON.parse(o.t));
      } catch (e) { console.error(e); sendJson(res, { error: String(e) }, 500); }
    });
    return;
  }
  // Pomocnicze endpointy testowe
  if (u.pathname === '/__reset') { load(); sendJson(res, { ok: 1 }); return; }
  if (u.pathname === '/__sheet') { const s = env.__ss.getSheetByName(u.searchParams.get('n')); sendJson(res, s ? s.d : null); return; }
  if (u.pathname === '/__stats') { sendJson(res, env.__stats); return; }
  if (u.pathname === '/__fail') { env.__stats.failNext = Number(u.searchParams.get('n')) || 1; sendJson(res, { ok: 1 }); return; }
  if (u.pathname === '/__set') { // ?label=..&value=..
    const s = env.__ss.getSheetByName('Sterowanie'); s.d.forEach(r => { if (r[0] === u.searchParams.get('label')) r[1] = u.searchParams.get('value'); }); sendJson(res, { ok: 1 }); return; }
  let f = u.pathname === '/' ? '/v2/index.html' : u.pathname;
  if (f.endsWith('/')) f += 'index.html';
  const fp = path.join(ROOT, f);
  if (!fp.startsWith(ROOT) || !fs.existsSync(fp)) { res.writeHead(404); res.end(); return; }
  let txt = fs.readFileSync(fp);
  if (fp.endsWith('.html')) { txt = String(txt).replace(/var SCRIPT_URL="[^"]*"/, 'var SCRIPT_URL="/exec"'); res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }); res.end(txt); return; }
  const ct = { '.svg': 'image/svg+xml', '.png': 'image/png', '.js': 'text/javascript' }[path.extname(fp)];
  res.writeHead(200, ct ? { 'Content-Type': ct } : {}); res.end(txt);
}).listen(PORT, () => console.log('mock na http://localhost:' + PORT + '/v2/'));
