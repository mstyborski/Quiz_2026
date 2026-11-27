# Test gry na żywo na mocku (node tests/server.js 8787).
# Sprawdza: dołączenie, podwójne kliknięcie, utratę żądania z odpowiedzią (awaria serwera), wynik po odsłonięciu.
import json, re, sys, time, urllib.request
from playwright.sync_api import sync_playwright

B = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:8787"
def api(p): return json.loads(urllib.request.urlopen(B + p).read())
def host(cmd):
    req = urllib.request.Request(B + "/exec", data=json.dumps({"type": "host", "cmd": cmd, "pin": "1213"}).encode(), method="POST")
    return json.loads(urllib.request.urlopen(req).read())

def block_cdn(ctx):
    # tylko zewnętrzne adresy przechodzą przez Pythona; lokalne żądania nie czekają na time.sleep
    ctx.route(re.compile(r"^https?://(?!localhost|127\.0\.0\.1)"), lambda r: r.abort())

results = []
def check(name, cond):
    results.append((name, bool(cond))); print(("OK  " if cond else "BŁĄD") + " " + name)

api("/__reset")
with sync_playwright() as p:
    brs = []
    players = {}
    for nm in ["Ala", "Bartek", "Celina"]:
        br = p.chromium.launch(); brs.append(br)   # osobne przeglądarki: brak dławienia timerów kart w tle
        ctx = br.new_context(viewport={"width": 390, "height": 800}, has_touch=True); block_cdn(ctx)
        pg = ctx.new_page(); pg.goto(B + "/v2/")
        pg.fill("#nm", nm); pg.click("#go"); players[nm] = pg
    time.sleep(4)
    check("trzech graczy w lobby", api("/exec?action=state&full=1")["joined"] == 3)
    host("start"); time.sleep(2.5)
    q = players["Ala"]
    q.wait_for_selector("button.tile")
    # Ala: podwójne kliknięcie tej samej odpowiedzi (A)
    q.dblclick('button.tile[data-i="0"]')
    # Bartek: pierwsze żądanie z odpowiedzią ginie (awaria serwera), kliknięcie B
    api("/__fail?n=1")
    players["Bartek"].click('button.tile[data-i="1"]')
    # Celina: szybkie dwa stuknięcia w C
    c = players["Celina"]; c.tap('button.tile[data-i="2"]'); c.tap('button.tile[data-i="2"]')
    time.sleep(5)
    rows = [r for r in api("/__sheet?n=Odpowiedzi%20live")[1:] if r[3] == 0]
    who = sorted(set(r[2] for r in rows))
    check("odpowiedzi wszystkich trzech graczy zapisane (" + ",".join(who) + ")", who == ["Ala", "Bartek", "Celina"])
    for nm, pg in players.items():
        txt = pg.inner_text("body")
        check(nm + ": telefon potwierdza zapis", "Odpowiedź zapisana" in txt)
    host("reveal"); time.sleep(3)
    for nm, pg in players.items():
        txt = pg.inner_text("body")
        check(nm + ": po odsłonięciu nie ma „Brak odpowiedzi”", "Brak odpowiedzi" not in txt)
    # pytanie 1 (indeks 0) ma poprawną odpowiedź B (n=1 -> 'ABCD'[1])
    check("Bartek: Dobrze!", "Dobrze" in players["Bartek"].inner_text("body"))
    for br in brs: br.close()

bad = [n for n, ok in results if not ok]
print("\n%d/%d testów zaliczonych" % (len(results) - len(bad), len(results)))
sys.exit(1 if bad else 0)
