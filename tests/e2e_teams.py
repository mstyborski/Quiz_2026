# Test bonusu za serię poprawnych odpowiedzi i drużyn (gra na żywo) oraz bonusu w trybie dowolnym.
import json, re, sys, time, urllib.request, urllib.parse
from playwright.sync_api import sync_playwright

B = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:8787"
def api(p): return json.loads(urllib.request.urlopen(B + p).read())
def setv(l, v): api("/__set?" + urllib.parse.urlencode({"label": l, "value": v}))
def post(o):
    r = urllib.request.Request(B + "/exec", data=json.dumps(o).encode(), method="POST"); return json.loads(urllib.request.urlopen(r).read())
def host(c): return post({"type": "host", "cmd": c, "pin": "1213"})
def join(n, team):
    api("/exec?action=claim&scope=live&" + urllib.parse.urlencode({"name": n, "tok": "t" + n})); return post({"type": "join", "name": n, "tok": "t" + n, "team": team})
def ans(n, i, ch): return post({"type": "answer", "name": n, "i": i, "choice": ch, "round": 1, "ms": 2000})
def st(name=""): return api("/exec?action=state&full=1" + ("&name=" + urllib.parse.quote(name) if name else ""))
results = []
def check(name, cond):
    results.append((name, bool(cond))); print(("OK  " if cond else "BŁĄD") + " " + name)
EXT = re.compile(r"^https?://(?!localhost|127\.0\.0\.1)")

api("/__reset")
setv("Bonus za serię poprawnych odpowiedzi", "Tak"); setv("Drużyny (nazwy po przecinku)", "Czerwoni, Niebiescy"); setv("Liczba pytań w quizie", "5")
with sync_playwright() as p:
    brs, errs = [], []
    def page(w, h):
        br = p.chromium.launch(); brs.append(br); ctx = br.new_context(viewport={"width": w, "height": h}); ctx.route(EXT, lambda r: r.abort())
        pg = ctx.new_page(); pg.on("pageerror", lambda e: errs.append(str(e))); pg.on("dialog", lambda d: d.accept()); return pg
    tv = page(1280, 800); tv.goto(B + "/v2/#tv")
    ala = page(390, 800); ala.goto(B + "/v2/"); ala.wait_for_selector("#teamSel")
    check("telefon: wybór drużyny przy dołączaniu", ala.locator("#teamSel option").count() == 3)
    ala.fill("#nm", "Ala"); ala.click("#go"); time.sleep(0.5)
    check("bez wybranej drużyny nie można dołączyć", ala.locator("#nm").count() == 1)
    ala.select_option("#teamSel", "Czerwoni"); ala.click("#go"); ala.wait_for_selector("text=Cześć, Ala")
    check("join z nieznaną drużyną odrzucony", join("Zenon", "Zieloni").get("team") == 1)
    join("Bartek", "Czerwoni"); join("Celina", "Niebiescy"); time.sleep(3)
    check("TV: lobby pogrupowane według drużyn", "Czerwoni (2)" in tv.inner_text("body") and "Niebiescy (1)" in tv.inner_text("body"))
    host("start")
    plan = {"Ala": [1, 2, 3], "Bartek": [1, 0, 3], "Celina": [1, 2, 3]}   # poprawne: B, C, D
    for q in range(3):
        if q: host("next")
        time.sleep(0.3)
        for n, a in plan.items():
            if n == "Ala" and q == 0: ala.wait_for_selector("button.tile"); ala.click('button.tile[data-i="1"]'); continue
            ans(n, q, a[q])
        time.sleep(1.5); host("reveal"); time.sleep(2.5)
    s = st("Ala"); rk = {r["name"]: r for r in s["rank"]}
    rows = [r for r in api("/__sheet?n=Odpowiedzi%20live")[1:] if r[3] >= 0]
    base = {n: sum(r[6] for r in rows if r[2] == n) for n in plan}
    check("bonus Ala: +100 i +200 = 300", rk["Ala"]["pts"] - base["Ala"] == 300 and s["last"]["bonus"] == 200)
    check("Bartek: seria przerwana, bonus 0", rk["Bartek"]["pts"] == base["Bartek"])
    check("seria Celiny = 3", rk["Celina"]["streak"] == 3)
    tr = {t["team"]: t for t in s["trank"]}
    check("ranking drużyn: średnia Czerwonych", tr["Czerwoni"]["pts"] == round((rk["Ala"]["pts"] + rk["Bartek"]["pts"]) / 2) and tr["Czerwoni"]["n"] == 2)
    body = ala.inner_text("body")
    check("telefon: bonus i seria po odsłonięciu", "bonus za serię +200" in body and "Seria poprawnych odpowiedzi: 3" in body)
    check("telefon: miejsce drużyny", "Drużyna Czerwoni: miejsce" in body)
    host("ranking"); time.sleep(2.5)
    check("TV: ranking drużyn i seria 🔥 przy graczu", "Drużyny" in tv.inner_text("body") and "🔥 3" in tv.inner_text("body"))
    tv.screenshot(path="/tmp/teams_tv.png")
    setv("Wynik drużyny", "Suma"); time.sleep(2.5)
    tr2 = {t["team"]: t for t in st()["trank"]}
    check("wynik drużyny jako suma", tr2["Czerwoni"]["pts"] == rk["Ala"]["pts"] + rk["Bartek"]["pts"])
    # tryb dowolny: bonus po stronie telefonu
    setv("Tryb gry", "Dowolny moment"); setv("Tryb dowolny: otwarty od", ""); setv("Tryb dowolny: otwarty do", ""); setv("Zmiana odpowiedzi w czasie pytania", "Nie")
    f = page(390, 800); f.goto(B + "/v2/"); f.fill("#nm", "Zuzia"); f.click("#go")
    for q, a in enumerate([1, 2]):
        f.wait_for_selector("button.tile"); f.click('button.tile[data-i="%d"]' % a); f.wait_for_selector("#nx")
        if q == 1: check("dowolny: bonus za serię w komunikacie", "w tym bonus za serię +100" in f.inner_text("body"))
        f.click("#nx")
    check("brak błędów JavaScript", not errs)
    if errs: print(errs)
    for br in brs: br.close()
bad = [n for n, ok in results if not ok]
print("\n%d/%d testów zaliczonych" % (len(results) - len(bad), len(results)))
sys.exit(1 if bad else 0)
