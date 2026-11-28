# Test nowych rodzajów pytań (Prawda/Fałsz, Liczba), wyjaśnienia i ponownego głosowania (peer instruction).
import json, re, sys, time, urllib.request, urllib.parse
from playwright.sync_api import sync_playwright

B = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:8787"
def api(p): return json.loads(urllib.request.urlopen(B + p).read())
def setv(l, v): api("/__set?" + urllib.parse.urlencode({"label": l, "value": v}))
def host(cmd):
    req = urllib.request.Request(B + "/exec", data=json.dumps({"type": "host", "cmd": cmd, "pin": "1213"}).encode(), method="POST")
    return json.loads(urllib.request.urlopen(req).read())
def state(): return api("/exec?action=state&full=1")
results = []
def check(name, cond):
    results.append((name, bool(cond))); print(("OK  " if cond else "BŁĄD") + " " + name)
def wait_phase(ph, t=15):
    for _ in range(t * 4):
        if state()["phase"] == ph: return True
        time.sleep(0.25)
    return False
EXT = re.compile(r"^https?://(?!localhost|127\.0\.0\.1)")

api("/__reset")
setv("Kategoria do gry", "Test rodzajów"); setv("Liczba pytań w quizie", "Wszystkie")
setv("Ponowne głosowanie (peer instruction)", "Tak"); setv("Czas na pytanie (s)", "6")
with sync_playwright() as p:
    brs, errs = [], []
    def page(w, h, init=None):
        br = p.chromium.launch(); brs.append(br); ctx = br.new_context(viewport={"width": w, "height": h})
        ctx.route(EXT, lambda r: r.abort())
        if init: ctx.add_init_script(init)
        pg = ctx.new_page(); pg.on("pageerror", lambda e: errs.append(str(e))); pg.on("dialog", lambda d: d.accept()); return pg
    tv = page(1280, 800, "localStorage.setItem('pin','1213')"); tv.goto(B + "/v2/#tv")
    P = {}
    for nm in ["Ala", "Bartek"]:
        pg = page(390, 800); pg.goto(B + "/v2/"); pg.fill("#nm", nm); pg.click("#go"); P[nm] = pg
    time.sleep(3)
    # --- pytanie 1: Prawda/Fałsz z ponownym głosowaniem ---
    host("start"); P["Ala"].wait_for_selector("button.tile")
    check("P/F: dwa kafelki", P["Ala"].locator("button.tile").count() == 2)
    P["Ala"].click('button.tile[data-i="0"]'); P["Bartek"].click('button.tile[data-i="1"]')
    check("P/F: po czasie TV sam przechodzi do dyskusji", wait_phase("discuss", 12))
    time.sleep(2.5)
    check("dyskusja: TV pokazuje rozkład 1/1", "dyskusja" in tv.inner_text("body").lower() and state()["dist1"] == [1, 1])
    check("dyskusja: telefon pokazuje pierwszą odpowiedź", "Twoja odpowiedź w pierwszym głosowaniu" in P["Bartek"].inner_text("body"))
    host("revote"); P["Bartek"].wait_for_selector("text=ponowne głosowanie")
    time.sleep(0.5); P["Bartek"].click('button.tile[data-i="0"]')
    check("ponowne głosowanie: TV sam odsłania odpowiedź", wait_phase("reveal", 12))
    time.sleep(2.5); st = state()
    check("rozkład po dyskusji [2,0], przed [1,1]", st["dist"] == [2, 0] and st["dist1"] == [1, 1])
    check("TV: liczniki „1 → 2”", "1 → 2" in tv.inner_text("body"))
    check("Bartek: Dobrze! po zmianie zdania", "Dobrze" in P["Bartek"].inner_text("body"))
    check("telefon: wyjaśnienie po odsłonięciu", "Heinego" in P["Ala"].inner_text("body"))
    # --- pytanie 2: Liczba ---
    host("next"); P["Ala"].wait_for_selector("#nv")
    P["Ala"].fill("#nv", "abc"); P["Ala"].click("#nsend")
    check("Liczba: walidacja niepoprawnego wpisu", "Wpisz liczbę" in P["Ala"].inner_text("body"))
    P["Ala"].fill("#nv", "29,5"); P["Ala"].click("#nsend")
    P["Bartek"].wait_for_selector("#nv"); P["Bartek"].fill("#nv", "40"); P["Bartek"].press("#nv", "Enter")
    time.sleep(2.5)
    check("Liczba: telefon potwierdza zapis", "Odpowiedź zapisana" in P["Ala"].inner_text("body"))
    host("reveal"); time.sleep(3)
    rows = [r for r in api("/__sheet?n=Odpowiedzi%20live")[1:] if r[3] == 1]
    a = [r for r in rows if r[2] == "Ala"][0]; b = [r for r in rows if r[2] == "Bartek"][0]
    check("Liczba: 29,5 przy 29±2 poprawne, punkty %d" % a[6], a[5] == 1 and 0 < a[6] < 1000)
    check("Liczba: 40 niepoprawne", b[5] == 0 and b[6] == 0)
    check("Liczba: TV rysuje rozkład na osi", tv.locator("svg.strip").count() >= 1)
    check("Liczba: telefon pokazuje poprawną wartość", "29 (± 2)" in P["Bartek"].inner_text("body"))
    host("next"); time.sleep(1); host("reveal"); host("next")
    # --- tryb dowolny ---
    setv("Tryb gry", "Dowolny moment"); setv("Tryb dowolny: otwarty od", ""); setv("Tryb dowolny: otwarty do", "")
    f = page(390, 800); f.goto(B + "/v2/"); f.fill("#nm", "Zenek"); f.click("#go")
    f.wait_for_selector("button.tile"); f.click('button.tile[data-i="0"]'); f.click("#ok")
    f.wait_for_selector("text=Dobrze"); check("dowolny: P/F i wyjaśnienie", "Heinego" in f.inner_text("body"))
    f.click("#nx"); f.wait_for_selector("#nv"); f.fill("#nv", "29"); f.click("#nsend"); f.wait_for_selector("text=Dobrze")
    check("dowolny: Liczba dokładnie = 1000 pkt maks.", "29 (± 2)" in f.inner_text("body"))
    # --- panel: nowe pytanie liczbowe ---
    ad = page(1280, 900); ad.goto(B + "/v2/#admin"); ad.fill("#apin", "1213"); ad.click("#ago"); ad.wait_for_selector(".agrid")
    ad.click(".atab[data-t=q]"); ad.click("#qnew"); ad.wait_for_selector(".amodal")
    ad.fill("#e0", "Test rodzajów"); ad.click('input[name=ek][value="Liczba"]', force=True)
    ad.fill("#e1", "Ile to 2+2?"); ad.fill("#e13", "4"); ad.fill("#e14", "0,5"); ad.fill("#e12", "Arytmetyka.")
    ad.screenshot(path="/tmp/kind_editor.png")
    ad.click("#esave"); ad.wait_for_selector("text=Pytanie dodane")
    last = api("/__sheet?n=Pytania")[-1]; head = api("/__sheet?n=Pytania")[0]
    check("panel: zapis pytania liczbowego (L, M, N, O)", last[11] == "Liczba" and last[12] == "Arytmetyka." and last[13] == "4" and last[14] == "0.5" and last[6] == "")
    check("panel: nagłówki kolumn L–O", head[11].startswith("Rodzaj") and head[14].startswith("Tolerancja"))
    ad.select_option("#fcat", "Test rodzajów"); ad.screenshot(path="/tmp/kind_list.png", full_page=True)
    check("brak błędów JavaScript", not errs)
    if errs: print(errs)
    for br in brs: br.close()
bad = [n for n, ok in results if not ok]
print("\n%d/%d testów zaliczonych" % (len(results) - len(bad), len(results)))
sys.exit(1 if bad else 0)
