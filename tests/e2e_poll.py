# Test trybu „Głosowanie i opinie”: tematy z panelu, anonimowe głosy i opinie, wykres, chmura, moderacja.
import json, re, sys, time, urllib.request, urllib.parse
from playwright.sync_api import sync_playwright

B = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:8787"
def api(p): return json.loads(urllib.request.urlopen(B + p).read())
def pst(): return api("/exec?action=pstate&full=1")
results = []
def check(name, cond):
    results.append((name, bool(cond))); print(("OK  " if cond else "BŁĄD") + " " + name)
EXT = re.compile(r"^https?://(?!localhost|127\.0\.0\.1)")

api("/__reset")
with sync_playwright() as p:
    brs, errs = [], []
    def page(w, h, init=None):
        br = p.chromium.launch(); brs.append(br); ctx = br.new_context(viewport={"width": w, "height": h})
        ctx.route(EXT, lambda r: r.abort())
        if init: ctx.add_init_script(init)
        pg = ctx.new_page(); pg.on("pageerror", lambda e: errs.append(str(e))); pg.on("dialog", lambda d: d.accept()); return pg
    # --- panel: tematy i tryb ---
    ad = page(1280, 900); ad.goto(B + "/v2/#admin"); ad.fill("#apin", "1213"); ad.click("#ago"); ad.wait_for_selector(".agrid")
    ad.click(".atab[data-t=p]"); ad.wait_for_selector("#pnew")
    ad.click("#pnew"); ad.fill("#p0", "Seminarium"); ad.fill("#p2", "Jak oceniasz wystąpienie?")
    for j, o in zip(range(3, 6), ["Bardzo dobrze", "Dobrze", "Słabo"]): ad.fill("#p%d" % j, o)
    ad.click("#psave"); ad.wait_for_selector("text=Temat dodany")
    ad.click("#pnew"); ad.fill("#p0", "Seminarium"); ad.click('input[name=pk][value="Opinia"]', force=True); ad.fill("#p2", "Mocne strony prelegenta")
    ad.click("#psave"); ad.wait_for_selector("text=Temat dodany")
    rows = api("/__sheet?n=Ankiety")
    check("panel: dwa tematy w karcie Ankiety", len(rows) == 3 and rows[1][1] == "Głosowanie" and rows[2][1] == "Opinia" and rows[1][5] == "Słabo")
    ad.screenshot(path="/tmp/poll_admin.png", full_page=True)
    ad.click(".atab[data-t=set]"); ad.locator(".afield", has=ad.locator("label", has_text="Tryb gry")).locator(".ain").select_option("Głosowanie i opinie")
    ad.click("#ssave"); ad.wait_for_selector("text=Ustawienia zapisane")
    # --- ekran i telefony ---
    tv = page(1280, 760, "localStorage.setItem('pin','1213')"); tv.goto(B + "/v2/#tv"); tv.wait_for_selector("text=Dołącz do głosowania")
    check("TV: ekran startowy głosowania", True)
    P = []
    for k in range(3):
        pg = page(390, 800); pg.goto(B + "/v2/"); pg.wait_for_selector("text=Jesteś w grze"); P.append(pg)
    check("telefon: dołączenie bez pseudonimu", True)
    tv.click('[data-c="pstart"]'); P[0].wait_for_selector("button.tile")
    check("telefon: trzy opcje głosowania", P[0].locator("button.tile").count() == 3)
    P[0].click('button.tile[data-i="0"]'); P[1].click('button.tile[data-i="1"]'); P[2].click('button.tile[data-i="0"]')
    time.sleep(1); P[2].click('button.tile[data-i="2"]'); time.sleep(3)
    check("głosy z ostatnią zmianą: [1,1,1]", pst()["agg"]["counts"] == [1, 1, 1])
    check("telefon: potwierdzenie głosu", "Głos zapisany" in P[2].inner_text("body"))
    check("TV: wykres z liczbą głosów", "Głosów: 3" in tv.inner_text("body"))
    tv.screenshot(path="/tmp/poll_tv_vote.png")
    tv.click('[data-c="pnext"]'); P[0].wait_for_selector("#pop")
    for pg, t in zip(P, ["Świetne przykłady", "świetne przykłady!", "Za szybkie tempo"]):
        pg.fill("#pop", t); pg.click("#psend")
    P[0].fill("#pop", "Dobry kontakt z salą"); P[0].click("#psend")
    time.sleep(3.5)
    body = tv.inner_text("body")
    check("TV: chmura opinii z połączeniem duplikatów (×2)", "×2" in body and "Za szybkie tempo" in body and "Dobry kontakt" in body)
    check("telefon: licznik własnych opinii", "Wysłane opinie: 2" in P[0].inner_text("body"))
    tv.screenshot(path="/tmp/poll_tv_cloud.png")
    tv.locator(".bub", has_text="Za szybkie tempo").click(); time.sleep(3)
    check("moderacja: ukryta opinia znika z ekranu", "Za szybkie tempo" not in tv.inner_text("body"))
    hid = [r for r in api("/__sheet?n=Odpowiedzi%20ankiet") if len(r) > 7 and r[7] == "TAK"]
    check("moderacja: opinia zostaje w arkuszu z oznaczeniem", len(hid) == 1 and hid[0][5] == "Za szybkie tempo")
    tv.click('[data-c="pnext"]'); P[1].wait_for_selector("text=Dziękujemy za udział")
    check("koniec sesji na telefonie i ekranie", "Dziękujemy" in tv.inner_text("body"))
    check("brak błędów JavaScript", not errs)
    if errs: print(errs)
    for br in brs: br.close()
bad = [n for n, ok in results if not ok]
print("\n%d/%d testów zaliczonych" % (len(results) - len(bad), len(results)))
sys.exit(1 if bad else 0)
