# Test panelu #admin na mocku (node tests/server.js 8787).
import json, re, sys, time, urllib.request
from playwright.sync_api import sync_playwright

B = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:8787"
def api(p): return json.loads(urllib.request.urlopen(B + p).read())
def host(cmd):
    req = urllib.request.Request(B + "/exec", data=json.dumps({"type": "host", "cmd": cmd, "pin": "1213"}).encode(), method="POST")
    return json.loads(urllib.request.urlopen(req).read())
def setting(label):
    for r in api("/__sheet?n=Sterowanie"):
        if r and r[0] == label: return r[1]
def qrows(): return api("/__sheet?n=Pytania")

results = []
def check(name, cond):
    results.append((name, bool(cond))); print(("OK  " if cond else "BŁĄD") + " " + name)

api("/__reset")
with sync_playwright() as p:
    br = p.chromium.launch(); ctx = br.new_context(viewport={"width": 1280, "height": 900})
    ctx.route(re.compile(r"^https?://(?!localhost|127\.0\.0\.1)"), lambda r: r.abort())
    pg = ctx.new_page(); errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    pg.on("dialog", lambda d: d.accept())
    pg.goto(B + "/v2/#admin")
    pg.fill("#apin", "0000"); pg.click("#ago"); pg.wait_for_selector("text=Nieprawidłowy PIN")
    check("zły PIN odrzucony", True)
    pg.fill("#apin", "1213"); pg.click("#ago"); pg.wait_for_selector(".agrid")
    check("logowanie poprawnym PIN-em", True)

    # ustawienia: tryb gry, liczba pytań, data
    lab = lambda t: pg.locator(".afield", has=pg.locator("label", has_text=t)).locator(".ain")
    lab("Tryb gry").select_option("Dowolny moment")
    lab("Liczba pytań w quizie").fill("3")
    lab("Tryb dowolny: otwarty od").fill("2026-12-01T18:30")
    check("licznik niezapisanych zmian = 3", "Niezapisane zmiany: 3" in pg.inner_text(".asave"))
    pg.click("#ssave"); pg.wait_for_selector("text=Ustawienia zapisane")
    check("zapis do arkusza: Tryb gry", setting("Tryb gry") == "Dowolny moment")
    check("zapis do arkusza: Liczba pytań", str(setting("Liczba pytań w quizie")) == "3")
    check("zapis do arkusza: data", setting("Tryb dowolny: otwarty od") == "2026-12-01 18:30")
    check("pasek „Do gry trafi: 3”", "Do gry trafi: 3" in pg.inner_text(".astat"))
    lab("Tryb gry").select_option("Na żywo"); lab("Liczba pytań w quizie").fill("5"); pg.click("#ssave"); pg.wait_for_selector("text=Ustawienia zapisane")

    # pytania
    pg.click(".atab[data-t=q]"); pg.wait_for_selector("#qlist")
    pg.locator('.qrow[data-r="8"] .sw').click(); pg.wait_for_selector("text=Pytanie włączone")
    check("włączenie pytania (wiersz 8 = TAK)", qrows()[7][7] == "TAK")
    # edycja
    pg.click('.qrow[data-r="3"] .qbody'); pg.wait_for_selector(".amodal")
    pg.fill("#e1", "Zmienione pytanie?"); pg.check('input[name=ec][value="D"]'); pg.click("#esave"); pg.wait_for_selector("text=Pytanie zapisane")
    r3 = qrows()[2]
    check("edycja treści i poprawnej odpowiedzi", r3[1] == "Zmienione pytanie?" and r3[6] == "D")
    # walidacja: brak odpowiedzi
    pg.click("#qnew"); pg.wait_for_selector(".amodal")
    pg.fill("#e0", "Nowa kategoria"); pg.fill("#e1", "Nowe pytanie?"); pg.fill("#e2", "1"); pg.fill("#e3", "2"); pg.fill("#e4", "3")
    pg.click("#esave"); pg.wait_for_selector("text=Uzupełnij odpowiedź D")
    check("walidacja brakującej odpowiedzi", True)
    pg.fill("#e5", "4"); pg.check('input[name=ec][value="B"]'); pg.click("#esave"); pg.wait_for_selector("text=Pytanie dodane")
    last = qrows()[-1]
    check("nowe pytanie na końcu", last[1] == "Nowe pytanie?" and last[6] == "B" and last[0] == "Nowa kategoria")
    check("nowa kategoria dopisana do Kategorii", ["Nowa kategoria"] in api("/__sheet?n=Kategorie"))
    # wstawienie pod wierszem 2
    pg.click('.qrow[data-r="2"] [data-a=ins]'); pg.wait_for_selector(".amodal")
    for i, v in enumerate(["W", "Wstawione?", "a", "b", "c", "d"]): pg.fill("#e%d" % i, v)
    pg.click("#esave"); pg.wait_for_selector("text=Pytanie dodane")
    check("wstawienie pod wierszem 2", qrows()[2][1] == "Wstawione?")
    # przesuwanie w obrębie filtra kategorii
    pg.select_option("#fcat", "Wiedza żeglarska")
    before = [r[1] for r in qrows()[1:5]]
    pg.click('.qrow[data-r="4"] [data-a=up]'); pg.wait_for_selector("text=Kolejność zmieniona")
    after = [r[1] for r in qrows()[1:5]]
    check("przesunięcie w górę z pominięciem innej kategorii (" + " | ".join(after) + ")", after[0] == before[2] and after[1] == before[0])
    # w trakcie gry usuwanie zablokowane
    host("start"); pg.click("#arf"); pg.wait_for_selector("text=Trwa gra na żywo")
    n0 = len(qrows())
    pg.click('.qrow[data-r="5"] [data-a=del]'); pg.wait_for_selector(".toast.bad")
    check("usuwanie zablokowane w trakcie gry", len(qrows()) == n0)
    host("reset"); pg.click("#arf"); time.sleep(0.5)
    pg.click('.qrow[data-r="5"] [data-a=del]'); pg.wait_for_selector("text=Pytanie usunięte")
    check("usuwanie po zakończeniu gry", len(qrows()) == n0 - 1)
    # zmiana PIN
    pg.click(".atab[data-t=set]"); lab("PIN prowadzącego").fill("4321"); pg.click("#ssave"); pg.wait_for_selector("text=Ustawienia zapisane")
    check("zmiana PIN-u zapisana, sesja trwa", str(setting("PIN prowadzącego")) == "4321" and pg.locator(".agrid").count() == 1)
    check("brak błędów JavaScript", not errs)
    if errs: print(errs)
    br.close()

bad = [n for n, ok in results if not ok]
print("\n%d/%d testów zaliczonych" % (len(results) - len(bad), len(results)))
sys.exit(1 if bad else 0)
