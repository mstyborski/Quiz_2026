# Testy lokalne

Mock uruchamia prawdziwy `v2/apps_script.gs` w Node na arkuszu w pamięci (`fixture.json`)
i serwuje `v2/index.html` z adresem serwera podmienionym na lokalny.

```
node tests/server.js 8787          # http://localhost:8787/v2/  (#tv, #gp, #admin; PIN 1213)
python3 tests/e2e_live.py          # gra na żywo: podwójne kliknięcie, utracone żądanie, potwierdzenia
python3 tests/e2e_admin.py         # panel #admin: ustawienia, pytania, blokady w trakcie gry
```

Wymagania: Node 18+, Python 3 z pakietem `playwright` i przeglądarką Chromium.
Pomocnicze adresy mocka: `/__reset`, `/__sheet?n=<karta>`, `/__set?label=..&value=..`, `/__fail?n=1`
(kolejne żądanie z odpowiedzią kończy się błędem 500).
