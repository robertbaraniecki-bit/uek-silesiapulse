import urllib.request
import json
import csv
from datetime import date

ZAKRESY = [
    ("2020-01-01", "2020-12-31"),
    ("2021-01-01", "2021-12-31"),
    ("2022-01-01", "2022-12-31"),
    ("2023-01-01", "2023-12-31"),
    ("2024-01-01", "2024-12-31"),
    ("2025-01-01", "2025-12-31"),
    ("2026-01-01", "2026-05-07")
]

def pobierz(waluta, zakresy):
    wiersze = []
    for od, do in zakresy:
        url = f"https://api.nbp.pl/api/exchangerates/rates/a/{waluta}/{od}/{do}/?format=json"
        print(f"  Pobieram {waluta.upper()} {od} → {do}...")
        with urllib.request.urlopen(url) as r:
            dane = json.loads(r.read())
            for rec in dane["rates"]:
                wiersze.append((rec["effectiveDate"], rec["mid"]))
    return wiersze

for waluta, nazwa_col in [("usd", "USD_PLN"), ("eur", "EUR_PLN")]:
    print(f"\nPobieram {waluta.upper()}...")
    wiersze = pobierz(waluta, ZAKRESY)
    plik = f"nbp_{waluta}_2020_2026.csv"
    with open(plik, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Data", nazwa_col])
        w.writerows(wiersze)
    print(f"  Zapisano {len(wiersze)} wierszy → {plik}")

print("\nGotowe!")