# =============================================================================
# pobierz_dane.py — SilesiaPulse / Projekt UEK 2026
# Agent: Nexus | Robert Baraniecki | dr hab. K. Kania
# Master skrypt pobierający wszystkie dane przed odświeżeniem Power Query
# =============================================================================

import urllib.request
import json
import csv
import os
import sys
import yfinance as yf
import pandas as pd
from datetime import datetime

# =============================================================================
# KONFIGURACJA
# =============================================================================

DATA_OD = "2020-01-01"
DATA_DO = datetime.now().strftime("%Y-%m-%d")
DATA_OD_NBP_ZAKRESY = [
    ("2020-01-01", "2020-12-31"),
    ("2021-01-01", "2021-12-31"),
    ("2022-01-01", "2022-12-31"),
    ("2023-01-01", "2023-12-31"),
    ("2024-01-01", "2024-12-31"),
    ("2025-01-01", "2025-12-31"),
    ("2026-01-01", datetime.now().strftime("%Y-%m-%d")),
]

FOLDER_CSV = r"C:\Users\Oloros\OneDrive\Pulpit\SilesiaPulse_SEMINARIUM_28.05\data\dane_csv"
STOOQ_KLUCZ = "Eg9CrBlcSu2kzNme7A5MsyOq0IEawQhK"

SPOLKI_YFINANCE = {
    "PKN.WA": "pkn_d.csv",
    "KGH.WA": "kghm_d.csv",
    "JSW.WA": "jsw_d.csv",
}

# =============================================================================
# MODUŁ 1 — NBP (USD/PLN, EUR/PLN)
# =============================================================================

def pobierz_nbp():
    print("\n" + "=" * 50)
    print("MODUŁ 1 — NBP (kursy walut)")
    print("=" * 50)
    
    bledy = 0

    def pobierz_walute(waluta, zakresy):
        wiersze = []
        for od, do in zakresy:
            url = f"https://api.nbp.pl/api/exchangerates/rates/a/{waluta}/{od}/{do}/?format=json"
            print(f"  Pobieram {waluta.upper()} {od} → {do}...")
            try:
                with urllib.request.urlopen(url) as r:
                    dane = json.loads(r.read())
                    for rec in dane["rates"]:
                        wiersze.append((rec["effectiveDate"], rec["mid"]))
            except Exception as e:
                print(f"  ⚠️ Błąd: {e}")
                return None
        return wiersze

    for waluta, nazwa_col in [("usd", "USD_PLN"), ("eur", "EUR_PLN")]:
        print(f"\nPobieram {waluta.upper()}...")
        wiersze = pobierz_walute(waluta, DATA_OD_NBP_ZAKRESY)
        if wiersze is None:
            print(f"  ❌ Nie udało się pobrać {waluta.upper()}")
            bledy += 1
            continue
        plik = os.path.join(FOLDER_CSV, f"nbp_{waluta}_2020_2026.csv")
        with open(plik, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Data", nazwa_col])
            w.writerows(wiersze)
        print(f"  ✅ Zapisano {len(wiersze)} wierszy → nbp_{waluta}_2020_2026.csv")
    
    return bledy

# =============================================================================
# MODUŁ 2 — Spółki GPW (yfinance)
# =============================================================================

def pobierz_spolki():
    print("\n" + "=" * 50)
    print("MODUŁ 2 — Spółki GPW (yfinance)")
    print("=" * 50)
    
    bledy = 0

    for ticker, plik in SPOLKI_YFINANCE.items():
        print(f"\nPobieram {ticker}...")
        try:
            df = yf.download(
                ticker,
                start=DATA_OD,
                end=DATA_DO,
                progress=False,
                auto_adjust=True
            )

            if df.empty:
                print(f"  ❌ Brak danych dla {ticker}")
                bledy += 1
                continue

            df = df[["Close"]].copy()
            df.index = pd.to_datetime(df.index)
            df.index = df.index.strftime("%Y-%m-%d")
            df.index.name = "Data"
            nazwa_col = ticker.replace(".WA", "")
            df.columns = [nazwa_col]
            df[nazwa_col] = df[nazwa_col].round(2)

            sciezka = os.path.join(FOLDER_CSV, plik)
            df.to_csv(sciezka, sep=",", decimal=",", encoding="utf-8-sig")

            print(f"  ✅ Zapisano {len(df)} wierszy → {plik}")
            print(f"  📅 Zakres: {df.index[0]} → {df.index[-1]}")

        except Exception as e:
            print(f"  ❌ Błąd: {e}")
            bledy += 1
    
    return bledy

# =============================================================================
# MODUŁ 3 — WIG20 (Stooq API)
# =============================================================================

def pobierz_wig20():
    print("\n" + "=" * 50)
    print("MODUŁ 3 — WIG20 (Stooq API)")
    print("=" * 50)

    data_od = "20200101"
    data_do = datetime.now().strftime("%Y%m%d")
    url = f"https://stooq.com/q/d/l/?s=wig20&d1={data_od}&d2={data_do}&i=d&apikey={STOOQ_KLUCZ}"

    print(f"\nPobieram WIG20...")

    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as r:
            zawartosc = r.read().decode('utf-8')

        linie = zawartosc.strip().split('\n')

        if len(linie) < 2 or 'Date' not in linie[0]:
            print(f"  ❌ Błąd — Stooq nie zwrócił danych. Sprawdź klucz API.")
            return 1

        sciezka = os.path.join(FOLDER_CSV, "wig20_d.csv")
        licznik = 0
        with open(sciezka, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(["Data", "Zamkniecie"])
            for linia in linie[1:]:
                cols = linia.split(',')
                if len(cols) >= 5:
                    writer.writerow([cols[0], cols[4]])
                    licznik += 1

        print(f"  ✅ Zapisano {licznik} wierszy → wig20_d.csv")
        return 0

    except Exception as e:
        print(f"  ❌ Błąd: {e}")
        return 1

# =============================================================================
# GŁÓWNA FUNKCJA
# =============================================================================

def main():
    print("=" * 50)
    print("SilesiaPulse — Pobieranie wszystkich danych")
    print(f"Start: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Zakres: {DATA_OD} → {DATA_DO}")
    print("=" * 50)

    bledy_nbp = pobierz_nbp()
    bledy_spolki = pobierz_spolki()
    bledy_wig20 = pobierz_wig20()

    total_bledy = bledy_nbp + bledy_spolki + bledy_wig20

    print("\n" + "=" * 50)
    if total_bledy == 0:
        print("✅ ZAKOŃCZONO — wszystkie dane pobrane")
    else:
        print(f"❌ ZAKOŃCZONO Z BŁĘDAMI — {total_bledy} źródła nie powiodły się")
    print(f"Koniec: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    if total_bledy > 0:
        sys.exit(1)

if __name__ == "__main__":
    main()