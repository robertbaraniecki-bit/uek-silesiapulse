# =============================================================================
# silesiapulse_ai.py — SilesiaPulse / Projekt UEK 2026
# Agent: Nexus | Robert Baraniecki | dr hab. K. Kania
# =============================================================================
# ARCHITEKTURA:
#   Excel + Python = warstwa wykonawcza (silnik, logika, AI)
#   Power BI       = warstwa wizualizacyjna (odczyt gotowych danych)
#
# WYWOŁANIE:
#   Z VBA przez: Shell "python silesiapulse_ai.py"
#   lub: xlwings RunPython (zależnie od konfiguracji po zajęciach dr Kaczmarzyka)
#
# LOGIKA:
#   1. Czyta DAYS_BACK oraz TYP_INWESTORA z arkusza Config
#   2. Pobiera ostatnie N wierszy z qry_MASTER
#   3. Agreguje dane (średnie, trendy, min/max)
#   4. Wysyła zagregowane dane do Claude API (prompt dostosowany do typu inwestora)
#   5. Zapisuje komentarz dzienny do kolumny Komentarz_AI (przyrostowo)
#   6. Zapisuje komentarz zbiorczy do arkusza AI_Summary
# =============================================================================

from dotenv import load_dotenv
import os
import xlwings as xw
import anthropic
import pandas as pd
from datetime import datetime, timedelta

# =============================================================================
# SŁOWNIKI PROMPTÓW — wg typu inwestora
# =============================================================================

PROMPT_ZBIORCZY = {
    'ostrożny': """Jesteś analitykiem ryzyka geopolitycznego specjalizującym się w gospodarce regionu Śląska (Polska).
Piszesz dla inwestora ostrożnego — priorytetem jest ochrona kapitału, unikanie strat i sygnały ostrzegawcze.

Przeanalizuj poniższe zagregowane dane finansowe i geopolityczne dla okresu {okres}:

RYZYKO GEOPOLITYCZNE:
- GPR Global (indeks Caldara-Iacoviello): średnia {gpr_srednia}, min {gpr_min}, max {gpr_max}, trend {gpr_trend:+.2f} pkt
- GPRC_POL (udział medialny Polski w globalnym GPR, skala 0–1): średnia {gprc_pol_srednia:.4f}
- Norma historyczna GPR: 100–150 pkt. Wartości powyżej 200 oznaczają podwyższone ryzyko.

RYNKI FINANSOWE:
- VIX (indeks strachu): średnia {vix_srednia}, max {vix_max}
- USD/PLN: średnia {usd_pln_srednia}, zmiana {usd_pln_trend:+.4f}
- EUR/PLN: średnia {eur_pln_srednia}

SPÓŁKI ŚLĄSKIE I WIG20:
- WIG20: średnia {wig20_srednia}, trend {wig20_trend:+.2f} pkt
- JSW: średnia {jsw_srednia}, trend {jsw_trend:+.2f} PLN
- KGHM: średnia {kghm_srednia}, trend {kghm_trend:+.2f} PLN
- PKN Orlen: średnia {pkn_srednia}, trend {pkn_trend:+.2f} PLN

Napisz zwięzły komentarz analityczny (max 150 słów) który:
1. Ocenia poziom ryzyka geopolitycznego — ze szczególnym uwzględnieniem sygnałów ostrzegawczych
2. Wskazuje zagrożenia dla kapitału w tym okresie
3. Wymienia spółki śląskie które zachowywały się najgorzej lub najniestabilniej
4. Formułuje rekomendację ochronną dla inwestora konserwatywnego (np. redukcja ekspozycji, czekanie)

Pisz po polsku, konkretnie, bez zbędnych wstępów. Bez formatowania markdown, bez gwiazdek, bez nagłówków — tylko czysty tekst z podziałem na akapity.""",

    'aktywny': """Jesteś analitykiem ryzyka geopolitycznego specjalizującym się w gospodarce regionu Śląska (Polska).
Piszesz dla inwestora aktywnego — szuka okazji rynkowych, akceptuje wyższe ryzyko, interesuje go momentum i potencjał wzrostu.

Przeanalizuj poniższe zagregowane dane finansowe i geopolityczne dla okresu {okres}:

RYZYKO GEOPOLITYCZNE:
- GPR Global (indeks Caldara-Iacoviello): średnia {gpr_srednia}, min {gpr_min}, max {gpr_max}, trend {gpr_trend:+.2f} pkt
- GPRC_POL (udział medialny Polski w globalnym GPR, skala 0–1): średnia {gprc_pol_srednia:.4f}
- Norma historyczna GPR: 100–150 pkt. Wartości powyżej 200 oznaczają podwyższone ryzyko.

RYNKI FINANSOWE:
- VIX (indeks strachu): średnia {vix_srednia}, max {vix_max}
- USD/PLN: średnia {usd_pln_srednia}, zmiana {usd_pln_trend:+.4f}
- EUR/PLN: średnia {eur_pln_srednia}

SPÓŁKI ŚLĄSKIE I WIG20:
- WIG20: średnia {wig20_srednia}, trend {wig20_trend:+.2f} pkt
- JSW: średnia {jsw_srednia}, trend {jsw_trend:+.2f} PLN
- KGHM: średnia {kghm_srednia}, trend {kghm_trend:+.2f} PLN
- PKN Orlen: średnia {pkn_srednia}, trend {pkn_trend:+.2f} PLN

Napisz zwięzły komentarz analityczny (max 150 słów) który:
1. Ocenia poziom ryzyka geopolitycznego — ze szczególnym uwzględnieniem okazji rynkowych jakie generuje
2. Wskazuje które spółki śląskie mogą skorzystać na bieżącej sytuacji
3. Opisuje momentum — które walory mają najsilniejszy pozytywny trend
4. Formułuje rekomendację aktywną dla inwestora nastawionego na zysk (np. zwiększenie pozycji, obserwacja konkretnych spółek)

Pisz po polsku, konkretnie, bez zbędnych wstępów. Bez formatowania markdown, bez gwiazdek, bez nagłówków — tylko czysty tekst z podziałem na akapity."""
}

PROMPT_DZIENNY = {
    'ostrożny': """Jesteś analitykiem ryzyka geopolitycznego. Data analizy: {data}.
Piszesz dla inwestora ostrożnego — priorytet: ochrona kapitału, sygnały ostrzegawcze.

Dane dzienne:
GPR Global: {gpr:.1f} | GPRC_POL: {gprc_pol:.4f}
VIX: {vix:.2f} | USD/PLN: {usd_pln:.4f} | EUR/PLN: {eur_pln:.4f}
WIG20: {wig20:.1f} | JSW: {jsw:.2f} | KGHM: {kghm:.2f} | PKN: {pkn:.2f}

Napisz dokładnie 2 zdania komentarza — skup się na zagrożeniach i sygnałach ostrzegawczych dla regionalnego inwestora śląskiego.
Zasady: bez tytułów, bez nagłówków, bez formatowania markdown, bez gwiazdek. Tylko czysty tekst. Maksymalnie 50 słów łącznie. Po polsku.""",

    'aktywny': """Jesteś analitykiem ryzyka geopolitycznego. Data analizy: {data}.
Piszesz dla inwestora aktywnego — priorytet: okazje rynkowe, momentum, potencjał wzrostu.

Dane dzienne:
GPR Global: {gpr:.1f} | GPRC_POL: {gprc_pol:.4f}
VIX: {vix:.2f} | USD/PLN: {usd_pln:.4f} | EUR/PLN: {eur_pln:.4f}
WIG20: {wig20:.1f} | JSW: {jsw:.2f} | KGHM: {kghm:.2f} | PKN: {pkn:.2f}

Napisz dokładnie 2 zdania komentarza — skup się na okazjach i potencjale wzrostu dla regionalnego inwestora śląskiego.
Zasady: bez tytułów, bez nagłówków, bez formatowania markdown, bez gwiazdek. Tylko czysty tekst. Maksymalnie 50 słów łącznie. Po polsku."""
}

# =============================================================================
# KROK 1 — POŁĄCZENIE Z EXCELEM I ODCZYT KONFIGURACJI
# =============================================================================

def polacz_z_excelem():
    """Łączy się z już otwartym plikiem UEK_Baraniecki_2026.xlsm"""
    app = xw.apps.active
    wb = app.books.active
    return wb

def czytaj_config(wb):
    load_dotenv()
    ws_config = wb.sheets['Config']
    typ = ws_config.range('B6').value
    if typ not in ('ostrożny', 'aktywny'):
        typ = 'ostrożny'
    config = {
        'api_key':        os.getenv("ANTHROPIC_API_KEY") or ws_config.range('B1').value,
        'model':          ws_config.range('B2').value,
        'max_tokens':     int(ws_config.range('B3').value),
        'days_back':      int(ws_config.range('B4').value),
        'typ_inwestora':  typ
    }
    return config

# =============================================================================
# KROK 2 — POBRANIE DANYCH Z qry_MASTER
# =============================================================================

def pobierz_dane_master(wb, days_back):
    ws_master = wb.sheets['qry_MASTER']
    dane = ws_master.range('A1').expand().value
    naglowki = dane[0]
    wiersze = dane[1:]
    df = pd.DataFrame(wiersze, columns=naglowki)
    df['Data'] = pd.to_datetime(df['Data'])
    df = df.sort_values('Data').reset_index(drop=True)
    data_graniczna = datetime.now() - timedelta(days=days_back)
    df_okno = df[df['Data'] >= data_graniczna].copy()
    return df, df_okno

# =============================================================================
# KROK 3 — AGREGACJA DANYCH
# =============================================================================

def agreguj_dane(df_okno):
    if df_okno.empty:
        return None
    data_od = df_okno['Data'].min().strftime('%Y-%m-%d')
    data_do = df_okno['Data'].max().strftime('%Y-%m-%d')
    liczba_dni = len(df_okno)
    agregat = {
        'okres':            f"{data_od} do {data_do} ({liczba_dni} dni roboczych)",
        'gpr_srednia':      round(df_okno['GPR'].mean(), 2),
        'gpr_min':          round(df_okno['GPR'].min(), 2),
        'gpr_max':          round(df_okno['GPR'].max(), 2),
        'gpr_trend':        round(df_okno['GPR'].iloc[-1] - df_okno['GPR'].iloc[0], 2),
        'gprc_pol_srednia': round(df_okno['GPRC_POL'].mean(), 4),
        'vix_srednia':      round(df_okno['VIX'].mean(), 2),
        'vix_max':          round(df_okno['VIX'].max(), 2),
        'usd_pln_srednia':  round(df_okno['USD_PLN'].mean(), 4),
        'usd_pln_trend':    round(df_okno['USD_PLN'].iloc[-1] - df_okno['USD_PLN'].iloc[0], 4),
        'eur_pln_srednia':  round(df_okno['EUR_PLN'].mean(), 4),
        'wig20_srednia':    round(df_okno['WIG20'].mean(), 2),
        'wig20_trend':      round(df_okno['WIG20'].iloc[-1] - df_okno['WIG20'].iloc[0], 2),
        'jsw_srednia':      round(df_okno['JSW'].mean(), 2),
        'jsw_trend':        round(df_okno['JSW'].iloc[-1] - df_okno['JSW'].iloc[0], 2),
        'kghm_srednia':     round(df_okno['KGHM'].mean(), 2),
        'kghm_trend':       round(df_okno['KGHM'].iloc[-1] - df_okno['KGHM'].iloc[0], 2),
        'pkn_srednia':      round(df_okno['PKN'].mean(), 2),
        'pkn_trend':        round(df_okno['PKN'].iloc[-1] - df_okno['PKN'].iloc[0], 2),
    }
    return agregat

# =============================================================================
# KROK 4 — BUDOWA PROMPTÓW
# =============================================================================

def buduj_prompt_zbiorczy(agregat, typ_inwestora):
    return PROMPT_ZBIORCZY[typ_inwestora].format(**agregat)

def buduj_prompt_dzienny(wiersz, typ_inwestora):
    return PROMPT_DZIENNY[typ_inwestora].format(
        data=wiersz['Data'].strftime('%Y-%m-%d'),
        gpr=wiersz['GPR'],
        gprc_pol=wiersz['GPRC_POL'],
        vix=wiersz['VIX'],
        usd_pln=wiersz['USD_PLN'],
        eur_pln=wiersz['EUR_PLN'],
        wig20=wiersz['WIG20'],
        jsw=wiersz['JSW'],
        kghm=wiersz['KGHM'],
        pkn=wiersz['PKN']
    )

# =============================================================================
# KROK 5 — WYWOŁANIE CLAUDE API
# =============================================================================

def wywolaj_claude(prompt, config):
    klient = anthropic.Anthropic(api_key=config['api_key'])
    odpowiedz = klient.messages.create(
        model=config['model'],
        max_tokens=config['max_tokens'],
        messages=[{"role": "user", "content": prompt}]
    )
    return odpowiedz.content[0].text.strip()

# =============================================================================
# KROK 6 — ZAPIS WYNIKÓW DO EXCELA
# =============================================================================

def zapisz_komentarz_zbiorczy(wb, komentarz, agregat):
    if 'AI_Summary' not in [s.name for s in wb.sheets]:
        wb.sheets.add('AI_Summary')
    ws_summary = wb.sheets['AI_Summary']
    ws_summary.range('A1').value = f"Wygenerowano: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws_summary.range('A2').value = f"Okres: {agregat['okres']}"
    ws_summary.range('A3').value = komentarz
    print(f"✅ Komentarz zbiorczy zapisany do AI_Summary")

def zapisz_komentarze_dzienne(wb, df, df_okno, config):
    ws_master = wb.sheets['qry_MASTER']
    naglowki = ws_master.range('A1').expand('right').value
    if 'Komentarz_AI' not in naglowki:
        print("⚠️ Brak kolumny Komentarz_AI w qry_MASTER")
        return
    idx_komentarz = naglowki.index('Komentarz_AI')
    col_letter_offset = idx_komentarz + 1
    wiersze_do_przetworzenia = 0
    wiersze_przetworzone = 0
    for i, row in df_okno.iterrows():
        excel_row = df.index.get_loc(i) + 2
        istniejacy = ws_master.range((excel_row, col_letter_offset)).value
        if istniejacy:
            continue
        wiersze_do_przetworzenia += 1
        try:
            prompt = buduj_prompt_dzienny(row, config['typ_inwestora'])
            komentarz = wywolaj_claude(prompt, config)
            ws_master.range((excel_row, col_letter_offset)).value = komentarz
            wiersze_przetworzone += 1
            print(f"✅ {row['Data'].strftime('%Y-%m-%d')} — komentarz zapisany")
        except Exception as e:
            print(f"⚠️ {row['Data'].strftime('%Y-%m-%d')} — błąd API: {e}")
            continue
    print(f"\n✅ Komentarze dzienne: {wiersze_przetworzone}/{wiersze_do_przetworzenia} wierszy")
    return wiersze_przetworzone

# =============================================================================
# GŁÓWNA FUNKCJA — WYWOŁYWANA PRZEZ VBA
# =============================================================================

def generuj_komentarz_ai():
    print("=" * 50)
    print("SilesiaPulse — Generowanie komentarzy AI")
    print(f"Start: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    try:
        # Krok 1 — połączenie
        wb = polacz_z_excelem()
        print(f"✅ Połączono z: {wb.name}")

        # Krok 2 — konfiguracja
        config = czytaj_config(wb)
        print(f"✅ Konfiguracja: model={config['model']}, days_back={config['days_back']}, typ_inwestora={config['typ_inwestora']}")

        # Krok 3 — dane
        df_pelny, df_okno = pobierz_dane_master(wb, config['days_back'])
        print(f"✅ Pobrano {len(df_okno)} wierszy za ostatnie {config['days_back']} dni")

        if df_okno.empty:
            print("⚠️ Brak danych dla wybranego okresu")
            return

        # Krok 4 — komentarz zbiorczy
        print("\n→ Generowanie komentarza zbiorczego...")
        agregat = agreguj_dane(df_okno)
        prompt_zbiorczy = buduj_prompt_zbiorczy(agregat, config['typ_inwestora'])
        komentarz_zbiorczy = wywolaj_claude(prompt_zbiorczy, config)
        zapisz_komentarz_zbiorczy(wb, komentarz_zbiorczy, agregat)

        # Krok 5 — komentarze dzienne (przyrostowo)
        print("\n→ Generowanie komentarzy dziennych (przyrostowo)...")
        wiersze_przetworzone = zapisz_komentarze_dzienne(wb, df_pelny, df_okno, config)

        # Krok 6 — aktualizacja Panelu
        panel = wb.sheets["Panel"]
        panel["C15"].value = datetime.now().strftime("%Y-%m-%d %H:%M")
        panel["C16"].value = f"{len(df_okno)} (ostatnie {config['days_back']} dni)"

        print("\n" + "=" * 50)
        print("✅ ZAKOŃCZONO POMYŚLNIE")
        print(f"Koniec: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 50)
        print("\n→ Odśwież Power BI aby zobaczyć zaktualizowane komentarze.")

    except Exception as e:
        print(f"\n❌ BŁĄD KRYTYCZNY: {e}")
        raise

# =============================================================================
# URUCHOMIENIE
# =============================================================================

if __name__ == "__main__":
    generuj_komentarz_ai()