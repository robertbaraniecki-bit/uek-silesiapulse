# AXIS_BRIEFING — SilesiaPulse · Kody referencyjne
**Agent:** Nexus · **Aktualizacja:** 25.05.2026 · **Fazy 1–4 zamknięte**

> Plik uzupełniający do Notion. Zawiera wyłącznie kody krytyczne.
> Stan projektu → fetch Notion `355f37b1-ed15-810e-ad76-d59f4f27fca5`

---

## ŚRODOWISKO

```
Excel:     C:\Users\Oloros\OneDrive\Pulpit\Projekt UEK Czerwiec 2026\PROJEKT\UEK_Baraniecki_2026.xlsm
Python:    C:\Projekty\uek-silesiapulse\python\silesiapulse_ai.py
Repo:      C:\Projekty\uek-silesiapulse\  (GitHub: robertbaraniecki-bit/uek-silesiapulse)
Model AI:  claude-sonnet-4-6  ← NIE claude-sonnet-4-20250514 (deprecated 15.06.2026)
DAX sep:   średnik ;  (nie przecinek!) · bez VAR/RETURN · cudzysłowy ręcznie
```

---

## POWER QUERY — WZORCE M

```m
// 1. Separator dziesiętny — ZAWSZE przed TransformColumnTypes
= Table.ReplaceValue(PreviousStep, ".", ",", Replacer.ReplaceText, {"Zamkniecie","USD","EUR"})

// 2. Scalanie tabel — NestedJoin (nie Table.Join — unika duplikatów kluczy)
= Table.NestedJoin(tabela_lewa, "Data", tabela_prawa, "Data", "Tmp", JoinKind.Left)
= Table.ExpandTableColumn(PreviousStep, "Tmp", {"VIX","GPR","GPRC_POL"}, {"VIX","GPR","GPRC_POL"})

// 3. FillDown — GPR miesięczny → dni robocze (LOCF)
= Table.FillDown(PreviousStep, {"GPR","GPRC_POL","VIX"})

// 4. Tabela kalendarza (Power Query M — nie DAX, stara wersja PP)
let
    StartDate = #date(2020,1,1), EndDate = #date(2026,12,31),
    DateList  = List.Dates(StartDate, Duration.Days(EndDate-StartDate)+1, #duration(1,0,0,0)),
    ToTable   = Table.FromList(DateList, Splitter.SplitByNothing()),
    Rename    = Table.RenameColumns(ToTable, {{"Column1","Data"}}),
    TypeDate  = Table.TransformColumnTypes(Rename, {{"Data", type date}})
in  TypeDate
```

---

## DAX — 5 MIAR (separatory: średniki!)

```dax
-- Relacja: qry_Kalendarz[Data] 1→* qry_MASTER[Data]

Avg VIX 30d :=
AVERAGE(qry_MASTER[VIX])

GPR Prev Month :=
CALCULATE(AVERAGE(qry_MASTER[GPR]); PREVIOUSMONTH(qry_Kalendarz[Data]))

GPR Change % :=
DIVIDE(AVERAGE(qry_MASTER[GPR]) - [GPR Prev Month]; [GPR Prev Month]; 0)

USD Trend :=
DIVIDE(
    AVERAGE(qry_MASTER[USD_PLN])
        - CALCULATE(AVERAGE(qry_MASTER[USD_PLN]); PREVIOUSMONTH(qry_Kalendarz[Data]));
    CALCULATE(AVERAGE(qry_MASTER[USD_PLN]); PREVIOUSMONTH(qry_Kalendarz[Data]));
    0)

Risk Label :=
IF([GPR Change %] > 0.05; 1; IF([GPR Change %] < -0.05; -1; 0))
-- 1=ryzyko rośnie · -1=stabilizacja · 0=neutralnie
```

---

## PYTHON — silesiapulse_ai.py

```python
import xlwings as xw
import anthropic
import pandas as pd
from datetime import datetime

def polacz_z_excelem():
    return xw.apps.active.books.active

def czytaj_config(wb):
    cfg = wb.sheets["Config"]
    return {
        "api_key":    cfg["B1"].value,
        "model":      cfg["B2"].value,       # claude-sonnet-4-6
        "max_tokens": int(cfg["B3"].value),  # 600
        "days_back":  int(cfg["B4"].value),  # 30
    }

def pobierz_dane_master(wb, days_back):
    ws  = wb.sheets["qry_MASTER"]
    df  = ws.range("A1").expand("table").options(pd.DataFrame, header=True).value
    df.index = pd.to_datetime(df.index)
    return df, df.tail(days_back)

def agreguj_dane(df_okno):
    return {
        "vix_avg":  round(df_okno["VIX"].mean(), 2),
        "vix_max":  round(df_okno["VIX"].max(), 2),
        "gpr_avg":  round(df_okno["GPR"].mean(), 2),
        "usd_avg":  round(df_okno["USD_PLN"].mean(), 4),
        "usd_last": round(df_okno["USD_PLN"].iloc[-1], 4),
        "wig_last": round(df_okno["WIG20"].iloc[-1], 0),
        "data_od":  df_okno.index[0].strftime("%Y-%m-%d"),
        "data_do":  df_okno.index[-1].strftime("%Y-%m-%d"),
        "n_days":   len(df_okno),
    }

def buduj_prompt_zbiorczy(a):
    return (
        f"Jesteś analitykiem ryzyka geopolitycznego. Przeanalizuj dane z okresu "
        f"{a['data_od']} do {a['data_do']} ({a['n_days']} dni). "
        f"GPR avg {a['gpr_avg']}, VIX avg {a['vix_avg']} max {a['vix_max']}, "
        f"USD/PLN avg {a['usd_avg']} ostatni {a['usd_last']}, WIG20 {a['wig_last']}. "
        "Zwięzłe podsumowanie sytuacji geopolityczno-rynkowej dla inwestora z ekspozycją "
        "na rynek śląski. Brak Markdown, czysty tekst, max 5 zdań."
    )

def buduj_prompt_dzienny(wiersz):
    return (
        f"Dane z dnia {wiersz.name.strftime('%Y-%m-%d')}: "
        f"GPR={wiersz.get('GPR','nd')}, VIX={wiersz.get('VIX','nd')}, "
        f"USD/PLN={wiersz.get('USD_PLN','nd')}, WIG20={wiersz.get('WIG20','nd')}, "
        f"JSW={wiersz.get('JSW','nd')}, KGHM={wiersz.get('KGHM','nd')}, PKN={wiersz.get('PKN','nd')}. "
        "2-zdaniowy komentarz analityczny dla regionalnego inwestora śląskiego. Brak Markdown."
    )

def wywolaj_claude(prompt, config):
    client = anthropic.Anthropic(api_key=config["api_key"])
    msg = client.messages.create(
        model=config["model"],
        max_tokens=config["max_tokens"],
        messages=[{"role": "user", "content": prompt}]
    )
    return msg.content[0].text

def zapisz_komentarz_zbiorczy(wb, komentarz, a):
    ai = wb.sheets["AI_Summary"]
    ai["B2"].value = datetime.now().strftime("%Y-%m-%d %H:%M")
    ai["B3"].value = f"{a['data_od']} — {a['data_do']}"
    ai["B4"].value = a["n_days"]
    ai["B5"].value = komentarz

def zapisz_komentarze_dzienne(wb, df, df_okno, config):
    ws = wb.sheets["qry_MASTER"]
    for i in df_okno.index:
        wiersz    = df.loc[i]
        komentarz = wywolaj_claude(buduj_prompt_dzienny(wiersz), config)
        excel_row = df.index.get_loc(i) + 2   # KRYTYCZNE: get_loc, nie i+2
        ws.range(f"K{excel_row}").value = komentarz
        print(f"  [{i.strftime('%Y-%m-%d')}] OK")

def generuj_komentarz_ai():
    wb     = polacz_z_excelem()
    config = czytaj_config(wb)
    df, df_okno = pobierz_dane_master(wb, config["days_back"])

    agregat = agreguj_dane(df_okno)
    zapisz_komentarz_zbiorczy(wb, wywolaj_claude(buduj_prompt_zbiorczy(agregat), config), agregat)

    zapisz_komentarze_dzienne(wb, df, df_okno, config)

    panel = wb.sheets["Panel"]
    panel["C15"].value = datetime.now().strftime("%Y-%m-%d %H:%M")
    panel["C16"].value = len(df_okno)
    panel["C17"].value = "OK"
    print(f"Gotowe. {len(df_okno)} komentarzy + AI_Summary.")

if __name__ == "__main__":
    generuj_komentarz_ai()
```

---

## VBA — moduł SilesiaPulse_Makra

```vba
Sub GenerujKomentarzAI()
    Dim cmd As String
    cmd = "python " & Chr(34) & "C:\Projekty\uek-silesiapulse\python\silesiapulse_ai.py" & Chr(34)
    Shell "cmd.exe /c " & cmd, vbNormalFocus
    MsgBox "Generowanie komentarzy AI uruchomione.", vbInformation, "SilesiaPulse"
End Sub

Sub OdswierzDane()
    Dim t As Double : t = Timer
    ThisWorkbook.RefreshAll
    Dim panel As Worksheet : Set panel = ThisWorkbook.Sheets("Panel")
    panel.Range("C14").Value = Now()
    panel.Range("C17").Value = "OK"
    ' UWAGA: Interior.Color = xlNone — kolor obsługuje formatowanie warunkowe, nie VBA
    MsgBox "Dane odświeżone (" & Round(Timer - t) & "s).", vbInformation, "SilesiaPulse"
End Sub
```

---

## ARKUSZ CONFIG (ukryty)

```
A1: ANTHROPIC_API_KEY  │ B1: sk-ant-... (aktywny)
A2: MODEL              │ B2: claude-sonnet-4-6
A3: MAX_TOKENS         │ B3: 600
A4: DAYS_BACK          │ B4: 30
A5: STOOQ_API_KEY      │ B5: [klucz Roberta] ← DO DODANIA
```

---

## STOOQ API — wzorzec Python

```python
def url(ticker: str, interval: str, api_key: str) -> str:
    return f"https://stooq.pl/q/d/l/?s={ticker}&i={interval}&apikey={api_key}"

def get_data(ticker: str, interval: str, api_key: str) -> pd.DataFrame:
    df = pd.read_csv(
        filepath_or_buffer=url(ticker, interval, api_key),
        index_col=0, usecols=[0, 4], parse_dates=True
        # usecols=[0,4]: Data + Zamkniecie (nie Close — Stooq PL!)
    )
    df.columns = [ticker]
    return df

# Tickery SilesiaPulse (małe litery): "jsw", "kgh", "pkn", "wig"
# Klucz API: stooq.pl/q/d/?s=wig&get_apikey (jednorazowa captcha)
# Klucz → Config!B5
```

---

## PANEL STEROWANIA — mapa komórek

```
Arkusz: Panel (pierwszy od lewej)
C14 → data ostatniego odświeżenia danych    (ustawia: OdswierzDane VBA)
C15 → data ostatniego komentarza AI         (ustawia: silesiapulse_ai.py)
C16 → wiersze przetworzone                  (ustawia: silesiapulse_ai.py)
C17 → OK / BŁĄD  ← formatowanie warunkowe: OK=zielony, BŁĄD=czerwony
```

---

## PUŁAPKI I ZASADY

```
PYTHON:
  ✅ excel_row = df.index.get_loc(i) + 2    ← prawidłowe
  ❌ excel_row = i + 2                       ← błąd (i to data, nie liczba)

DAX:
  ✅ separatory: średniki ;
  ❌ przecinki , (błąd składni w tej wersji Power Pivot)
  ❌ VAR/RETURN  (nieobsługiwane)
  ❌ Nowa tabela obliczeniowa (brak przycisku — kalendarz w Power Query M)

GIT:
  Konwencja: python: / excel: / vba: / data: / docs:
  Zasada: commit po każdej istotnej zmianie, nie tylko po sesjach

SEPARACJA:
  SilesiaPulse = Excel + Power BI (UEK academic POC)
  Geo Pulse    = Python → JSON → HTML (produkcja)
  Nazewnictwo NIE transferuje się między projektami nigdy.
```

---

*Nexus · SilesiaPulse · UEK 2026 · 25.05.2026*
