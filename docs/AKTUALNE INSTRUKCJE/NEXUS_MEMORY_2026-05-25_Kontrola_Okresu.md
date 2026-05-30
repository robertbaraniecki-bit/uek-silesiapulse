# NEXUS_MEMORY — Kontrola okresu analizy · SilesiaPulse
**Data sesji:** 25.05.2026  
**Status:** Zaprojektowane, NIE zaimplementowane — do wdrożenia w osobnej sesji  
**Powiązane fazy:** Faza 4 (Python xlwings) — rozszerzenie

---

## Problem który rozwiązujemy

`Komentarz_AI` to statyczna kolumna tekstowa — Python zapisuje tekst do konkretnych komórek.  
Power BI slicer daty **ukrywa wiersze** spoza okresu, ale **nie regeneruje** komentarzy dynamicznie.  
Brakuje jednego zbiorczego komentarza AI dla **dowolnie wybranego okresu** (dzień / tydzień / miesiąc / rok / kryzys).

---

## Rozwiązanie: dwuścieżkowy system kontroli okresu

### Nowe komórki w arkuszu Config (ukryty)

```
A6: TRYB     │ B6: LAST_N        ← domyślny tryb bieżący
A7: DATA_OD  │ B7: (puste)       ← aktywne gdy TRYB = DATE_RANGE
A8: DATA_DO  │ B8: (puste)       ← aktywne gdy TRYB = DATE_RANGE
```

Dwa tryby:
- **`LAST_N`** — ostatnie N dni (z Config B4, domyślnie 30). Tryb codzienny.
- **`DATE_RANGE`** — konkretny zakres dat. Tryb analityczny / demo.

---

## Zmiany w Python: silesiapulse_ai.py

### czytaj_config() — rozszerzenie

```python
def czytaj_config(wb):
    cfg = wb.sheets["Config"]
    return {
        "api_key":    cfg["B1"].value,
        "model":      cfg["B2"].value,       # claude-sonnet-4-6
        "max_tokens": int(cfg["B3"].value),  # 600
        "days_back":  int(cfg["B4"].value),  # 30
        "tryb":       cfg["B6"].value or "LAST_N",  # LAST_N / DATE_RANGE
        "data_od":    cfg["B7"].value,               # datetime lub None
        "data_do":    cfg["B8"].value,               # datetime lub None
    }
```

### pobierz_dane_master() — logika wyboru okna

```python
def pobierz_dane_master(wb, config):
    ws  = wb.sheets["qry_MASTER"]
    df  = ws.range("A1").expand("table").options(pd.DataFrame, header=True).value
    df.index = pd.to_datetime(df.index)

    if config["tryb"] == "DATE_RANGE" and config["data_od"] and config["data_do"]:
        d_od = pd.to_datetime(config["data_od"])
        d_do = pd.to_datetime(config["data_do"])
        df_okno = df[(df.index >= d_od) & (df.index <= d_do)]
    else:
        df_okno = df.tail(config["days_back"])

    return df, df_okno
```

**Uwaga:** `generuj_komentarz_ai()` nie wymaga zmian — woła te dwie funkcje, reszta działa automatycznie. AI_Summary i Komentarz_AI per-row generują się dla wybranego okna.

---

## Zmiany w VBA: dwa nowe makra

### AnalizujOkres() — tryb DATE_RANGE

```vba
Sub AnalizujOkres()
    Dim cfg As Worksheet
    Set cfg = ThisWorkbook.Sheets("Config")

    Dim d_od As String, d_do As String
    d_od = InputBox("Data OD (RRRR-MM-DD):", "SilesiaPulse — Analiza okresu", "2022-02-24")
    If d_od = "" Then Exit Sub
    d_do = InputBox("Data DO (RRRR-MM-DD):", "SilesiaPulse — Analiza okresu", "2022-06-30")
    If d_do = "" Then Exit Sub

    cfg.Range("B6").Value = "DATE_RANGE"
    cfg.Range("B7").Value = d_od
    cfg.Range("B8").Value = d_do

    Dim cmd As String
    cmd = "python " & Chr(34) & "C:\Projekty\uek-silesiapulse\python\silesiapulse_ai.py" & Chr(34)
    Shell "cmd.exe /c " & cmd, vbMinimizedNoFocus

    MsgBox "Analiza okresu " & d_od & " — " & d_do & " uruchomiona." & vbCrLf & _
           "Wyniki pojawia sie w AI_Summary i kolumnie Komentarz_AI.", _
           vbInformation, "SilesiaPulse"
End Sub
```

### PrzywrocTrybDomyslny() — powrót do LAST_N

```vba
Sub PrzywrocTrybDomyslny()
    ThisWorkbook.Sheets("Config").Range("B6").Value = "LAST_N"
    MsgBox "Tryb przywrocony: ostatnie " & _
           ThisWorkbook.Sheets("Config").Range("B4").Value & " dni.", _
           vbInformation, "SilesiaPulse"
End Sub
```

---

## Zmiany w Power Query: kolumna Kryzys

Nowa kolumna w `qry_MASTER` (dodana w Power Query M jako krok własny po scaleniu):

```m
= Table.AddColumn(PreviousStep, "Kryzys", each
    if [Data] >= #date(2020,3,1)  and [Data] <= #date(2020,6,30)  then "COVID"
    else if [Data] >= #date(2022,2,24) and [Data] <= #date(2022,6,30) then "Ukraina"
    else if [Data] >= #date(2023,10,7) and [Data] <= #date(2023,12,31) then "Czarna Sobota"
    else if [Data] >= #date(2026,2,1)  and [Data] <= #date(2026,4,30)  then "Epic Fury"
    else "Bez kryzysu")
```

Ta kolumna zasila slicer predefiniowany w Power BI — klik „Ukraina" filtruje wszystkie wykresy do tego okresu.

---

## Zmiany w Power BI: 3 typy slicerów daty

| Typ | Konfiguracja | Zastosowanie |
|---|---|---|
| **Relative date** | Slicer → Data → Relative date | Ostatni tydzień / miesiąc / rok |
| **Between** | Slicer → Data → Between | Dowolny zakres od-do (suwak) |
| **Kryzys** | Slicer → kolumna Kryzys | Przyciski: COVID / Ukraina / Czarna Sobota / Epic Fury |

Wszystkie trzy słuchają tej samej osi `Data` z relacji `qry_Kalendarz[Data] → qry_MASTER[Data]`.

---

## Mapa komórek Panelu sterowania — stan po rozszerzeniu

```
C14 → data ostatniego odświeżenia danych    (OdswierzDane VBA)
C15 → data ostatniego komentarza AI         (silesiapulse_ai.py)
C16 → wiersze przetworzone                  (silesiapulse_ai.py)
C17 → OK / BŁĄD  ← formatowanie warunkowe
```

Proponowane nowe komórki:
```
C18 → aktywny tryb: LAST_N / DATE_RANGE     (AnalizujOkres VBA)
C19 → data OD (gdy DATE_RANGE)              (AnalizujOkres VBA)
C20 → data DO (gdy DATE_RANGE)              (AnalizujOkres VBA)
```

---

## Workflow demo na obronę (sekwencja 6 kroków)

1. Klik **„Analizuj okres"** → wpisz `2022-02-24` i `2022-06-30`
2. Python generuje w tle: AI_Summary dla okresu + Komentarz_AI per dzień
3. W Power BI: slicer **„Ukraina"** (lub zakres dat manualnie)
4. Wykresy GPR + VIX + KGHM filtrują się → widać dynamikę kryzysu
5. Zakładka **AI_Summary**: zbiorczy komentarz Claude dla całego okresu
6. Zakładka **qry_MASTER**: per-row komentarze dla każdego dnia z kryzysu

→ Demonstracja poziomów prognostycznego i preskryptywnego (Kania) jednocześnie.

---

## Ograniczenia do zaznaczenia na obronę

- Slicer Power BI **nie regeneruje** Komentarz_AI — filtruje istniejące wiersze
- Komentarz zbiorczy (AI_Summary) generuje się na żądanie przez `AnalizujOkres()`
- Dla bardzo wąskich okien (1 dzień) Python generuje 1 komentarz per-row + 1 zbiorczy — sensowne minimum
- `DATE_RANGE` zapisuje się do Config — po sesji analitycznej wywołaj `PrzywrocTrybDomyslny()` lub `OdswierzWszystko()` automatycznie przywróci LAST_N (do zdecydowania przy implementacji)

---

## Stan implementacji

| Element | Status |
|---|---|
| Config B6/B7/B8 — nowe komórki | ⬜ NIE zaimplementowane |
| czytaj_config() — rozszerzenie | ⬜ NIE zaimplementowane |
| pobierz_dane_master() — logika trybu | ⬜ NIE zaimplementowane |
| AnalizujOkres() VBA | ⬜ NIE zaimplementowane |
| PrzywrocTrybDomyslny() VBA | ⬜ NIE zaimplementowane |
| Kolumna Kryzys w qry_MASTER | ⬜ NIE zaimplementowane |
| Slicery Power BI | ⬜ (Faza 6 — Power BI jeszcze nie istnieje) |

---

*Nexus · SilesiaPulse · 25.05.2026 · sesja kontroli okresu analizy*  
*"Zaprojektowane razem — do wdrożenia w osobnej sesji przed 28.05.2026"*
