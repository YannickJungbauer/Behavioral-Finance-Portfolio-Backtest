# Behavioral Finance Portfolio Backtest

Dieses Repository enthält die Code-, Daten- und Ergebnisdateien zur
Bachelorarbeit **„Behavioral Finance im Portfoliomanagement: Die Rolle von Fear of Loss und Fear of Missing Out 
im Asset Management
“**. Der empirische
Teil vergleicht klassische und verhaltensorientierte Portfoliomodelle auf Basis
historischer S&P-500-Daten.

Im Mittelpunkt stehen die Behavioral-Finance-Komponenten **Fear of Loss (FOL)**
und **Fear of Missing Out (FOMO)**. Der Code führt den Backtest durch, erzeugt
Out-of-Sample-Renditen, berechnet Performance- und Risikokennzahlen und erstellt
Excel-Auswertungen, Grafiken sowie ein Sammel-PDF.

## Repository-Struktur

```text
.
├── Modellvergleich.R
├── Auswertung Ergebnisse Modellvergleich.R
├── S&P_500_Daten.xlsx
├── Backtest_Ergebnisse.xlsx
├── Backtest_Auswertung.xlsx
├── Backtest_Charts.pdf
├── README.md
└── Plots/
    ├── 01_Equity_Curves.png
    ├── 02_Drawdown.png
    ├── 03_Rolling_Volatility.png
    ├── 04_Annual_Returns.png
    ├── 05_Metrics_Heatmap.png
    ├── 06_Krise_Boom.png
    ├── 07_Konzentration_EffN.png
    ├── 08_Konzentration_AnzPos.png
    ├── 09_Top10_Holdings.png
    ├── 10_Skew_Semi_Scatter.png
    ├── 11_Renditeverteilungen.png
    ├── 12_Sharpe_vs_Sortino.png
    ├── 13_Outperformance.png
    ├── 14_Korrelationsmatrix.png
    └── 15_Krise_Boom_Bars.png
```

## Dateiübersicht

| Datei | Inhalt |
| --- | --- |
| `Modellvergleich.R` | Hauptskript für Datenimport, Portfoliooptimierung, Backtest, Benchmark-Download und Export der Roh-Ergebnisse. |
| `Auswertung Ergebnisse Modellvergleich.R` | Auswertungsskript für Kennzahlen, Tabellen, Grafiken, Excel-Bericht und Sammel-PDF. |
| `S&P_500_Daten.xlsx` | Eingabedatei mit S&P-500-Einzeltiteldaten und Rebalancing-Sheets. |
| `Backtest_Ergebnisse.xlsx` | Roh-Output des Backtests mit Out-of-Sample-Renditen, Metriken, Diagnostik und Portfolio-Gewichten. |
| `Backtest_Auswertung.xlsx` | Präsentationsfertige Excel-Auswertung mit formatierten Tabellen und eingebetteten Grafiken. |
| `Backtest_Charts.pdf` | Sammel-PDF mit den wichtigsten Ergebnisgrafiken. |
| `Plots/` | Einzelne PNG-Grafiken der empirischen Auswertung. |

## Voraussetzungen

Der Code wurde für R ab Version 4.2 entwickelt und zuletzt mit R 4.5.2 geprüft.

Benötigte R-Pakete:

```r
install.packages(c(
  "readxl", "xts", "zoo", "PerformanceAnalytics",
  "quadprog", "openxlsx", "quantmod", "DEoptim",
  "dplyr", "tidyr", "ggplot2", "scales"
))
```

`quantmod` wird für den Download des S&P 500 Total Return Index (`^SP500TR`)
verwendet. Dafür ist eine Internetverbindung notwendig.

## Ausführung

Die Skripte sind in dieser Reihenfolge auszuführen:

```r
source("Modellvergleich.R")
source("Auswertung Ergebnisse Modellvergleich.R")
```

Alternativ über die Konsole:

```bash
Rscript "Modellvergleich.R"
Rscript "Auswertung Ergebnisse Modellvergleich.R"
```

Wichtig: In beiden Skripten ist `BASE_PATH` aktuell auf
`F:/FH/Bachelorarbeit/` gesetzt. Wird das Repository an einen anderen Ort
kopiert oder geklont, muss dieser Pfad im Setup-Block der beiden Skripte
angepasst werden.

## Datengrundlage

Die Datei `S&P_500_Daten.xlsx` enthält die für den Backtest verwendeten
S&P-500-Einzeltiteldaten. Die Daten wurden über Datastream/LSEG bezogen und in
eine für den Code lesbare Excel-Struktur gebracht.

Verwendet wurden:

- tägliche Return-Index-Zeitreihen (`RI`) je Aktie
- zusätzliche Kontroll- und Metafelder:
  - `RI`: Return Index
  - `MV`: Market Value
  - `UP`: Price bzw. Unadjusted Price
  - `VO`: Volume
- jährliche Rebalancing-Sheets für 2016 bis 2026
- S&P 500 Total Return Index als Benchmark; dieser wird im Code über
  `quantmod` von Yahoo Finance als `^SP500TR` geladen

Der zentrale Input für die Renditeberechnung ist der Return Index:

```text
R_t = RI_t / RI_{t-1} - 1
```

## Aufbau der Datei `S&P_500_Daten.xlsx`

### Sheet `S&P500`

Das Hauptsheet enthält die täglichen Return-Index-Zeitreihen der Aktien.

Erwarteter Aufbau:

```text
Zeile 1:  Formel-/Metatext in Spalte A, Unternehmens- oder Tickerspalten ab B
Zeile 2:  "RI" in Spalte A, letzter RI-Wert je Aktie ab B
Zeile 3:  "MV" in Spalte A, letzter MV-Wert je Aktie ab B
Zeile 4:  "UP" in Spalte A, letzter UP-Wert je Aktie ab B
Zeile 5:  "VO" in Spalte A, letztes Volumen je Aktie ab B
Zeile 6+: Datum in Spalte A, tägliche RI-Zeitreihe je Aktie ab B
```

Der Code liest zuerst die Spaltennamen aus Zeile 1 und anschließend die
eigentlichen RI-Zeitreihen ab Zeile 6.

### Sheets `Rebalancing YYYY`

Zusätzlich enthält die Datei je Jahr ein Rebalancing-Sheet:

```text
Rebalancing 2016
Rebalancing 2017
...
Rebalancing 2026
```

Diese Sheets werden genutzt, um das investierbare Universum zum jeweiligen
Rebalancing-Stichtag zu bestimmen.

Erwarteter Aufbau:

```text
Zeile 1:  Formel-/Metatext in Spalte A, Unternehmens- oder Tickerspalten ab B
Zeile 4:  Jahresendpreis bzw. gültiger Preis je Aktie ab B
```

Ein gültiger Preis bedeutet, dass der Titel im jeweiligen Jahr als investierbar
berücksichtigt werden kann. Fehlende Preise werden als nicht investierbar
interpretiert.

Zusätzlich filtert der Code nach Datenqualität:

- mindestens `MIN_HISTORY_DAYS = 504` gültige Handelstage vor dem
  Rebalancing-Stichtag
- höchstens `MAX_NA_SHARE = 5 %` fehlende Werte im jüngsten Historienfenster
- keine vollständig fehlenden oder praktisch konstanten Renditereihen

## Backtest-Design

- Zeitraum Out-of-Sample: 2017 bis 2026
- Rebalancing: jährlich auf Basis eines Expanding Windows
- Anlageuniversum: S&P-500-Konstituenten je Rebalancing-Jahr
- Benchmark: S&P 500 Total Return Index (`SP500_TR`)
- Risikofreier Zinssatz: konstant 2 % p.a.
- Gemeinsame Constraints:
  - long-only
  - voll investiert
  - maximal 10 % Gewicht je Einzeltitel
  - identische Constraints für alle Modellportfolios

Die Out-of-Sample-Renditen werden als Strategie mit konstanten Zielgewichten je
Haltedauer berechnet. Mathematisch entspricht `R_oos %*% w` einer täglichen
Rückführung auf diese Zielgewichte und ist daher kein klassisches
Buy-and-Hold-Portfolio mit driftenden Einzeltitelgewichten.

## Modellübersicht

### MaxSharpe

Klassisches Mean-Variance- bzw. Tangency-Portfolio. Optimiert wird die
Sharpe-Ratio über eine quadratische Programmierung mit `quadprog`.

### Behavioral

Das Behavioral-Hauptmodell optimiert FOMO und FOL aktiv in einer gemeinsamen
Zielfunktion:

```text
max Skew(w) / SemiDev_0(w)
```

Dabei gilt:

- FOMO wird über positive Portfolio-Schiefe gemessen.
- FOL wird über annualisierte Semideviation echter Verluste unter 0 gemessen.
- Die ökonomische Interpretation lautet: positive Schiefe pro Einheit
  Downside-Risiko.
- Negative Schiefe wird in der Zielfunktion über eine Penalty behandelt, damit
  keine Vorzeichenfalle entsteht.

Die direkte Optimierung von Sample-Schiefe bleibt schätzsensitiv und
nicht-konvex. Dieser Punkt ist methodisch bewusst dokumentiert und wird in der
Interpretation als empirischer Trade-off diskutiert.

### Arnott

Der Arnott-Blend dient als Robustheitsmodell. Er kombiniert:

- FOL-Schicht: defensives Min-Semicovariance-Portfolio
- FOMO-Schicht: 1/N-Marktproxy
- zeitvariabler Alpha-Faktor auf Basis jüngster gegenüber vollständiger
  Semideviation

Dieses Modell optimiert Schiefe nicht direkt, sondern bildet FOMO passiv über
Marktpartizipation ab.

### Hybrid

Das Hybrid-Portfolio ist ein transparenter linearer Blend aus MaxSharpe und
Behavioral:

```text
Hybrid = 50 % MaxSharpe + 50 % Behavioral
```

Es dient dazu, die Forschungsfrage „ersetzen oder ergänzen“ empirisch
nachvollziehbar zu beantworten.

## Ergebnisdateien

`Modellvergleich.R` erzeugt bzw. aktualisiert:

```text
Backtest_Ergebnisse.xlsx
```

Diese Datei enthält:

- `README`
- `Returns_OOS`
- `Metrics_Overall`
- `Metrics_Crisis_Corona`
- `Metrics_Boom_Tech23`
- `Diagnostics`
- `Weights_MaxSharpe`
- `Weights_Behavioral`
- `Weights_Arnott`
- `Weights_Hybrid`

`Auswertung Ergebnisse Modellvergleich.R` erzeugt bzw. aktualisiert:

```text
Backtest_Auswertung.xlsx
Backtest_Charts.pdf
Plots/*.png
```

Die Auswertungsdatei enthält 16 formatierte Sheets, darunter Executive Summary,
Wertentwicklung, Risiko-Profil, Detailmetriken, Jahresrenditen, Teilperioden,
Konzentrationsanalyse, Top-Holdings, Chart-Galerie und Gewichtungsmatrizen.

## Wichtige Limitationen

- Der risikofreie Zinssatz wird konstant mit 2 % p.a. angenommen.
- Die direkte Sample-Schiefe-Optimierung ist empirisch sensibel gegenüber
  Tail-Beobachtungen und nicht-konvex.
- Die Corona-Stressphase wird bewusst breiter als der reine Crash-Zeitraum
  definiert: 01.02.2020 bis 30.06.2020.
- Ein möglicher Survivorship Bias wird nicht korrigiert.
  
## Reproduzierbarkeit

Zur Reproduzierbarkeit wird im Hauptskript `set.seed(42)` gesetzt. Das betrifft
insbesondere die zufallsbasierten Kandidaten der Local Search im Behavioral-
Modell. Bei `FOMO_FOL_SOLVER = "auto"` gilt:

- DEoptim wird verwendet, wenn das reduzierte Universum höchstens
  `FOMO_FOL_DE_N_MAX` Titel enthält.
- Bei größeren Universen wird Local Search mit derselben Zielfunktion genutzt.

Die Solverwahl wird im Konsolenoutput und im Sheet `README` der Datei
`Backtest_Ergebnisse.xlsx` dokumentiert.

## Git-Hinweis

Die ursprünglichen Daten- und Ergebnisdateien sind in diesem Repository
enthalten. Temporäre Dateien, R-Arbeitsdateien, Office-Lockdateien und lokale
Render-Ordner sollten dagegen nicht versioniert werden.

## Autor

Yannick Jungbauer  
Bachelorarbeit: Behavioral Finance im Portfoliomanagement: Die Rolle von Fear of Loss und Fear of Missing Out 
im Asset Management


