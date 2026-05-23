[README.md](https://github.com/user-attachments/files/28179041/README.md)
# Behavioral-Finance-Portfolio-Backtest
Dieses Repository enthält den R-Code, die Eingabedaten und die Ergebnisdateien zur empirischen Portfolioanalyse der Bachelorarbeit „Behavioral Finance im Portfoliomanagement: Die Rolle von Fear of Loss und Fear of Missing Out im Asset Management".
# Behavioral Finance Portfolio Backtest

Dieses Repository enthaelt die digitalen Anhangsdateien zur Bachelorarbeit
**"Behavioral Finance im Portfoliomanagement: Die Rolle von Fear of Loss und Fear of Missing Out im Asset Management"**. Der empirische Teil vergleicht
klassische und verhaltensorientierte Portfolioansaetze auf Basis historischer
S&P-500-Daten. Im Mittelpunkt stehen klassische Rendite-Risiko-Logiken sowie
die Operationalisierung von **Fear of Loss (FOL)** ueber Semivarianz bzw.
Semideviation und **Fear of Missing Out (FOMO)** ueber positive Schiefe.

Der Code dient der Nachvollziehbarkeit der Datenaufbereitung,
Portfoliokonstruktion, Out-of-Sample-Auswertung und grafischen Darstellung der
in der Arbeit berichteten Ergebnisse.

## Inhalt des Repositorys

```text
.
+-- Modellvergleich.R
+-- Auswertung Ergebnisse Modellvergleich.R
+-- S&P_500_Daten.xlsx
+-- Backtest_Ergebnisse.xlsx
+-- Backtest_Auswertung.xlsx
+-- Backtest_Charts.pdf
+-- Plots/
    +-- 01_Equity_Curves.png
    +-- 02_Drawdown.png
    +-- 03_Rolling_Volatility.png
    +-- 04_Annual_Returns.png
    +-- 05_Metrics_Heatmap.png
    +-- 06_Krise_Boom.png
    +-- 07_Konzentration_EffN.png
    +-- 08_Konzentration_AnzPos.png
    +-- 09_Top10_Holdings.png
    +-- 10_Skew_Semi_Scatter.png
    +-- 11_Renditeverteilungen.png
    +-- 12_Sharpe_vs_Sortino.png
    +-- 13_Outperformance.png
    +-- 14_Korrelationsmatrix.png
    +-- 15_Krise_Boom_Bars.png
```

## Dateiuebersicht

- `Modellvergleich.R`: Fuehrt den eigentlichen Backtest durch. Das Skript
  liest `S&P_500_Daten.xlsx`, konstruiert die Modellportfolios und erzeugt
  `Backtest_Ergebnisse.xlsx`.

- `Auswertung Ergebnisse Modellvergleich.R`: Liest `Backtest_Ergebnisse.xlsx`
  ein und erstellt die aufbereitete Ergebnisdatei `Backtest_Auswertung.xlsx`,
  die Chart-PDF `Backtest_Charts.pdf` sowie die PNG-Grafiken im Ordner
  `Plots/`.

- `S&P_500_Daten.xlsx`: Eingabedatei mit S&P-500-Daten. Die Datei enthaelt ein
  Hauptsheet `S&P500` sowie Rebalancing-Sheets fuer die Jahre 2016 bis 2026.
  Die Datenbasis beruht auf ueber Datastream/LSEG bezogenen Einzeltiteldaten.

- `Backtest_Ergebnisse.xlsx`: Rohere Ergebnisdatei des Backtests. Enthalten
  sind unter anderem Out-of-Sample-Renditen, Gesamtmetriken, Phasenmetriken,
  Diagnostikdaten und Portfolio-Gewichte.

- `Backtest_Auswertung.xlsx`: Aufbereitete Ergebnisdatei fuer die Interpretation
  der empirischen Analyse. Die Datei enthaelt zusammenfassende Tabellen,
  Kennzahlen, Rohdaten, Gewichtungsmatrizen und eingebettete Grafiken.

- `Backtest_Charts.pdf`: Sammel-PDF mit den wichtigsten grafischen
  Ergebnisdarstellungen.

- `Plots/`: Ordner mit den einzeln exportierten Grafiken im PNG-Format.

## Methodischer Kurzueberblick

Der Backtest verwendet eine Expanding-Window-Logik mit jaehrlichem Rebalancing.
Die Out-of-Sample-Phase beginnt am 1. Januar 2017 und endet am 20. April 2026.
Als Benchmark wird der S&P 500 Total Return Index (`^SP500TR`) ueber `quantmod`
eingebunden.

Verglichen werden folgende Strategien:

- `MaxSharpe`: Klassisches Mean-Variance-Portfolio bzw. Tangency-Portfolio auf
  Basis der Sharpe Ratio. Die Optimierung erfolgt mit `quadprog`.

- `Behavioral`: Aktives Behavioral-Portfolio. FOMO wird ueber positive
  Portfolio-Schiefe operationalisiert, FOL ueber Semideviation unter dem
  Referenzwert null. Die Zielfunktion maximiert das Verhaeltnis aus Schiefe und
  Semideviation.

- `Arnott`: Robustheitsmodell mit FOL-MinSemiCov-Komponente und passiver
  FOMO-Schicht ueber eine gleichgewichtete Marktproxy-Komponente.

- `Hybrid`: Lineare Kombination aus `MaxSharpe` und `Behavioral` mit einem
  Hybrid-Anteil von 50 Prozent.

- `SP500_TR`: S&P 500 Total Return als marktbreite Benchmark.

Alle optimierten Modellportfolios verwenden dieselben Grundrestriktionen:

- long-only
- voll investiert
- maximales Einzelgewicht von 10 Prozent
- mindestens 504 Handelstage Historie vor dem jeweiligen Rebalancing-Zeitpunkt
- maximal 5 Prozent fehlende Werte im juengsten Historienfenster
- konstanter risikofreier Zinssatz von 2 Prozent p. a.

## Voraussetzungen

Die Skripte sind fuer R ab Version 4.2 ausgelegt. Benoetigt werden die folgenden
R-Pakete:

```r
install.packages(c(
  "readxl",
  "xts",
  "zoo",
  "PerformanceAnalytics",
  "quadprog",
  "openxlsx",
  "quantmod",
  "DEoptim",
  "dplyr",
  "tidyr",
  "ggplot2",
  "scales"
))
```

Hinweis: `DEoptim` wird im Skript nur verwendet, wenn es verfuegbar ist und die
Universumsgroesse im Auto-Modus unter der festgelegten Schwelle liegt. Bei
groesseren Universen nutzt das Skript eine Local-Search-Logik fuer das aktive
Behavioral-Portfolio. Fuer die Benchmark wird eine Internetverbindung benoetigt,
da `quantmod` den S&P 500 Total Return Index von Yahoo Finance abruft.

## Ausfuehrung

Vor der Ausfuehrung muss in beiden R-Skripten der Pfad `BASE_PATH` auf den
lokalen Projektordner angepasst werden. Im vorliegenden Arbeitsstand ist dieser
Pfad auf `F:/FH/Bachelorarbeit/` gesetzt.

Die empfohlene Ausfuehrungsreihenfolge lautet:

```r
source("Modellvergleich.R")
source("Auswertung Ergebnisse Modellvergleich.R")
```

Alternativ koennen die Skripte ueber die Kommandozeile ausgefuehrt werden:

```bash
Rscript "Modellvergleich.R"
Rscript "Auswertung Ergebnisse Modellvergleich.R"
```

## Erwartete Outputs

Nach erfolgreicher Ausfuehrung von `Modellvergleich.R` entsteht:

- `Backtest_Ergebnisse.xlsx`

Nach erfolgreicher Ausfuehrung von `Auswertung Ergebnisse Modellvergleich.R`
entstehen:

- `Backtest_Auswertung.xlsx`
- `Backtest_Charts.pdf`
- der Ordner `Plots/` mit 15 PNG-Grafiken

## Ergebnisdateien

`Backtest_Ergebnisse.xlsx` enthaelt unter anderem folgende Sheets:

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

`Backtest_Auswertung.xlsx` bereitet diese Ergebnisse in auswertungsnaher Form
auf. Enthalten sind unter anderem eine Executive Summary, Wertentwicklung,
Risiko-Profil, Detailmetriken, Jahresrenditen, Marktphasenanalyse,
risikoadjustierte Kennzahlen, Konzentrationsanalysen, Top-Holdings,
Rohdaten-Renditen und Gewichtungsmatrizen.

## Hinweise zur Reproduzierbarkeit

- Das Skript `Modellvergleich.R` setzt `set.seed(42)`, um die zufallsbasierte
  Local-Search-Komponente reproduzierbarer zu machen.
- Die Ergebnisse koennen sich geringfuegig aendern, wenn externe Datenquellen
  wie Yahoo Finance nachtraeglich korrigiert werden oder wenn Paketversionen
  voneinander abweichen.
- Die Eingabedatei `S&P_500_Daten.xlsx` enthaelt lizenzierte Marktdaten. Eine
  oeffentliche Weitergabe kann daher eingeschraenkt sein.
- Dieses Repository dient der Dokumentation und Nachvollziehbarkeit der
  Bachelorarbeit und ersetzt nicht die methodische Beschreibung im Textteil der
  Arbeit.

## Bezug zur Bachelorarbeit

Die Dateien beziehen sich insbesondere auf die Kapitel zur Portfoliokonstruktion
und empirischen Analyse. Sie dokumentieren die technische Umsetzung des
Backtests, die Berechnung der Rendite-, Risiko- und Performancekennzahlen sowie
die grafische Aufbereitung der zentralen Ergebnisse.
