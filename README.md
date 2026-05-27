# Behavioral Finance Portfolio Backtest

Dieses Repository enthält ausschließlich den R-Code zur empirischen Analyse der
Bachelorarbeit **„Behavioral Finance im Portfoliomanagement: Die Rolle von Fear of Loss und Fear of Missing Out im Asset Management“**. 
Die zugehörigen Excel-Daten, Ergebnisdateien, Grafiken und PDF-Exports werden aus
Lizenz-, Speicher- und Reproduzierbarkeitsgründen nicht versioniert.

Der Code vergleicht klassische und verhaltensorientierte Portfoliomodelle auf
Basis historischer S&P-500-Daten. Im Mittelpunkt stehen die Behavioral-Finance-
Komponenten **Fear of Loss (FOL)** und **Fear of Missing Out (FOMO)**.

## Repository-Inhalt

Im Git-Repository sollen nur folgende Dateien liegen:

```text
.
├── Modellvergleich.R
├── Auswertung Ergebnisse Modellvergleich.R
├── README.md
└── .gitignore
```

Nicht Bestandteil des Repositorys sind insbesondere:

- `S&P_500_Daten.xlsx`
- `Backtest_Ergebnisse.xlsx`
- `Backtest_Auswertung.xlsx`
- `Backtest_Charts.pdf`
- der Ordner `Plots/`
- Word-, PowerPoint-, PDF-, CSV- und sonstige Arbeitsdateien

Diese Dateien werden lokal benötigt oder lokal erzeugt, aber nicht in Git
gespeichert.

## Skripte

| Datei | Zweck |
| --- | --- |
| `Modellvergleich.R` | Hauptskript für Datenimport, Portfoliooptimierung, Backtest, Benchmark-Download und Export der Roh-Ergebnisse. |
| `Auswertung Ergebnisse Modellvergleich.R` | Auswertungsskript für Kennzahlen, Tabellen, Grafiken, Excel-Bericht und Sammel-PDF. |

## Lokale Dateien

Damit der Code ausgeführt werden kann, muss die Eingabedatei lokal vorhanden
sein:

```text
S&P_500_Daten.xlsx
```

Diese Datei wird nicht mitversioniert. Sie muss vor dem Start von
`Modellvergleich.R` im lokalen Arbeitsordner liegen oder der Pfad `BASE_PATH`
im Skript muss entsprechend angepasst werden.

Nach der Ausführung entstehen lokal folgende Dateien:

```text
Backtest_Ergebnisse.xlsx
Backtest_Auswertung.xlsx
Backtest_Charts.pdf
Plots/*.png
```

Auch diese Outputs bleiben außerhalb von Git.

## Datenbeschaffung

Die Rohdaten werden nicht im Repository bereitgestellt. Wer den Code
reproduzieren möchte, muss die Datei `S&P_500_Daten.xlsx` selbst erzeugen.
In der Bachelorarbeit wurden die Einzeltiteldaten über Datastream/LSEG bezogen.

Verwendet wurden:

- historische S&P-500-Einzeltitel bzw. das für die Analyse definierte
  S&P-500-Universum
- tägliche Return-Index-Zeitreihen (`RI`) je Aktie
- zusätzliche statische Felder zur Kontrolle bzw. Dokumentation:
  - `RI`: Return Index
  - `MV`: Market Value
  - `UP`: Price bzw. Unadjusted Price
  - `VO`: Volume
- jährliche Rebalancing-Listen für 2016 bis 2026
- S&P 500 Total Return Index als Benchmark; dieser wird im Code separat über
  `quantmod` von Yahoo Finance als `^SP500TR` geladen

Der wichtigste Input für den Backtest ist der Datastream-Return-Index (`RI`),
weil daraus die diskreten Tagesrenditen berechnet werden:

```text
R_t = RI_t / RI_{t-1} - 1
```

Für die erste Rebalancing-Periode 2016 werden mindestens zwei Jahre Historie
benötigt. Die Daten sollten deshalb mindestens bis Anfang 2014 zurückreichen
und bis zum Ende des Untersuchungszeitraums laufen.

## Aufbau der lokalen Datendatei

Der Code erwartet lokal eine Excel-Datei mit dem Namen:

```text
S&P_500_Daten.xlsx
```

Diese Datei muss im lokalen `BASE_PATH` liegen oder der Pfad im Skript
`Modellvergleich.R` muss angepasst werden.

### Sheet `S&P500`

Das Hauptsheet enthält die täglichen Return-Index-Zeitreihen der Aktien. Der
Code liest das Sheet in zwei Schritten ein: zuerst die Namen aus Zeile 1, danach
die Daten ab Zeile 6.

Erwarteter Aufbau:

```text
Zeile 1:  Formel-/Metatext in Spalte A, Unternehmens- oder Tickerspalten ab B
Zeile 2:  "RI" in Spalte A, letzter RI-Wert je Aktie ab B
Zeile 3:  "MV" in Spalte A, letzter MV-Wert je Aktie ab B
Zeile 4:  "UP" in Spalte A, letzter UP-Wert je Aktie ab B
Zeile 5:  "VO" in Spalte A, letztes Volumen je Aktie ab B
Zeile 6+: Datum in Spalte A, tägliche RI-Zeitreihe je Aktie ab B
```

Wichtig ist, dass die Spaltenreihenfolge der Aktien in diesem Sheet konsistent
bleibt. Die Werte ab Zeile 6 werden vom Skript als numerische RI-Zeitreihe
interpretiert.

### Sheets `Rebalancing YYYY`

Zusätzlich erwartet der Code je Jahr ein Rebalancing-Sheet:

```text
Rebalancing 2016
Rebalancing 2017
...
Rebalancing 2026
```

Diese Sheets dienen dazu, das investierbare Universum am jeweiligen
Rebalancing-Stichtag zu bestimmen.

Erwarteter Aufbau:

```text
Zeile 1:  Formel-/Metatext in Spalte A, Unternehmens- oder Tickerspalten ab B
Zeile 4:  Jahresendpreis bzw. gültiger Preis je Aktie ab B
```

Der Code liest aus jedem Rebalancing-Sheet die Ticker aus Zeile 1 und die
Jahresendpreise aus Zeile 4. Ein gültiger Preis bedeutet, dass der Titel für
das betreffende Jahr als investierbar berücksichtigt werden kann. Fehlende
Preise werden als nicht investierbar interpretiert.

Zusätzlich filtert der Code später nochmals nach Datenqualität:

- mindestens `MIN_HISTORY_DAYS = 504` gültige Handelstage vor dem
  Rebalancing-Stichtag
- höchstens `MAX_NA_SHARE = 5 %` fehlende Werte im jüngsten Historienfenster
- keine vollständig fehlenden oder praktisch konstanten Renditereihen

Dadurch wird verhindert, dass Aktien mit zu kurzer Historie oder instabiler
Datenlage in die Optimierung eingehen.

## Nachvollziehbarer Datenworkflow

Der Datenworkflow war:

1. Historisches S&P-500-Universum bzw. die verwendeten S&P-500-Titel für die
   Analyse festlegen.
2. Für alle Aktien tägliche Datastream-Return-Index-Reihen (`RI`) für den
   gesamten Untersuchungszeitraum herunterladen.
3. Zusätzlich `MV`, `UP` und `VO` als Kontroll-/Metafelder aus Datastream
   exportieren.
4. Die Daten in die beschriebene Excel-Struktur `S&P_500_Daten.xlsx`
   übertragen.
5. Für jedes Rebalancing-Jahr 2016 bis 2026 ein Sheet `Rebalancing YYYY`
   anlegen.
6. Im jeweiligen Rebalancing-Sheet nur für die am Stichtag investierbaren Titel
   einen gültigen Preis hinterlegen; fehlende Preise bleiben `NA`.
7. `Modellvergleich.R` ausführen. Das Skript berechnet aus den RI-Reihen die
   Tagesrenditen, filtert das investierbare Universum und erzeugt lokal
   `Backtest_Ergebnisse.xlsx`.
8. `Auswertung Ergebnisse Modellvergleich.R` ausführen. Das Skript erstellt
   daraus lokal die Ergebnisgrafiken, den Excel-Auswertungsbericht und das
   Sammel-PDF.

Da Datastream/LSEG-Daten lizenzpflichtig sein können, werden diese Rohdaten
nicht veröffentlicht. Die genaue Datenstruktur ist jedoch dokumentiert, sodass
die Datei mit einem eigenen Datenzugang reproduziert werden kann.

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
geklont, muss dieser Pfad im Setup-Block angepasst werden.

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

## Output

`Modellvergleich.R` erzeugt lokal:

```text
Backtest_Ergebnisse.xlsx
```

Diese Datei enthält unter anderem:

- Out-of-Sample-Renditen
- Gesamtmetriken
- Krisen- und Boomphasenmetriken
- Constraint-Diagnostik
- Portfolio-Gewichte
- Methoden-README

`Auswertung Ergebnisse Modellvergleich.R` erzeugt lokal:

```text
Backtest_Auswertung.xlsx
Backtest_Charts.pdf
Plots/*.png
```

Diese Ausgaben sind Ergebnisartefakte und werden nicht in Git gespeichert.

## Wichtige Limitationen

- Der risikofreie Zinssatz wird konstant mit 2 % p.a. angenommen.
- Die direkte Sample-Schiefe-Optimierung ist empirisch sensibel gegenüber
  Tail-Beobachtungen und nicht-konvex.
- Die Corona-Stressphase wird bewusst breiter als der reine Crash-Zeitraum
  definiert: 01.02.2020 bis 30.06.2020.
- Ein möglicher Survivorship Bias wird nicht vollständig korrigiert, da die
  lokale Datei `S&P_500_Daten.xlsx` als Datengrundlage verwendet wird.
- Die S&P-500-Einzeltiteldaten können lizenzrechtlichen Beschränkungen
  unterliegen und werden deshalb nicht im Repository abgelegt.

## Reproduzierbarkeit

Zur Reproduzierbarkeit wird im Hauptskript `set.seed(42)` gesetzt. Das betrifft
insbesondere die zufallsbasierten Kandidaten der Local Search im Behavioral-
Modell. Bei `FOMO_FOL_SOLVER = "auto"` gilt:

- DEoptim wird verwendet, wenn das reduzierte Universum höchstens
  `FOMO_FOL_DE_N_MAX` Titel enthält.
- Bei größeren Universen wird Local Search mit derselben Zielfunktion genutzt.

Die Solverwahl wird im Konsolenoutput und lokal im Sheet `README` der Datei
`Backtest_Ergebnisse.xlsx` dokumentiert.

## Git-Hinweis

Die Datei `.gitignore` ist so ausgelegt, dass Excel-Daten, Ergebnisexports,
Grafiken, Office-Dateien, temporäre Render-Ordner und R-Arbeitsdateien nicht
versehentlich versioniert werden. Das Repository bleibt dadurch ein reines
Code-Repository.

## Autor

Yannick Jungbauer  
Bachelorarbeit: Behavioral Finance im Portfoliomanagement: Die Rolle von Fear of Loss und Fear of Missing Out 
im Asset Management

