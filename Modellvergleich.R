################################################################################
#                                                                              #
#   BACHELORARBEIT: BEHAVIORAL PORTFOLIO OPTIMIZATION                          #
#   -----------------------------------------------------------------------    #
#   Vergleich MaxSharpe, Behavioral-FOMO/FOL, Arnott und Hybrid auf S&P 500    #
#   Backtest: Expanding-Window, Out-of-Sample 2017 - 2026                      #
#                                                                              #
#   Autor:   [Dein Name]                                                       #
#   Datum:   2026-04                                                           #
#   R-Ver.:  >= 4.2.0                                                          #
#                                                                              #
#   Ausfuehrung:                                                               #
#     1) Datei S&P_500_Daten.xlsx in BASE_PATH ablegen                         #
#     2) Skript komplett ausfuehren (source() oder Run-All)                    #
#     3) Ergebnisdatei wird in OUTPUT_PATH erzeugt                             #
#                                                                              #
################################################################################

# ==============================================================================
# 1. SETUP: PAKETE, PARAMETER, PFADE
# ==============================================================================

## ---- 1.1 Pakete -------------------------------------------------------------
# Bei Erstausfuehrung ggf. installieren:
# install.packages(c("readxl","xts","zoo","PerformanceAnalytics",
#                    "quadprog","openxlsx","quantmod","DEoptim"))

suppressPackageStartupMessages({
  library(readxl)              # Excel-Import
  library(xts)                 # Zeitreihen-Container
  library(zoo)                 # Hilfsfunktionen fuer Zeitreihen
  library(PerformanceAnalytics)# Performance-Metriken
  library(openxlsx)            # Excel-Export (write.xlsx)
})

# quantmod wird nur bei Bedarf geladen (Online-Benchmark)
has_quantmod <- suppressWarnings(suppressMessages(
  requireNamespace("quantmod", quietly = TRUE)
))

# User-Library explizit aufnehmen, falls Rscript sie nicht automatisch in
# .libPaths() setzt. Das hilft, wenn Pakete in RStudio installiert wurden.
user_lib <- Sys.getenv("R_LIBS_USER")
if (nzchar(user_lib) && dir.exists(user_lib) && !user_lib %in% .libPaths()) {
  .libPaths(c(user_lib, .libPaths()))
}
has_DEoptim <- suppressWarnings(suppressMessages(
  requireNamespace("DEoptim", quietly = TRUE)
))

## ---- 1.2 Parameter ----------------------------------------------------------
set.seed(42)  # Reproduzierbarkeit der zufallsbasierten Local-Search-Kandidaten

# --- Pfade ---
BASE_PATH   <- "F:/FH/Bachelorarbeit/"
DATA_PATH   <- paste0(BASE_PATH, "S&P_500_Daten.xlsx")
OUTPUT_PATH <- paste0(BASE_PATH, "Backtest_Ergebnisse.xlsx")

# --- Backtest-Parameter ---
REBAL_YEARS <- 2016:2025                 # Rebalancing-Stichtage (jeweils letzter Handelstag des Jahres)
OOS_START   <- as.Date("2017-01-01")     # Start der Out-of-Sample-Phase
OOS_END     <- as.Date("2026-04-20")     # Ende der Out-of-Sample-Phase (= letzter Datenpunkt)

# --- Risikofreie Rate (fuer Sharpe / Sortino) ---
# Vereinfachung: konstant. Fuer zeitvariable rf -> TBill-Zeitreihe laden.
RF_ANNUAL   <- 0.02                      # 2 % p.a.
RF_DAILY    <- (1 + RF_ANNUAL)^(1/252) - 1

# --- Universumsfilter (pragmatischer Datenverfuegbarkeitsfilter) ---
MIN_HISTORY_DAYS <- 504                  # min. 2 Jahre Historie vor Rebalancing-Datum
MAX_NA_SHARE     <- 0.05                 # max. 5 % NAs im juengsten Historienfenster

# --- Gemeinsame Portfolio-Constraints ---------------------------------------
# Diese Restriktionen gelten fuer ALLE Modellportfolios. Damit wird im Backtest
# nicht durch unterschiedliche Anlageregeln verzerrt. Long-only, voll
# investiert, maximal 10 % pro Einzeltitel.
# MIN_WEIGHT ist nur eine numerische Untergrenze fuer QP-/Projektionsschritte.
# Oekonomisch wird kein Mindestgewicht erzwungen: sehr kleine Gewichte sind
# praktisch Null und dienen nur der Solver-Stabilitaet.
MIN_WEIGHT <- 1e-6
MAX_WEIGHT <- 0.10
MIN_ASSETS_FOR_MAX_WEIGHT <- ceiling(1 / MAX_WEIGHT)

# --- Behavioral-Hauptmodell: aktive FOMO/FOL-Ratio ---------------------------
# Das Behavioral-Modell optimiert beide behavioralen Komponenten gemeinsam:
#
#   max_w  Skew(w) / SemiDev_0(w)
#
# FOMO wird ueber positive Portfolio-Schiefe (Fisher-Pearson, korrigiert)
# operationalisiert. FOL wird ueber annualisierte Semideviation echter Verluste
# unter 0 operationalisiert. Damit lautet die oekonomische Interpretation:
# "positive Schiefe pro Einheit Downside-Risiko".
#
# Die bekannten Stabilitaetsprobleme direkter Sample-Schiefe-Optimierung werden
# nicht ausgeblendet, sondern bewusst in der Diskussion aufgegriffen. Der
# Arnott-Blend bleibt deshalb zusaetzlich als Robustheitsmodell im Backtest.
# "DEoptim": DEoptim erzwingen (kann bei vollem S&P-500-Universum sehr lange dauern).
# "auto": DEoptim bis FOMO_FOL_DE_N_MAX, darueber Local Search.
# "local": immer Local Search mit gleicher Zielfunktion.
FOMO_FOL_SOLVER        <- "auto"
FOMO_FOL_DE_N_MAX      <- 120            # nur im auto-Modus: bis hier DEoptim, darueber Local Search
FOMO_FOL_DE_ITERMAX    <- 200
FOMO_FOL_LOCAL_TRIALS  <- 2500
FOMO_FOL_LOCAL_JITTER  <- 500
FOMO_FOL_LOCAL_TOPPOOL <- 120
FOMO_FOL_ACTIVE_MIN    <- MIN_ASSETS_FOR_MAX_WEIGHT
FOMO_FOL_ACTIVE_MAX    <- 40

# --- Robustheitsmodell nach Arnott & McQuarrie -------------------------------
# Passives Schiefe-Harvesting: FOL-Schicht = Min-Semicovariance-Portfolio,
# FOMO-Schicht = 1/N-Marktproxy, alpha_t aus aktueller vs. voller
# Semideviation. Dieses Modell ist NICHT das Behavioral-Hauptmodell, sondern
# dient als Vergleich "aktive FOMO/FOL-Optimierung vs. passives Harvesting".
ARNOTT_RECENT_WINDOW_DAYS <- 252
ARNOTT_ALPHA_BASE <- 0.50
ARNOTT_ALPHA_MIN  <- 0.20
ARNOTT_ALPHA_MAX  <- 0.80
BEHAVIORAL_METHOD_NOTE <- paste(
  "Das Behavioral-Hauptmodell maximiert aktiv Skew(w)/SemiDev_0(w).",
  "FOMO wird ueber korrigierte Fisher-Pearson-Schiefe gemessen, FOL ueber",
  "annualisierte Semideviation unter 0. Der Arnott-Blend bleibt als",
  "Robustheitsmodell fuer passives Schiefe-Harvesting enthalten."
)
SKEW_DISCUSSION_NOTE <- paste(
  "Direkte Sample-Schiefe-Optimierung bleibt tail-sensitiv und nicht-konvex.",
  "Wenn das aktive Behavioral-Modell konzentrierter ist, in Krisen schlechter",
  "abschneidet oder MaxSharpe OOS eine hoehere Schiefe zeigt, ist das kein",
  "Codefehler, sondern empirisches Material fuer die Diskussion zur Frage",
  "'ersetzen oder ergaenzen'."
)

# --- Hybrid-Ergaenzungslogik ------------------------------------------------
# Der Hybrid ist kein weiterer nichtlinearer Optimierer. Er kombiniert das
# klassische MaxSharpe-Portfolio und das aktive Behavioral-Portfolio linear.
# Damit bleibt die Forschungsfrage "ergaenzen" transparent interpretierbar.
HYBRID_ALPHA <- 0.50                     # Anteil MaxSharpe im Hybrid-Blend

cat("[Setup] Behavioral: aktive FOMO/FOL-Ratio max Skew(w)/SemiDev_0(w)\n")
cat("[Setup] Solver:", FOMO_FOL_SOLVER,
    "- DEoptim verfuegbar:", has_DEoptim, "\n")
if (tolower(FOMO_FOL_SOLVER) == "auto") {
  cat("[Setup] Auto-Schwelle: DEoptim bis n <=", FOMO_FOL_DE_N_MAX,
      ", darueber Local Search\n")
} else if (tolower(FOMO_FOL_SOLVER) == "deoptim") {
  cat("[Setup] DEoptim wird fuer Behavioral unabhaengig von der Universumsgroesse erzwungen\n")
}
cat(sprintf("[Setup] Arnott-alpha: %.2f * SemiDev_%dT / SemiDev_full, begrenzt auf [%.2f, %.2f]\n",
            ARNOTT_ALPHA_BASE, ARNOTT_RECENT_WINDOW_DAYS,
            ARNOTT_ALPHA_MIN, ARNOTT_ALPHA_MAX))
cat(sprintf("[Setup] Hybrid: %.0f%% MaxSharpe + %.0f%% Behavioral-FOMO/FOL\n",
            100 * HYBRID_ALPHA, 100 * (1 - HYBRID_ALPHA)))
cat(sprintf("[Setup] Gemeinsame Constraints: long-only, voll investiert, max %.1f%% je Titel (MIN_WEIGHT %.6f nur numerisch)\n",
            100 * MAX_WEIGHT, MIN_WEIGHT))
ES_ALPHA <- 0.95                         # Konfidenz fuer ES (wird fuer Metrik CVaR wiederverwendet)

# --- Modell-Schalter ---
# RUN_HYBRID = TRUE schaltet das Hybrid-Modell zu. Der Hybrid ist ein
# transparenter Blend aus MaxSharpe und aktivem Behavioral.
RUN_HYBRID <- TRUE

# --- Krisen-/Boomphasen-Definitionen ---
CRISIS_START  <- as.Date("2020-02-01")   # Corona-Stressphase: breiteres Fenster fuer robuste Tail-Metriken
CRISIS_END    <- as.Date("2020-06-30")   # inkl. Crash und erste Erholungsphase
BOOM_START    <- as.Date("2023-01-01")   # Tech-Rallye 2023 Start
BOOM_END      <- as.Date("2023-12-31")   # Tech-Rallye 2023 Ende


# ==============================================================================
# 2. DATEN-IMPORT
# ==============================================================================

## ---- 2.1 Haupt-Zeitreihe (Total-Return-Index) -------------------------------
# Aufbau des Sheets "S&P500":
#   Zeile 1:  Formel + Unternehmensnamen
#   Zeile 2:  "RI"  + Letzter RI-Wert (statisch)
#   Zeile 3:  "MV"  + Letzter MV-Wert (statisch)
#   Zeile 4:  "UP"  + Letzter UP-Wert (statisch)
#   Zeile 5:  "VO"  + Letztes Volumen (statisch)
#   Zeile 6+: Datum + tägliche RI-Zeitreihe
#
# Vorgehen: zweiphasiger Import, um Header (Zeile 1) und Daten (ab Zeile 6) sauber zu trennen.

cat("[1/7] Lade Kursdaten aus", DATA_PATH, "...\n")

# Phase 1: Unternehmensnamen aus Zeile 1
header_row <- read_excel(DATA_PATH, sheet = "S&P500",
                         n_max = 1, col_names = FALSE, .name_repair = "minimal")
ticker_names <- as.character(header_row[1, -1])  # ohne erste Spalte (Formel-Text)
ticker_names[is.na(ticker_names)] <- paste0("UNKNOWN_", which(is.na(ticker_names)))

# Phase 2: Preisdaten ab Zeile 6 (4 Metazeilen + Headerzeile = 5 Zeilen überspringen)
prices_raw <- read_excel(DATA_PATH, sheet = "S&P500",
                         skip = 5, col_names = FALSE, .name_repair = "minimal")
names(prices_raw) <- c("Date", ticker_names)

# Datum als Date konvertieren
prices_raw$Date <- as.Date(prices_raw$Date)

# Doppelte Datumszeilen entfernen (Feiertage stehen in Datastream z.T. doppelt drin)
prices_raw <- prices_raw[!duplicated(prices_raw$Date), ]

# In xts konvertieren
RI <- xts(x = as.matrix(prices_raw[, -1]),
          order.by = prices_raw$Date)

# Numerisch erzwingen (readxl liefert gelegentlich Text-Spalten bei #ERROR-Zellen)
storage.mode(RI) <- "numeric"

cat("    -> RI-Matrix:", nrow(RI), "Tage x", ncol(RI), "Aktien\n")
cat("    -> Zeitraum:", as.character(start(RI)), "bis", as.character(end(RI)), "\n")


## ---- 2.2 Historische Konstituenten-Listen -----------------------------------
# Ein Sheet je Jahr: "Rebalancing YYYY". Struktur analog zum Hauptsheet,
# jedoch nur EINE Datenzeile (= Jahresendpreis).
# Wir nutzen fehlende Preise (NA) als Hinweis auf Nicht-Zugehoerigkeit;
# Da in dieser Datei alle 503 Ticker in allen Rebal-Sheets vorhanden sind
# (siehe Methodik-Hinweis in der Dokumentation), fungiert der tatsaechliche
# Filter ueber den RI-Historien-Check (siehe Funktion get_universe() unten).

cat("[2/7] Lade historische Konstituenten-Listen ...\n")

constituent_list <- list()
for (yr in 2016:2026) {
  sheet_name <- paste0("Rebalancing ", yr)
  
  hd <- read_excel(DATA_PATH, sheet = sheet_name, n_max = 1,
                   col_names = FALSE, .name_repair = "minimal")
  ticker_yr <- as.character(hd[1, -1])
  
  # Jahresendpreise aus Zeile 4 lesen
  row4 <- read_excel(DATA_PATH, sheet = sheet_name, skip = 3, n_max = 1,
                     col_names = FALSE, .name_repair = "minimal")
  prices_yr <- suppressWarnings(as.numeric(row4[1, -1]))
  
  # Nur Ticker mit gueltigem Preis behalten
  valid <- !is.na(prices_yr) & !is.na(ticker_yr)
  constituent_list[[as.character(yr)]] <- ticker_yr[valid]
}
cat("    -> Konstituenten pro Jahr (Anzahl):\n")
print(sapply(constituent_list, length))


# ==============================================================================
# 3. RENDITEN & UNIVERSUMSFUNKTION
# ==============================================================================

## ---- 3.1 Berechnung diskreter Tagesrenditen ---------------------------------
# R_t = RI_t / RI_{t-1} - 1
# Wir nutzen PerformanceAnalytics::Return.calculate als kanonische, robuste
# Implementierung (behandelt NAs automatisch korrekt).

cat("[3/7] Berechne diskrete Tagesrenditen ...\n")

returns <- Return.calculate(RI, method = "discrete")
returns <- returns[-1, ]  # erste Zeile ist komplett NA

# Extreme Ausreisser abschneiden (ggf. Fehler in Datastream-Daten)
# Wir begrenzen einzelne Tagesrenditen auf [-0.5, 2.0]; Werte ausserhalb
# werden hier als Datastream-Ausreisser bzw. Datenfehler behandelt.
returns_clean <- returns
returns_clean[returns_clean < -0.5] <- NA
returns_clean[returns_clean >  2.0] <- NA

cat("    -> Renditen-Matrix:", nrow(returns_clean), "Tage x", ncol(returns_clean), "Aktien\n")


## ---- 3.2 Funktion: Universum zum Rebalancing-Zeitpunkt ----------------------
# Kombination aus zwei Filtern:
#   (a) Ticker muss im Rebalancing-Sheet des betreffenden Jahres mit
#       gueltigem Preis enthalten sein.
#   (b) Ticker muss mindestens MIN_HISTORY_DAYS Handelstage mit Daten vor
#       dem Rebalancing-Datum haben und im juengsten Historienfenster hoechstens
#       MAX_NA_SHARE Lueckentage aufweisen.
#
# SURVIVORSHIP BIAS:
# In dieser Skriptvariante wird bewusst die Datei S&P_500_Daten.xlsx verwendet.
# Der Survivorship Bias wird damit NICHT methodisch korrigiert. Die Ergebnisse
# sind daher als nicht-bias-korrigierter Modellvergleich zu interpretieren.
# Fuer eine bias-korrigierte Variante muss DATA_PATH wieder auf
# S&P_500_Daten_BIAS_KORRIGIERT.xlsx gesetzt und die Methodik entsprechend
# dokumentiert werden.

get_universe <- function(rebal_date, returns_mat, constituents_year) {
  # Trainingsfenster = alles vor und inkl. rebal_date
  R_train <- returns_mat[index(returns_mat) <= rebal_date, , drop = FALSE]
  
  # (a) Rebalancing-Sheet-Filter: nur Aktien mit gueltigem Preis im
  # betreffenden Jahressheet.
  in_index <- colnames(R_train) %in% constituents_year
  
  # (b) Historienfilter: mindestens MIN_HISTORY_DAYS gueltige Beobachtungen
  n_valid  <- colSums(!is.na(R_train))
  enough_history <- n_valid >= MIN_HISTORY_DAYS
  
  # (c) NA-Anteil im juengsten Fenster (letzte MIN_HISTORY_DAYS Tage) unter Schwellwert
  recent_window <- tail(R_train, MIN_HISTORY_DAYS)
  na_share_recent <- colMeans(is.na(recent_window))
  not_too_many_na <- na_share_recent <= MAX_NA_SHARE
  
  keep <- in_index & enough_history & not_too_many_na
  colnames(R_train)[keep]
}


# ==============================================================================
# 4. OPTIMIERUNGS-FUNKTIONEN
# ==============================================================================

## ---- 4.1 Hilfsfunktion: bereinigtes Renditepanel fuer Optimierung -----------
# WICHTIG: Wir filtern Ticker mit NAs im aktuellen Fenster AUS - nicht Zeilen.
# Hintergrund: na.omit() auf einer (T x N)-Matrix wuerde jeden Tag verwerfen,
# an dem auch nur ein einziger Ticker keinen Wert hat (z. B. Meta vor IPO
# 2012). Bei mehreren hundert Assets waere das ein Datenkiller - es bleibt ein
# rank-defizientes Panel, das den MaxSharpe-QP in NaNs treibt und numerische
# Optimierer mit Typ-Inkonsistenzen abwuergen kann. Loesung: Assets mit
# Luecken raus, verbleibende
# haben volle Historie. Das Universum schrumpft, die Kovarianz ist aber sauber.
#
# Zusaetzlich: Ticker mit Null-Varianz (konstante Zeitreihen, z. B. durch
# Handels-Aussetzung, die als konstanter RI durchging) erzeugen eine singulaere
# Kovarianzmatrix und bringen quadprog zum Absturz mit dem Fehler
# "Fehlender Wert, wo TRUE/FALSE noetig ist", weil Solver intern
# Nicht-Finitheit in der Kovarianzstruktur nicht robust behandeln.
prep_returns <- function(R_xts, tickers) {
  R_sub <- R_xts[, tickers, drop = FALSE]
  # (1) Ticker mit vollstaendiger Historie
  complete_cols <- colSums(is.na(R_sub)) == 0
  R_sub <- R_sub[, complete_cols, drop = FALSE]
  if (ncol(R_sub) == 0) {
    return(list(R = matrix(numeric(0), 0, 0), tickers = character(0)))
  }
  # (2) Ticker mit Mindest-Varianz (numerisch stabil)
  col_sd <- apply(R_sub, 2, sd, na.rm = TRUE)
  nonzero_var <- is.finite(col_sd) & col_sd > 1e-8
  R_sub <- R_sub[, nonzero_var, drop = FALSE]
  list(
    R       = as.matrix(R_sub),
    tickers = colnames(R_sub)
  )
}

# Hilfsfunktion: Gewichte, die vom Optimizer auf einem reduzierten Universum
# zurueckkommen, auf das gewuenschte Originaluniversum zurueckmappen (0-Padding).
expand_weights <- function(w_reduced, full_tickers) {
  w_full <- setNames(rep(0, length(full_tickers)), full_tickers)
  common <- intersect(names(w_reduced), full_tickers)
  w_full[common] <- as.numeric(w_reduced[common])
  w_full
}

`%||%` <- function(x, y) {
  if (is.null(x) || length(x) == 0) y else x
}

constraint_diagnostic_row <- function(w, model, yr, universe_size) {
  if (is.null(w)) {
    return(data.frame(
      Year = yr, Model = model, Success = FALSE,
      Solver = NA_character_,
      UniverseSize = universe_size, WeightSum = NA_real_,
      MinWeight = NA_real_, MaxWeight = NA_real_, ActivePositions = NA_integer_,
      FullyInvested = FALSE, LongOnly = FALSE, AllActive = FALSE,
      MaxWeightOK = FALSE,
      InSampleSkew = NA_real_, InSampleSemiDev = NA_real_,
      InSampleFomoFolRatio = NA_real_,
      stringsAsFactors = FALSE
    ))
  }
  ww <- as.numeric(w)
  ww[!is.finite(ww)] <- 0
  solver <- attr(w, "solver", exact = TRUE)
  if (is.null(solver) || !nzchar(solver)) {
    solver <- switch(model,
                     "MaxSharpe" = "quadprog",
                     "Arnott" = "deterministischer Arnott-Blend",
                     "Hybrid" = "linearer Blend",
                     NA_character_)
  }
  data.frame(
    Year = yr,
    Model = model,
    Success = TRUE,
    Solver = solver,
    UniverseSize = universe_size,
    WeightSum = sum(ww),
    MinWeight = min(ww),
    MaxWeight = max(ww),
    ActivePositions = sum(ww > 1e-8),
    FullyInvested = abs(sum(ww) - 1) <= 1e-5,
    LongOnly = all(ww >= -1e-8),
    AllActive = all(ww > 1e-8),
    MaxWeightOK = all(ww <= MAX_WEIGHT + 1e-5),
    InSampleSkew = as.numeric(attr(w, "insample_skew", exact = TRUE) %||% NA_real_),
    InSampleSemiDev = as.numeric(attr(w, "insample_semidev", exact = TRUE) %||% NA_real_),
    InSampleFomoFolRatio = as.numeric(attr(w, "insample_fomo_fol_ratio", exact = TRUE) %||% NA_real_),
    stringsAsFactors = FALSE
  )
}

portfolio_behavioral_stats <- function(r) {
  r <- as.numeric(r)
  r <- r[is.finite(r)]
  n <- length(r)
  if (n < 3) {
    return(c(excess_return_ann = NA_real_,
             semidev_ann = NA_real_,
             skew = NA_real_))
  }
  m <- mean(r)
  excess_return_ann <- m * 252 - RF_ANNUAL
  semidev_ann <- sqrt(mean(pmin(r, 0)^2) * 252)
  s2 <- sum((r - m)^2) / n
  if (!is.finite(s2) || s2 <= 1e-20) {
    skew <- NA_real_
  } else {
    s3 <- sum((r - m)^3) / n
    skew <- (sqrt(n * (n - 1)) / (n - 2)) * (s3 / s2^1.5)
  }
  c(excess_return_ann = excess_return_ann,
    semidev_ann = semidev_ann,
    skew = skew)
}

project_box_weights <- function(w, max_weight = MAX_WEIGHT) {
  w <- as.numeric(w)
  w[!is.finite(w)] <- 0
  w[w < 0] <- 0
  if (sum(w) <= 0) {
    w <- rep(1 / length(w), length(w))
  } else {
    w <- w / sum(w)
  }
  
  # Projektion auf long-only, sum=1, max_weight. Das ist bewusst einfach
  # gehalten und wird in der lokalen Suche sehr oft aufgerufen.
  for (iter in 1:100) {
    over <- w > max_weight
    if (!any(over)) break
    excess <- sum(w[over] - max_weight)
    w[over] <- max_weight
    under <- !over & w > 0
    if (!any(under) || sum(w[under]) <= 1e-12) break
    w[under] <- w[under] + excess * w[under] / sum(w[under])
  }
  
  w[w < 1e-12] <- 0
  if (sum(w) <= 0) {
    w <- rep(1 / length(w), length(w))
  } else {
    w <- w / sum(w)
  }
  w
}

optimize_fomo_fol_ratio <- function(R_train_xts, tickers,
                                    max_weight = MAX_WEIGHT) {
  pr           <- prep_returns(R_train_xts, tickers)
  R_mat        <- pr$R
  tickers_used <- pr$tickers
  n <- length(tickers_used)
  
  if (n * max_weight < 1) {
    stop("zu wenige Ticker fuer max_weight-Constraint")
  }
  
  project_weights <- function(w) {
    project_box_weights(w, max_weight = max_weight)
  }
  
  # Zielfunktion: DEoptim und Local Search minimieren. Daher wird
  # -Skew/SemiDev verwendet, was dem Maximieren des FOMO/FOL-Verhaeltnisses
  # entspricht. Bei negativer Schiefe wird eine stetige Penalty genutzt, damit
  # die Vorzeichenfalle vermieden wird und trotzdem nach niedrigem FOL-Risiko
  # gesucht werden kann.
  objective <- function(x) {
    w <- project_weights(x)
    active <- which(w > 1e-12)
    if (length(active) == 0) return(1e6)
    
    r_p <- as.numeric(R_mat[, active, drop = FALSE] %*% w[active])
    r_p <- r_p[is.finite(r_p)]
    n_obs <- length(r_p)
    if (n_obs < 3) return(1e6)
    
    semidev <- sqrt(mean(pmin(r_p, 0)^2) * 252)
    if (!is.finite(semidev) || semidev < 1e-8) return(1e6)
    
    m  <- mean(r_p)
    s2 <- sum((r_p - m)^2) / n_obs
    if (!is.finite(s2) || s2 <= 1e-20) return(1e6)
    s3 <- sum((r_p - m)^3) / n_obs
    skew <- (sqrt(n_obs * (n_obs - 1)) / (n_obs - 2)) * (s3 / s2^1.5)
    if (!is.finite(skew)) return(1e6)
    
    if (skew <= 0) return(100 - 100 * skew + 10 * semidev)
    -skew / semidev
  }
  
  eval_candidate <- function(w, best) {
    val <- objective(w)
    if (is.finite(val) && val < best$value) {
      best$value <- val
      best$w <- project_weights(w)
    }
    best
  }
  
  local_search <- function() {
    active_min <- max(FOMO_FOL_ACTIVE_MIN, ceiling(1 / max_weight))
    active_max <- min(FOMO_FOL_ACTIVE_MAX, n)
    if (active_min > active_max) active_max <- active_min
    
    asset_stats <- t(apply(R_mat, 2, portfolio_behavioral_stats))
    asset_ratio <- asset_stats[, "skew"] /
      pmax(asset_stats[, "semidev_ann"], 1e-8)
    asset_ratio[!is.finite(asset_ratio)] <- -Inf
    if (all(!is.finite(asset_ratio))) {
      asset_ratio <- -asset_stats[, "semidev_ann"]
      asset_ratio[!is.finite(asset_ratio)] <- -Inf
    }
    ranked <- order(asset_ratio, decreasing = TRUE)
    ranked <- ranked[is.finite(asset_ratio[ranked])]
    if (length(ranked) < active_min) ranked <- seq_len(n)
    pool_n <- min(n, max(active_min, FOMO_FOL_LOCAL_TOPPOOL))
    pool <- ranked[seq_len(min(pool_n, length(ranked)))]
    
    best <- list(value = Inf, w = rep(1 / n, n))
    best <- eval_candidate(rep(1 / n, n), best)
    
    # FOL-Anker und einige deterministische Top-k-Startpunkte geben der Suche
    # sinnvolle Startwerte, bevor zufaellige Sparse-Kandidaten erzeugt werden.
    w_fol <- tryCatch(optimize_min_semicov(R_train_xts, tickers),
                      error = function(e) NULL)
    if (!is.null(w_fol)) {
      best <- eval_candidate(w_fol[tickers_used], best)
    }
    
    for (k in unique(pmin(c(active_min, 15, 20, 30, active_max), n))) {
      idx <- ranked[seq_len(min(k, length(ranked)))]
      w0 <- rep(0, n)
      w0[idx] <- 1 / length(idx)
      best <- eval_candidate(w0, best)
    }
    
    for (trial in seq_len(FOMO_FOL_LOCAL_TRIALS)) {
      k <- sample(active_min:active_max, 1)
      idx <- sample(pool, size = min(k, length(pool)), replace = FALSE)
      w0 <- rep(0, n)
      w0[idx] <- rexp(length(idx), rate = 1)
      best <- eval_candidate(w0, best)
    }
    
    for (trial in seq_len(FOMO_FOL_LOCAL_JITTER)) {
      w0 <- best$w
      active <- which(w0 > 1e-8)
      if (length(active) == 0) active <- sample(pool, active_min)
      
      add_pool <- setdiff(pool, active)
      if (runif(1) < 0.35 && length(active) < active_max &&
          length(add_pool) > 0) {
        add <- sample(add_pool, 1)
        active <- c(active, add)
        w0[add] <- median(w0[w0 > 0], na.rm = TRUE)
      }
      if (runif(1) < 0.25 && length(active) > active_min) {
        drop <- sample(active, 1)
        w0[drop] <- 0
        active <- setdiff(active, drop)
      }
      w0[active] <- w0[active] * exp(rnorm(length(active), sd = 0.35))
      best <- eval_candidate(w0, best)
    }
    
    best$w
  }
  
  solver_mode <- tolower(FOMO_FOL_SOLVER)
  if (!solver_mode %in% c("auto", "deoptim", "local")) {
    stop("FOMO_FOL_SOLVER muss 'auto', 'DEoptim' oder 'local' sein.")
  }
  
  if (solver_mode == "deoptim" && !has_DEoptim) {
    stop("FOMO_FOL_SOLVER verlangt DEoptim, aber das Paket ist in dieser ",
         "R-Installation nicht verfuegbar. Installiere es mit ",
         "install.packages('DEoptim') in genau dieser R-Version oder setze ",
         "FOMO_FOL_SOLVER <- 'local'. Aktuelle .libPaths(): ",
         paste(.libPaths(), collapse = " | "))
  }
  
  use_deoptim <- has_DEoptim &&
    (solver_mode == "deoptim" ||
       (solver_mode == "auto" && n <= FOMO_FOL_DE_N_MAX))
  
  solver <- "Local Search"
  if (use_deoptim) {
    solver <- "DEoptim"
    ctrl <- DEoptim::DEoptim.control(
      NP = max(10 * n, 50),
      itermax = FOMO_FOL_DE_ITERMAX,
      trace = FALSE
    )
    res <- DEoptim::DEoptim(
      fn = objective,
      lower = rep(0, n),
      upper = rep(1, n),
      control = ctrl
    )
    w <- project_weights(res$optim$bestmem)
  } else {
    if (solver_mode == "auto" && !has_DEoptim && n <= FOMO_FOL_DE_N_MAX) {
      message("      Hinweis: DEoptim ist in dieser R-Installation nicht ",
              "verfuegbar; nutze Local Search mit gleicher Zielfunktion.")
    }
    w <- local_search()
  }
  
  names(w) <- tickers_used
  r_p <- as.numeric(R_mat %*% w)
  st  <- portfolio_behavioral_stats(r_p)
  ratio <- st["skew"] / max(st["semidev_ann"], 1e-8)
  cat(sprintf(
    "          FOMO/FOL (%s): Skew=%.3f, SemiDev=%.4f, Ratio=%.3f, Positionen=%d\n",
    solver, st["skew"], st["semidev_ann"], ratio, sum(w > 1e-6)
  ))
  
  w_full <- expand_weights(w, tickers)
  attr(w_full, "solver") <- solver
  attr(w_full, "insample_skew") <- as.numeric(st["skew"])
  attr(w_full, "insample_semidev") <- as.numeric(st["semidev_ann"])
  attr(w_full, "insample_fomo_fol_ratio") <- as.numeric(ratio)
  if (any(w_full < -1e-8) ||
      any(w_full > max_weight + 1e-5) ||
      abs(sum(w_full) - 1) > 1e-5) {
    stop("FOMO/FOL-Loesung verletzt gemeinsame Portfolio-Constraints")
  }
  
  w_full
}

optimize_behavioral_arnott <- function(R_train_xts, tickers) {
  pr           <- prep_returns(R_train_xts, tickers)
  R_mat        <- pr$R
  tickers_used <- pr$tickers
  if (length(tickers_used) < MIN_ASSETS_FOR_MAX_WEIGHT) {
    stop("zu wenige Ticker fuer Arnott-Robustheitsportfolio")
  }
  
  # FOL-Schicht: defensiver Downside-Anker.
  w_fol <- optimize_min_semicov(R_train_xts, tickers)
  
  # FOMO-Schicht: breite Marktpartizipation als passives
  # Schiefe-Harvesting im Robustheitsmodell.
  w_market <- setNames(rep(0, length(tickers)), tickers)
  w_market[tickers_used] <- 1 / length(tickers_used)
  
  R_market <- as.numeric(R_mat %*% w_market[tickers_used])
  recent_n <- min(ARNOTT_RECENT_WINDOW_DAYS, length(R_market))
  R_recent <- tail(R_market, recent_n)
  
  semidev_recent <- sqrt(mean(pmin(R_recent, 0)^2) * 252)
  semidev_full   <- sqrt(mean(pmin(R_market, 0)^2) * 252)
  
  alpha_t <- ARNOTT_ALPHA_BASE *
    (semidev_recent / max(semidev_full, 1e-8))
  alpha_t <- min(ARNOTT_ALPHA_MAX, max(ARNOTT_ALPHA_MIN, alpha_t))
  
  w <- alpha_t * w_fol + (1 - alpha_t) * w_market
  names(w) <- tickers
  w <- project_box_weights(w, MAX_WEIGHT)
  names(w) <- tickers
  attr(w, "solver") <- "deterministischer Arnott-Blend"
  
  r_p <- as.numeric(R_mat %*% w[tickers_used])
  st <- portfolio_behavioral_stats(r_p)
  attr(w, "insample_skew") <- as.numeric(st["skew"])
  attr(w, "insample_semidev") <- as.numeric(st["semidev_ann"])
  attr(w, "insample_fomo_fol_ratio") <-
    as.numeric(st["skew"] / max(st["semidev_ann"], 1e-8))
  cat(sprintf(
    "          Arnott-Robustheit: alpha=%.2f (FOL=%.0f%%, FOMO=%.0f%%), SemiDev_recent=%.4f, SemiDev_full=%.4f, Skew=%.3f, Excess=%.4f, Positionen=%d\n",
    alpha_t, 100 * alpha_t, 100 * (1 - alpha_t),
    semidev_recent, semidev_full, st["skew"], st["excess_return_ann"],
    sum(w > 1e-8)
  ))
  
  if (any(w < -1e-8) ||
      any(w > MAX_WEIGHT + 1e-5) ||
      abs(sum(w) - 1) > 1e-5) {
    stop("Arnott-Robustheitsloesung verletzt gemeinsame Portfolio-Constraints")
  }
  
  w
}

# FOL-Ankerportfolio fuer Arnott und als Startwert der Behavioral-Suche.
# Exakte Portfolio-Target-Semivarianz mit pmin(R %*% w, 0)^2 ist ein
# quadratisches Programm mit vielen Hilfsvariablen. Fuer ein robustes und
# schnelles Bachelorarbeits-Setup nutzen wir die uebliche Downside-
# Semicovariance-Naeherung: positive Asset-Renditen werden auf 0 gesetzt,
# daraus wird eine Downside-Momentmatrix gebildet und per quadprog minimiert.
optimize_min_semicov <- function(R_train_xts, tickers) {
  pr           <- prep_returns(R_train_xts, tickers)
  R_mat        <- pr$R
  tickers_used <- pr$tickers
  if (length(tickers_used) < MIN_ASSETS_FOR_MAX_WEIGHT) {
    stop("zu wenige Ticker fuer Min-Semicovariance-Portfolio")
  }
  
  n <- length(tickers_used)
  R_down <- pmin(R_mat, 0)
  Sigma_down <- crossprod(R_down) / nrow(R_down)
  eps <- 1e-6 * mean(diag(Sigma_down))
  if (!is.finite(eps) || eps <= 0) eps <- 1e-8
  Sigma_down <- Sigma_down + eps * diag(n)
  
  Dmat <- 2 * Sigma_down
  dvec <- rep(0, n)
  Amat <- cbind(rep(1, n), diag(n), -diag(n))
  bvec <- c(1, rep(MIN_WEIGHT, n), rep(-MAX_WEIGHT, n))
  
  sol <- tryCatch(
    quadprog::solve.QP(Dmat = Dmat, dvec = dvec,
                       Amat = Amat, bvec = bvec, meq = 1),
    error = function(e) stop("quadprog solve.QP (MinSemiCov) fehlgeschlagen: ",
                             e$message)
  )
  
  w <- sol$solution
  names(w) <- tickers_used
  w[w < 1e-10] <- 0
  if (sum(w) > 0) w <- w / sum(w)
  if (any(w < MIN_WEIGHT - 1e-8) || any(w > MAX_WEIGHT + 1e-5)) {
    stop("MinSemiCov-Loesung verletzt gemeinsame Max-Gewicht-Constraint")
  }
  expand_weights(w, tickers)
}


## ---- 4.2 Modell 1: Maximum Sharpe Ratio (Mean-Variance) via quadprog --------
# Max Sharpe ist ein Fractional QP und nicht direkt in solve.QP einsetzbar.
# Standard-Transformation (siehe Cornuejols/Tuetuencue 2007, Kapitel 8.3):
#
#   Substitution y = k*w mit k = 1 / (mu'w - r_f)  (k > 0 angenommen)
#   Dann loest das aequivalente QP:
#       min  y' Sigma y  s.t.  (mu - r_f)' y = 1,  y >= 0
#   Rueckrechnung:  w = y / sum(y)
#
# Voraussetzung: max(mu - r_f) > 0 im Universum, sonst existiert keine
# Tangency-Loesung. Der Code prueft diese Bedingung und bricht sonst sauber ab.
optimize_maxsharpe <- function(R_train_xts, tickers, rf = RF_DAILY) {
  pr           <- prep_returns(R_train_xts, tickers)
  R_mat        <- pr$R
  tickers_used <- pr$tickers
  if (length(tickers_used) < 2) stop("zu wenige Ticker mit kompletter Historie")
  
  n     <- length(tickers_used)
  if (n * MAX_WEIGHT < 1) {
    stop("zu wenige Ticker fuer gemeinsame Max-Gewicht-Constraint: ",
         n, " * ", MAX_WEIGHT, " < 1")
  }
  mu    <- colMeans(R_mat)
  Sigma <- cov(R_mat)
  eps   <- 1e-6 * mean(diag(Sigma))
  Sigma_reg <- Sigma + eps * diag(n)
  
  mu_excess <- mu - rf
  if (max(mu_excess) <= 0) {
    stop("kein Ticker mit positivem Excess-Return im Trainingsfenster - ",
         "Tangency-Loesung existiert nicht")
  }
  
  Dmat <- 2 * Sigma_reg
  dvec <- rep(0, n)
  # Erste Spalte = Gleichheitsconstraint (mu-rf)'y = 1,
  # danach w_i >= MIN_WEIGHT und die gemeinsame Box-Constraint w_i <= MAX_WEIGHT.
  #
  # Wegen w = y / sum(y) werden Unter- und Obergrenze linear als
  #   y_i >= MIN_WEIGHT * sum(y)
  #   y_i <= MAX_WEIGHT * sum(y)
  # formuliert, also:
  #   y_i - MIN_WEIGHT * sum(y) >= 0
  #   MAX_WEIGHT * sum(y) - y_i >= 0
  lower_Amat <- sapply(seq_len(n), function(j) {
    a <- rep(-MIN_WEIGHT, n)
    a[j] <- 1 - MIN_WEIGHT
    a
  })
  upper_Amat <- sapply(seq_len(n), function(j) {
    a <- rep(MAX_WEIGHT, n)
    a[j] <- MAX_WEIGHT - 1
    a
  })
  Amat <- cbind(mu_excess, lower_Amat, upper_Amat)
  bvec <- c(1, rep(0, n), rep(0, n))
  
  sol <- tryCatch(
    quadprog::solve.QP(Dmat = Dmat, dvec = dvec,
                       Amat = Amat, bvec = bvec, meq = 1),
    error = function(e) stop("quadprog solve.QP (MaxSharpe) fehlgeschlagen: ",
                             e$message)
  )
  
  y <- sol$solution
  if (sum(y) <= 1e-10) stop("Max-Sharpe-Loesung degeneriert (sum(y) <= 0)")
  
  w <- y / sum(y)
  names(w) <- tickers_used
  w[w < 1e-10] <- 0
  if (sum(w) > 0) w <- w / sum(w)
  if (any(w < MIN_WEIGHT - 1e-8) || any(w > MAX_WEIGHT + 1e-5)) {
    stop("MaxSharpe-Loesung verletzt gemeinsame Max-Gewicht-Constraint")
  }
  expand_weights(w, tickers)
}


## ---- 4.3 Gemeinsame Kennzahlen- und Blend-Helfer ----------------------------

# Custom Risk-Funktion: Target-Semivarianz
# Berechnet fuer eine Renditereihe den Mittelwert der quadrierten echten
# Verluste unter 0. Das passt zur Interpretation von Fear of Loss: relevant
# ist nicht Unterperformance gegenueber dem eigenen Mittelwert, sondern ein
# tatsaechlich negativer Tagesertrag.
SemiVariance <- function(R, ...) {
  r <- tryCatch(as.numeric(R), error = function(e) NA_real_)
  r <- r[is.finite(r)]
  if (length(r) < 2) return(0)
  out <- mean(pmin(r, 0)^2)
  as.numeric(out[1])
}

make_hybrid_blend <- function(w_sharpe, w_behav, tickers,
                              alpha = HYBRID_ALPHA) {
  if (is.null(w_sharpe) || is.null(w_behav)) {
    stop("Hybrid-Blend benoetigt MaxSharpe- und Behavioral-Gewichte")
  }
  ws <- expand_weights(w_sharpe, tickers)
  wb <- expand_weights(w_behav,  tickers)
  w <- alpha * ws + (1 - alpha) * wb
  names(w) <- tickers
  w <- project_box_weights(w, MAX_WEIGHT)
  names(w) <- tickers
  attr(w, "solver") <- "linearer Blend"
  
  if (any(w < -1e-8) ||
      any(w > MAX_WEIGHT + 1e-5) ||
      abs(sum(w) - 1) > 1e-5) {
    stop("Hybrid-Blend verletzt gemeinsame Portfolio-Constraints")
  }
  
  w
}


# ==============================================================================
# 5. BACKTEST-SCHLEIFE (EXPANDING WINDOW)
# ==============================================================================

cat("[4/7] Starte Backtest-Schleife ueber", length(REBAL_YEARS), "Rebalancing-Zeitpunkte ...\n")

# Bestimme reale (letzte Handelstage) Rebalancing-Termine aus den Daten
rebal_dates <- sapply(REBAL_YEARS, function(y) {
  days_in_y <- index(returns_clean)[format(index(returns_clean), "%Y") == as.character(y)]
  as.character(max(days_in_y))
})
rebal_dates <- as.Date(rebal_dates)
names(rebal_dates) <- as.character(REBAL_YEARS)

# Speicher fuer Ergebnisse
weights_sharpe   <- list()
weights_behav    <- list()
weights_arnott   <- list()
weights_hybrid   <- list()
oos_ret_sharpe   <- list()
oos_ret_behav    <- list()
oos_ret_arnott   <- list()
oos_ret_hybrid   <- list()
constraint_diag  <- list()

for (i in seq_along(rebal_dates)) {
  yr         <- names(rebal_dates)[i]
  rd         <- rebal_dates[i]
  
  # Ende der Haltedauer = Rebalancing des Folgejahres (letzter Tag vor neuem Rebal)
  if (i < length(rebal_dates)) {
    holding_end <- rebal_dates[i + 1]
  } else {
    holding_end <- OOS_END
  }
  
  cat(sprintf("  [%d/%d] Rebal %s -> OOS bis %s\n",
              i, length(rebal_dates), rd, holding_end))
  
  # Universum zum Rebalancing-Zeitpunkt
  uni <- get_universe(
    rebal_date        = rd,
    returns_mat       = returns_clean,
    constituents_year = constituent_list[[yr]]
  )
  # Zusaetzlich: pre-check, wie viele davon tatsaechlich volle Historie haben
  pr_peek <- prep_returns(returns_clean[index(returns_clean) <= rd, ], uni)
  cat(sprintf("      Universum: %d Titel (davon mit voller Historie + Varianz > 0: %d)\n",
              length(uni), length(pr_peek$tickers)))
  
  # Trainingsdaten (expanding window)
  R_train <- returns_clean[index(returns_clean) <= rd, uni]
  
  # --- Modell 1: Mean-Variance (MaxSharpe / Tangency) ---
  cat("      optimiere MaxSharpe ...\n")
  w_sharpe <- tryCatch(
    optimize_maxsharpe(R_train, uni),
    error = function(e) {
      message(sprintf("      !! MaxSharpe %s FEHLGESCHLAGEN: %s", yr, e$message))
      NULL
    }
  )
  
  # --- Modell 2: Behavioral (aktive FOMO/FOL-Ratio) -------------------------
  # FOMO = korrigierte Schiefe, FOL = Semideviation unter 0.
  # Optimiert wird Skew(w) / SemiDev_0(w) unter denselben Constraints.
  cat("      optimiere Behavioral (aktive FOMO/FOL-Ratio) ...\n")
  t0 <- Sys.time()
  w_behav <- tryCatch(
    optimize_fomo_fol_ratio(R_train, uni),
    error = function(e) {
      message(sprintf("      !! Behavioral %s FEHLGESCHLAGEN: %s", yr, e$message))
      NULL
    }
  )
  if (!is.null(w_behav)) {
    cat(sprintf(
      "      Behavioral-Solver: %s | In-Sample Skew=%.3f, SemiDev=%.4f, Ratio=%.3f\n",
      attr(w_behav, "solver") %||% "unbekannt",
      as.numeric(attr(w_behav, "insample_skew") %||% NA_real_),
      as.numeric(attr(w_behav, "insample_semidev") %||% NA_real_),
      as.numeric(attr(w_behav, "insample_fomo_fol_ratio") %||% NA_real_)
    ))
  }
  cat(sprintf("      ... fertig in %.1f sek\n",
              as.numeric(difftime(Sys.time(), t0, units = "secs"))))
  
  # --- Robustheitsmodell: Arnott-Blend --------------------------------------
  # Passives Schiefe-Harvesting ueber 1/N-Marktbein; dient dem methodischen
  # Vergleich zur aktiven FOMO/FOL-Optimierung.
  cat("      berechne Arnott-Robustheitsmodell (FOL-MinSemiCov + FOMO-1/N) ...\n")
  t0 <- Sys.time()
  w_arnott <- tryCatch(
    optimize_behavioral_arnott(R_train, uni),
    error = function(e) {
      message(sprintf("      !! Arnott %s FEHLGESCHLAGEN: %s", yr, e$message))
      NULL
    }
  )
  cat(sprintf("      ... fertig in %.1f sek\n",
              as.numeric(difftime(Sys.time(), t0, units = "secs"))))
  
  # --- Modell 3: Hybrid (Blend aus MaxSharpe + aktivem Behavioral) ----------
  # Der Hybrid testet die Forschungsfrage "ergaenzen" transparent als lineare
  # Kombination der klassischen und der verhaltensorientierten Allokation.
  if (RUN_HYBRID) {
    cat("      erstelle Hybrid-Blend (MaxSharpe + Behavioral-FOMO/FOL) ...\n")
    t0 <- Sys.time()
    w_hybrid <- tryCatch({
      if (!is.null(w_sharpe) && !is.null(w_behav)) {
        make_hybrid_blend(w_sharpe, w_behav, uni, alpha = HYBRID_ALPHA)
      } else {
        message(sprintf("      !! Hybrid %s: MaxSharpe oder Behavioral fehlt", yr))
        NULL
      }
    },
    error = function(e) {
      message(sprintf("      !! Hybrid %s FEHLGESCHLAGEN: %s", yr, e$message))
      NULL
    }
    )
    cat(sprintf("      ... fertig in %.1f sek\n",
                as.numeric(difftime(Sys.time(), t0, units = "secs"))))
  } else {
    w_hybrid <- NULL
  }
  
  # Mit einfacher [[<- NULL wuerde R das Listenelement loeschen. Die [<- Form
  # behaelt auch fehlgeschlagene Jahre als NULL-Eintrag fuer Diagnose/Export.
  weights_sharpe[yr] <- list(w_sharpe)
  weights_behav[yr]  <- list(w_behav)
  weights_arnott[yr] <- list(w_arnott)
  weights_hybrid[yr] <- list(w_hybrid)
  constraint_diag[[yr]] <- rbind(
    constraint_diagnostic_row(w_sharpe, "MaxSharpe",  yr, length(uni)),
    constraint_diagnostic_row(w_behav,  "Behavioral", yr, length(uni)),
    constraint_diagnostic_row(w_arnott, "Arnott",     yr, length(uni)),
    constraint_diagnostic_row(w_hybrid, "Hybrid",     yr, length(uni))
  )
  
  # --- OOS-Renditen (statische Zielgewichte; taeglich auf Zielgewichte rebalanciert) ---
  # Mathematisch entspricht as.matrix(R_oos) %*% w einer taeglich
  # rebalancierten Strategie mit konstanten Zielgewichten. Es ist keine
  # Strategie mit driftenden Einzeltitelgewichten.
  R_oos <- returns_clean[index(returns_clean) > rd & index(returns_clean) <= holding_end, uni]
  #
  # DELISTING-/NA-BEHANDLUNG:
  # Wenn ein Ticker waehrend der OOS-Periode delistet wird (Pleite, Uebernahme),
  # bricht die RI-Reihe ab und es kommen NAs. Die korrekte oekonomische
  # Behandlung haengt vom Grund ab, ist aber in Ermangelung von Detail-
  # informationen folgendermassen approximiert:
  #
  # 1) Letzte verfuegbare Rendite vor dem Abbruch -> Total-Loss-Tag (-100 %)
  #    annehmen, wenn der RI plausibel "abrupt" endet (Pleite-Annahme).
  # 2) Danach gibt der Ticker keine weitere Rendite (Position ist verfallen).
  #
  # Vereinfachte robuste Implementierung: am letzten gueltigen Tag jedes
  # delisteten Tickers wird eine Rendite von -1 (Total Loss) gesetzt.
  # WICHTIG: Da diese Skriptvariante S&P_500_Daten.xlsx nutzt, ist der
  # Survivorship Bias insgesamt NICHT korrigiert; diese Regel behandelt nur
  # NAs innerhalb einer OOS-Haltedauer, falls sie in der Datei auftreten.
  #
  # Hinweis: Falls dein LSEG-Datensatz fuer Uebernahmen/Mergers den letzten
  # RI-Wert oekonomisch korrekt enthaelt (Uebernahmepreis), ueberschaetzt
  # dieser Mechanismus den Verlust geringfuegig. Fuer eine reine Pleite-
  # Annahme ist er korrekt.
  for (col in colnames(R_oos)) {
    s <- as.numeric(R_oos[, col])
    last_valid <- max(which(!is.na(s)), -Inf)
    if (last_valid > 0 && last_valid < length(s)) {
      # Ticker delistet INNERHALB der Periode
      s[last_valid] <- -1.0    # Total-Loss am letzten Tag
      s[(last_valid + 1):length(s)] <- 0   # danach keine Position mehr
      R_oos[, col] <- s
    }
  }
  # NAs am Anfang (Ticker noch nicht existent) auf 0 setzen
  R_oos[is.na(R_oos)] <- 0
  
  if (!is.null(w_sharpe)) {
    oos_ret_sharpe[[yr]] <- xts(as.matrix(R_oos) %*% w_sharpe[uni],
                                order.by = index(R_oos))
  }
  if (!is.null(w_behav)) {
    oos_ret_behav[[yr]]  <- xts(as.matrix(R_oos) %*% w_behav[uni],
                                order.by = index(R_oos))
  }
  if (!is.null(w_arnott)) {
    oos_ret_arnott[[yr]] <- xts(as.matrix(R_oos) %*% w_arnott[uni],
                                order.by = index(R_oos))
  }
  if (!is.null(w_hybrid)) {
    oos_ret_hybrid[[yr]] <- xts(as.matrix(R_oos) %*% w_hybrid[uni],
                                order.by = index(R_oos))
  }
}

# Zusammenfuegen zu einer einzigen OOS-Zeitreihe je Modell
# ---- Sicherer Helper: gibt NULL zurueck statt Absturz, wenn Liste leer/komplett NULL ----
safe_rbind_xts <- function(lst, col_name) {
  lst <- lst[!vapply(lst, is.null, logical(1))]
  if (length(lst) == 0) {
    warning(sprintf(
      "Keine Ergebnisse fuer Strategie '%s' - Optimierung in allen Perioden fehlgeschlagen.",
      col_name))
    return(NULL)
  }
  out <- do.call(rbind, lst)
  colnames(out) <- col_name
  out
}

R_sharpe_oos <- safe_rbind_xts(oos_ret_sharpe, "MaxSharpe")
R_behav_oos  <- safe_rbind_xts(oos_ret_behav,  "Behavioral")
R_arnott_oos <- safe_rbind_xts(oos_ret_arnott, "Arnott")
R_hybrid_oos <- safe_rbind_xts(oos_ret_hybrid, "Hybrid")

cat("[5/7] Backtest abgeschlossen\n")

# ---- Diagnostik: zeige, welche Rebalancings erfolgreich waren -----------------
diag_df <- data.frame(
  Year       = names(rebal_dates),
  MaxSharpe  = vapply(names(rebal_dates), function(y) !is.null(weights_sharpe[[y]]), logical(1)),
  Behavioral = vapply(names(rebal_dates), function(y) !is.null(weights_behav[[y]]),  logical(1)),
  Arnott     = vapply(names(rebal_dates), function(y) !is.null(weights_arnott[[y]]), logical(1)),
  Hybrid     = vapply(names(rebal_dates), function(y) !is.null(weights_hybrid[[y]]), logical(1)),
  row.names  = NULL
)
constraint_diag_df <- if (length(constraint_diag) > 0) {
  do.call(rbind, constraint_diag)
} else {
  data.frame()
}
cat("\n--- Erfolgsstatus pro Rebalancing ---\n")
print(diag_df)
cat("\n--- Constraint-Check pro Modell ---\n")
print(constraint_diag_df)

if (!is.null(R_sharpe_oos)) cat("    MaxSharpe-OOS:",  nrow(R_sharpe_oos), "Tage\n")
if (!is.null(R_behav_oos))  cat("    Behavioral-OOS:", nrow(R_behav_oos),  "Tage\n")
if (!is.null(R_arnott_oos)) cat("    Arnott-OOS:",     nrow(R_arnott_oos), "Tage\n")
if (!is.null(R_hybrid_oos)) cat("    Hybrid-OOS:",     nrow(R_hybrid_oos), "Tage\n")

# ---- Fruehzeitiger Abbruch mit klarer Meldung, wenn keines der Modelle liefert ----
if (is.null(R_sharpe_oos) && is.null(R_behav_oos) &&
    is.null(R_arnott_oos) && is.null(R_hybrid_oos)) {
  stop("Keines der Modelle hat OOS-Renditen produziert. ",
       "Pruefe die Fehlermeldungen (warnings()) und den Diagnose-Output oben.")
}
if (is.null(R_behav_oos)) {
  stop("Behavioral hat keine OOS-Renditen produziert. ",
       "Die Excel wird nicht mit einem unvollstaendigen Modellvergleich ueberschrieben. ",
       "Pruefe die Behavioral-Fehlermeldungen direkt oberhalb im Konsolenoutput.")
}
if (is.null(R_arnott_oos)) {
  warning("Arnott-Robustheitsmodell hat keine OOS-Renditen produziert. ",
          "Der Hauptvergleich bleibt moeglich, der Robustheitsvergleich fehlt.")
}
if (RUN_HYBRID && is.null(R_hybrid_oos)) {
  stop("Hybrid hat keine OOS-Renditen produziert, weil MaxSharpe oder Behavioral fehlt. ",
       "Die Excel wird nicht mit einem unvollstaendigen Modellvergleich ueberschrieben.")
}


# ==============================================================================
# 6. BENCHMARK
# ==============================================================================

cat("[6/7] Lade Benchmark S&P 500 TR ...\n")

## ---- Echter S&P 500 TR via quantmod -----------------------------------------
R_sp500tr <- NULL
if (!has_quantmod) {
  stop("Paket 'quantmod' nicht installiert. Es wird zwingend fuer den ",
       "S&P 500 TR-Benchmark benoetigt. Bitte 'install.packages(\"quantmod\")'.")
}
suppressPackageStartupMessages(library(quantmod))
tryCatch({
  quantmod::getSymbols("^SP500TR", src = "yahoo",
                       from = OOS_START - 5, to = OOS_END, auto.assign = TRUE)
  sp_px <- Cl(SP500TR)
  R_sp500tr <- dailyReturn(sp_px, type = "arithmetic")
  R_sp500tr <- R_sp500tr[paste0(OOS_START, "/", OOS_END)]
  colnames(R_sp500tr) <- "SP500_TR"
  cat("    -> SP500TR geladen:", nrow(R_sp500tr), "Tage\n")
}, error = function(e) {
  stop("quantmod-Download von ^SP500TR fehlgeschlagen: ", e$message,
       "\nBitte Internetverbindung pruefen oder Daten manuell beilegen.")
})


# ==============================================================================
# 7. PERFORMANCE-METRIKEN
# ==============================================================================

cat("[7/7] Berechne Performance-Metriken & exportiere Excel ...\n")

## ---- 7.1 Gesamtes Panel aller OOS-Returns zusammenfuehren -------------------
# Nur nicht-NULL-Ergebnisse beruecksichtigen
panel_list <- list()
if (!is.null(R_sharpe_oos)) panel_list$MaxSharpe   <- R_sharpe_oos
if (!is.null(R_behav_oos))  panel_list$Behavioral  <- R_behav_oos
if (!is.null(R_arnott_oos)) panel_list$Arnott      <- R_arnott_oos
if (!is.null(R_hybrid_oos)) panel_list$Hybrid      <- R_hybrid_oos
panel_list$SP500_TR <- R_sp500tr

panel <- do.call(merge.xts, panel_list)
panel <- na.omit(panel)  # alle Serien auf gemeinsamen Zeitraum trimmen

## ---- 7.2 Metriken-Funktion --------------------------------------------------
# Annualisierte Rendite, Varianz, Target-Semivarianz, Schiefe, MaxDrawdown,
# CVaR, Sortino.
# Frequenzannahme: 252 Handelstage / Jahr.

calc_metrics <- function(R, scale = 252) {
  metr <- data.frame(row.names = colnames(R))
  for (col in colnames(R)) {
    r <- as.numeric(R[, col])
    r <- r[!is.na(r)]
    
    # Annualisierte geometrische Rendite aus der kumulierten Rendite
    ann_ret <- prod(1 + r)^(scale / length(r)) - 1
    # Annualisierte Varianz
    ann_var <- var(r) * scale
    # Annualisierte Target-Semivarianz (echte Verluste unter 0)
    ann_semivar <- SemiVariance(r) * scale
    # Schiefe
    skw <- PerformanceAnalytics::skewness(r)
    # Maximum Drawdown
    mdd <- as.numeric(maxDrawdown(R[, col]))
    # CVaR bei 95 % (PerformanceAnalytics-Default: historisch)
    cvar95 <- as.numeric(CVaR(R[, col], p = ES_ALPHA, method = "historical"))
    # Sortino-aehnlich: rf-adjustierte Rendite pro Semideviation unter 0
    sortino <- (ann_ret - RF_ANNUAL) / sqrt(ann_semivar)
    
    metr[col, "AnnReturn"]   <- ann_ret
    metr[col, "AnnVariance"] <- ann_var
    metr[col, "AnnSemiVar"]  <- ann_semivar
    metr[col, "Skewness"]    <- skw
    metr[col, "MaxDD"]       <- mdd
    metr[col, "CVaR_95"]     <- cvar95
    metr[col, "Sortino"]     <- sortino
  }
  metr
}

metrics_overall <- calc_metrics(panel)

## ---- 7.3 Subperioden-Analyse: Krise & Boom ----------------------------------
panel_crisis <- panel[paste0(CRISIS_START, "/", CRISIS_END)]
panel_boom   <- panel[paste0(BOOM_START,  "/", BOOM_END)]

metrics_crisis <- if (nrow(panel_crisis) > 5) calc_metrics(panel_crisis) else NULL
metrics_boom   <- if (nrow(panel_boom)   > 5) calc_metrics(panel_boom)   else NULL

cat("\n=== Metriken Gesamtperiode (OOS) ===\n")
print(round(metrics_overall, 4))
cat("\n=== Metriken Corona-Stressphase ===\n")
print(round(metrics_crisis, 4))
cat("\n=== Metriken Tech-Rallye 2023 ===\n")
print(round(metrics_boom, 4))


# ==============================================================================
# 8. EXCEL-EXPORT
# ==============================================================================

## ---- 8.1 Gewichte-Tabellen vorbereiten --------------------------------------
# Sparse-Format: wide-Tabelle mit Jahren als Zeilen, Tickern als Spalten.
# Viele Ticker kommen nur in wenigen Jahren vor, daher NA's zulassen.

weights_to_df <- function(weights_list) {
  all_tickers <- unique(unlist(lapply(weights_list, names)))
  df <- data.frame(Rebalancing = names(weights_list),
                   matrix(NA_real_, nrow = length(weights_list),
                          ncol = length(all_tickers),
                          dimnames = list(NULL, all_tickers)),
                   check.names = FALSE, stringsAsFactors = FALSE)
  for (i in seq_along(weights_list)) {
    w <- weights_list[[i]]
    if (is.null(w)) next
    df[i, names(w)] <- w
  }
  df
}

df_w_sharpe <- weights_to_df(weights_sharpe)
df_w_behav  <- weights_to_df(weights_behav)
df_w_arnott <- weights_to_df(weights_arnott)
df_w_hybrid <- weights_to_df(weights_hybrid)

## ---- 8.2 OOS-Renditen-Zeitreihen als Data Frame -----------------------------
df_returns <- data.frame(Date = index(panel), coredata(panel))

## ---- 8.3 Workbook schreiben -------------------------------------------------
wb <- createWorkbook()

# Metadaten-Sheet
addWorksheet(wb, "README")
readme_text <- data.frame(
  Eintrag = c(
    "Titel", "Erstellt am", "Zeitraum OOS", "Rebalancing",
    "Benchmark",
    "Modell 1", "Modell 2", "Robustheitsmodell", "Modell 3",
    "Gemeinsame Constraints", "Behavioral-Parameter",
    "Solver Behavioral",
    "Begruendung Behavioral",
    "Diskussion Schiefe",
    "Universum (jaehrlich)", "Rf (p.a.)",
    "Limitation rf",
    "Hinweis Survivorship"
  ),
  Wert = c(
    "Backtest Mean-Variance vs. Behavioral vs. Arnott vs. Hybrid (S&P 500)",
    as.character(Sys.time()),
    paste(OOS_START, "bis", OOS_END),
    paste0("Jaehrlich, ", paste(REBAL_YEARS, collapse = ", ")),
    "S&P 500 Total Return (Yahoo, ^SP500TR)",
    "Modell 1 (Mean-Variance / Max Sharpe, Tangency, quadprog)",
    "Modell 2 (Behavioral aktiv: max Skew(w) / SemiDev_0(w); FOMO/FOL-Ratio)",
    "Arnott-Blend (Robustheit: FOL-MinSemiCov + FOMO-1/N-Marktproxy mit zeitvariablem alpha)",
    paste0("Modell 3 (Hybrid-Blend: ", 100*HYBRID_ALPHA,
           "% MaxSharpe + ", 100*(1-HYBRID_ALPHA),
           "% Behavioral-FOMO/FOL)"),
    paste0("long-only, voll investiert, max. ", 100*MAX_WEIGHT,
           "% je Einzeltitel fuer alle optimierten Modelle; MIN_WEIGHT=",
           MIN_WEIGHT, " nur numerische Solver-Untergrenze"),
    paste0("Zielfunktion Behavioral: max Skew(w)/SemiDev_0(w); Schiefe = ",
           "Fisher-Pearson stichprobenkorrigiert; SemiDev_0 = annualisierte ",
           "Semideviation unter 0; negative Schiefe mit stetiger Penalty; ",
           "Hybrid-Alpha=", HYBRID_ALPHA, "; Arnott-alpha=", ARNOTT_ALPHA_BASE,
           " * SemiDev_", ARNOTT_RECENT_WINDOW_DAYS,
           "T / SemiDev_full, begrenzt auf [", ARNOTT_ALPHA_MIN,
           ", ", ARNOTT_ALPHA_MAX, "]"),
    paste0("FOMO_FOL_SOLVER=", FOMO_FOL_SOLVER,
           "; DEoptim verfuegbar=", has_DEoptim,
           if (tolower(FOMO_FOL_SOLVER) == "deoptim") {
             "; DEoptim wird fuer Behavioral erzwungen; FOMO_FOL_DE_N_MAX ist nur im auto-Modus relevant"
           } else {
             paste0("; DEoptim bis n<=", FOMO_FOL_DE_N_MAX,
                    " im auto-Modus, Local Search fuer groessere Universen")
           },
           "; Local Search bleibt als alternative Solveroption dokumentiert",
           "; Local Trials=", FOMO_FOL_LOCAL_TRIALS,
           ", Jitter=", FOMO_FOL_LOCAL_JITTER),
    BEHAVIORAL_METHOD_NOTE,
    SKEW_DISCUSSION_NOTE,
    paste0("gefiltert: >= ", MIN_HISTORY_DAYS, " Tage Historie, <= ",
           100*MAX_NA_SHARE, "% NAs"),
    paste0(100*RF_ANNUAL, "%"),
    paste0("Konstanter risikofreier Zinssatz von ", 100*RF_ANNUAL,
           "% p.a. ueber 2017-2026; bewusste Vereinfachung, in der Interpretation/Limitations zu nennen."),
    "Nicht korrigiert: Es wird S&P_500_Daten.xlsx verwendet; Survivorship Bias wird bewusst ausser Acht gelassen."
  ),
  stringsAsFactors = FALSE
)
writeData(wb, "README", readme_text)

# Returns OOS
addWorksheet(wb, "Returns_OOS")
writeData(wb, "Returns_OOS", df_returns)

# Metriken
addWorksheet(wb, "Metrics_Overall")
writeData(wb, "Metrics_Overall", data.frame(Strategy = rownames(metrics_overall),
                                            metrics_overall,
                                            row.names = NULL))

if (!is.null(metrics_crisis)) {
  addWorksheet(wb, "Metrics_Crisis_Corona")
  writeData(wb, "Metrics_Crisis_Corona",
            data.frame(Strategy = rownames(metrics_crisis),
                       metrics_crisis, row.names = NULL))
}
if (!is.null(metrics_boom)) {
  addWorksheet(wb, "Metrics_Boom_Tech23")
  writeData(wb, "Metrics_Boom_Tech23",
            data.frame(Strategy = rownames(metrics_boom),
                       metrics_boom, row.names = NULL))
}

# Diagnose: sichtbar machen, falls ein Modell in einzelnen Jahren nicht
# optimiert werden konnte. Das verhindert "stille" leere Sheets.
addWorksheet(wb, "Diagnostics")
writeData(wb, "Diagnostics", diag_df)
writeData(wb, "Diagnostics", "Constraint-Check: gleiche Anlagebedingungen fuer alle Modelle",
          startRow = nrow(diag_df) + 4, startCol = 1)
writeData(wb, "Diagnostics", constraint_diag_df,
          startRow = nrow(diag_df) + 5, startCol = 1)

# Gewichte
addWorksheet(wb, "Weights_MaxSharpe")
writeData(wb, "Weights_MaxSharpe", df_w_sharpe)

addWorksheet(wb, "Weights_Behavioral")
writeData(wb, "Weights_Behavioral", df_w_behav)

addWorksheet(wb, "Weights_Arnott")
writeData(wb, "Weights_Arnott", df_w_arnott)

addWorksheet(wb, "Weights_Hybrid")
writeData(wb, "Weights_Hybrid", df_w_hybrid)

# Workbook speichern
saveWorkbook(wb, OUTPUT_PATH, overwrite = TRUE)

cat("\nFERTIG. Ergebnisdatei:", normalizePath(OUTPUT_PATH, mustWork = FALSE), "\n")


################################################################################
#  ANMERKUNGEN FUER DIE BACHELORARBEIT                                         #
#  -----------------------------------------------------------------------     #
#  * Modell 2 optimiert beide behavioralen Komponenten aktiv:                  #
#    Fear of Missing Out wird als korrigierte Portfolio-Schiefe gemessen,      #
#    Fear of Loss als annualisierte Semideviation echter Verluste unter 0.     #
#    Die Zielfunktion max Skew(w)/SemiDev_0(w) bedeutet: positive Schiefe pro  #
#    Einheit Downside-Risiko.                                                  #
#                                                                              #
#  * Bei negativer Schiefe nutzt die Zielfunktion eine stetige Penalty. Damit  #
#    wird die Vorzeichenfalle vermieden: Ein negativer Quotient darf nicht     #
#    durch groessere Semideviation kuenstlich "besser" werden.                 #
#                                                                              #
#  * Die direkte Sample-Schiefe-Optimierung bleibt schaetzsensitiv,            #
#    tail-getrieben und nicht-konvex. Genau deshalb wird der Arnott-Blend als  #
#    zusaetzliches Robustheitsmodell beibehalten: aktive FOMO/FOL-Optimierung  #
#    versus passives Schiefe-Harvesting ueber ein 1/N-Marktbein.               #
#                                                                              #
#  * Ein konzentrierteres Behavioral-Portfolio und schwaechere Krisenperformance#
#    gegenueber Arnott waeren erwartete Trade-offs, keine Codefehler. Diese    #
#    Befunde gehoeren in die Diskussion der Frage "ersetzen oder ergaenzen".   #
#                                                                              #
#  * Modell 3 ist ein transparenter Hybrid-Blend aus MaxSharpe und aktivem     #
#    Behavioral. Dadurch ist klar erkennbar, was durch die Behavioral-         #
#    Komponente ergaenzt wird, ohne weitere Praeferenzparameter einzufuehren.  #
#                                                                              #
#  * Die Box-Constraint max=0.10 je Titel gilt fuer ALLE Modellportfolios.     #
#    Dadurch werden MaxSharpe, Behavioral, Arnott und Hybrid unter identischen #
#    Anlagebedingungen verglichen.                                             #
#                                                                              #
#  * Die Sortino-Kennzahl nutzt annualisierte Target-Semivarianz unter 0. Eine #
#    alternative Spezifikation mit MAR = rf ist schnell implementiert          #
#    (SortinoRatio aus PerformanceAnalytics).                                  #
#                                                                              #
#  * Die Corona-Phase wird bewusst breiter als 02/19-03/23/2020 definiert      #
#    (02/2020-06/2020). Damit beruhen Schiefe und CVaR auf mehr Beobachtungen  #
#    und sind statistisch besser interpretierbar als im reinen Crash-Tief.     #
#                                                                              #
#  * Der risikofreie Zinssatz ist konstant mit 2 % p.a. angesetzt. Das ist     #
#    eine bewusste Vereinfachung und muss im Limitationen-Abschnitt der        #
#    Arbeit genannt werden.                                                    #
################################################################################
