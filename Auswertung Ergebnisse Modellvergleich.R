# -*- coding: UTF-8 -*-
################################################################################
#                                                                              #
#   AUSWERTUNG: BACKTEST-ERGEBNISSE                                            #
#   -----------------------------------------------------------------------    #
#   Liest die Datei `Backtest_Ergebnisse.xlsx` ein und erzeugt:                #
#     1) Aufgewertete Excel mit Master-Tabelle, Jahresrenditen,                #
#        Konzentrationsmetriken, Top-Holdings                                  #
#     2) Publikationsreife Grafiken (PNG 300dpi + Sammel-PDF)                  #
#                                                                              #
#   Ausführung: Rscript "Auswertung Ergebnisse Modellvergleich.R"             #
#   Laufzeit:    abhängig von System, Datenumfang und Grafikexport             #
#                                                                              #
################################################################################

# ==============================================================================
# 1. SETUP
# ==============================================================================

# install.packages(c("openxlsx","xts","zoo","PerformanceAnalytics",
#                    "dplyr","tidyr","ggplot2","scales"))

suppressPackageStartupMessages({
  library(openxlsx)
  library(xts)
  library(zoo)
  library(PerformanceAnalytics)
  library(dplyr)
  library(tidyr)
  library(ggplot2)
  library(scales)
})

# --- Pfade ---
BASE_PATH    <- "F:/FH/Bachelorarbeit/"
INPUT_FILE   <- paste0(BASE_PATH, "Backtest_Ergebnisse.xlsx")
OUTPUT_XLSX  <- paste0(BASE_PATH, "Backtest_Auswertung.xlsx")
PLOTS_DIR    <- paste0(BASE_PATH, "Plots/")
COMBINED_PDF <- paste0(BASE_PATH, "Backtest_Charts.pdf")

dir.create(PLOTS_DIR, showWarnings = FALSE, recursive = TRUE)

# --- Krisen/Boom-Phasen (muss zum Backtest-Skript passen) ---
CRISIS_START <- as.Date("2020-02-01")
CRISIS_END   <- as.Date("2020-06-30")
BOOM_START   <- as.Date("2023-01-01")
BOOM_END     <- as.Date("2023-12-31")

RF_ANNUAL <- 0.02
RF_DAILY  <- (1 + RF_ANNUAL)^(1/252) - 1


# ==============================================================================
# 2. AKADEMISCHES PLOT-THEMA & FARBPALETTE
# ==============================================================================


strategy_colors <- c(
  "MaxSharpe"  = "#2C3E50",  # dunkelblau 
  "Behavioral" = "#C0392B",  # rot 
  "Arnott"     = "#D68910",  # orange 
  "Hybrid"     = "#8E44AD",  # violett 
  "SP500_TR"   = "#27AE60"   # grün 
)
strategy_linetypes <- c(
  "MaxSharpe"  = "solid",
  "Behavioral" = "solid",
  "Arnott"     = "longdash",
  "Hybrid"     = "solid",
  "SP500_TR"   = "dotted"
)

# Strategien, die in der Auswertung berücksichtigt werden.
KEEP_STRATEGIES <- c("MaxSharpe", "Behavioral", "Arnott", "Hybrid", "SP500_TR")

academic_theme <- theme_bw(base_size = 11) +
  theme(
    plot.title       = element_text(size = 12, face = "bold"),
    plot.subtitle    = element_text(size = 10, color = "grey30"),
    plot.caption     = element_text(size = 8,  color = "grey50",
                                    hjust = 0, face = "italic"),
    legend.position  = "bottom",
    legend.title     = element_blank(),
    legend.key.width = unit(1.2, "cm"),
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(color = "grey92", linewidth = 0.3),
    axis.title       = element_text(size = 10),
    strip.background = element_rect(fill = "grey95", color = NA),
    strip.text       = element_text(face = "bold", size = 10)
  )

theme_set(academic_theme)

# Helper für das einheitliche Speichern
save_plot <- function(p, filename, width = 18, height = 11, dpi = 300) {
  ggsave(filename = file.path(PLOTS_DIR, filename),
         plot = p, width = width, height = height,
         units = "cm", dpi = dpi)
}


# ==============================================================================
# 3. DATEN-IMPORT
# ==============================================================================

cat("[1/6] Lese Backtest-Ergebnisse aus", INPUT_FILE, "...\n")

if (!file.exists(INPUT_FILE)) {
  stop("Datei ", INPUT_FILE, " nicht gefunden. ",
       "Bitte zuerst Modellvergleich.R ausführen.")
}

# 3.1 OOS-Renditen
returns_df <- read.xlsx(INPUT_FILE, sheet = "Returns_OOS",
                        detectDates = TRUE)
returns_xts_raw <- xts(returns_df[, -1, drop = FALSE],
                       order.by = as.Date(returns_df$Date))

# Filter: nur die für diese Auswertung relevanten Strategien
available <- intersect(KEEP_STRATEGIES, colnames(returns_xts_raw))
if (length(available) == 0) {
  stop("Keine der erwarteten Strategien (", paste(KEEP_STRATEGIES, collapse=", "),
       ") in der Excel gefunden. Vorhandene Spalten: ",
       paste(colnames(returns_xts_raw), collapse=", "))
}
returns_xts <- returns_xts_raw[, available]
strategies  <- colnames(returns_xts)
# Auch das DataFrame für spätere Plots reduzieren
returns_df  <- returns_df[, c("Date", strategies)]
cat("    -> Strategien:", paste(strategies, collapse = ", "), "\n")
cat("    -> Zeitraum:", as.character(start(returns_xts)),
    "bis", as.character(end(returns_xts)), "\n")

# 3.2 Metadaten aus dem Backtest-Workbook lesen, damit die Auswertung automatisch
# zur aktuellen Modellvergleich-Spezifikation passt.
readme_input <- tryCatch(
  read.xlsx(INPUT_FILE, sheet = "README"),
  error = function(e) NULL
)
get_readme_value <- function(key, fallback = "") {
  if (is.null(readme_input) ||
      !all(c("Eintrag", "Wert") %in% colnames(readme_input))) {
    return(fallback)
  }
  hit <- readme_input$Wert[readme_input$Eintrag == key]
  if (length(hit) == 0 || is.na(hit[1]) || !nzchar(hit[1])) fallback else hit[1]
}

bt_model1 <- get_readme_value(
  "MaxSharpe",
  "(Mean-Variance / Max Sharpe, Tangency, quadprog)"
)
bt_model2 <- get_readme_value(
  "Behavioral",
  "(Behavioral aktiv: max Skew(w) / SemiDev_0(w); FOMO/FOL-Ratio)"
)
bt_arnott <- get_readme_value(
  "Arnott",
  "(Robustheit: FOL-MinSemiCov + FOMO-1/N-Marktproxy mit zeitvariablem alpha)"
)
bt_model3 <- get_readme_value(
  "Hybrid",
  "(Hybrid-Blend aus MaxSharpe und Behavioral-FOMO/FOL)"
)
bt_constraints <- get_readme_value(
  "Gemeinsame Constraints",
  "long-only, voll investiert, gleiche Max-Gewichtung für alle optimierten Modelle"
)
bt_behavioral_params <- get_readme_value(
  "Behavioral-Parameter",
  "Behavioral maximiert Skew(w)/SemiDev_0(w); Arnott bleibt Robustheitsmodell; Hybrid=linearer Blend"
)
bt_behavioral_reason <- get_readme_value(
  "Begründung Behavioral",
  "Für die Forschungsfrage werden FOMO und FOL aktiv in einer Ratio optimiert; die Stabilität wird über Arnott als Robustheitsmodell diskutiert."
)
bt_skew_discussion <- get_readme_value(
  "Diskussion Schiefe",
  "Direkte Sample-Schiefe-Optimierung ist tail-sensitiv und nicht-konvex; OOS-Abweichungen sind als empirischer Trade-off zu diskutieren."
)
bt_rf_limitation <- get_readme_value(
  "Limitation rf",
  paste0("Konstanter risikofreier Zinssatz von ", RF_ANNUAL * 100,
         "% p.a. über den gesamten Backtest; bewusste Vereinfachung.")
)
bt_survivorship <- get_readme_value(
  "Hinweis Survivorship",
  "Datenbasis siehe Backtest-Workbook"
)

# 3.3 Gewichte je Strategie
read_weights <- function(sheet) {
  df <- tryCatch(read.xlsx(INPUT_FILE, sheet = sheet),
                 error = function(e) NULL)
  if (is.null(df)) return(NULL)
  rownames(df) <- df$Rebalancing
  df$Rebalancing <- NULL
  as.matrix(df)
}
W_sharpe <- read_weights("Weights_MaxSharpe")
W_behav  <- read_weights("Weights_Behavioral")
W_arnott <- read_weights("Weights_Arnott")
W_hybrid <- read_weights("Weights_Hybrid")

weights_list <- list(
  MaxSharpe  = W_sharpe,
  Behavioral = W_behav,
  Arnott     = W_arnott,
  Hybrid     = W_hybrid
)
weights_list <- weights_list[!sapply(weights_list, is.null)]


# ==============================================================================
# 4. ERWEITERTE METRIKEN
# ==============================================================================

cat("[2/6] Berechne erweiterte Performance-Metriken ...\n")

#' Berechnet einen umfassenden Satz an Performance-Metriken für eine xts-
#' Renditematrix mit beliebig vielen Strategien als Spalten.
calc_full_metrics <- function(R, rf_annual = RF_ANNUAL, scale = 252) {
  out <- data.frame(row.names = colnames(R))
  
  for (col in colnames(R)) {
    r <- as.numeric(R[, col])
    r <- r[is.finite(r)]
    if (length(r) < 2) next
    
    # === Rendite ===
    cum_ret <- prod(1 + r) - 1
    ann_ret <- (1 + cum_ret)^(scale / length(r)) - 1
    
    # === Risiko ===
    ann_vol     <- sd(r) * sqrt(scale)
    ann_var     <- var(r) * scale
    ann_semivar <- mean(pmin(r, 0)^2) * scale
    ann_semidev <- sqrt(ann_semivar)
    
    # === Höhere Momente ===
    skw  <- {
      n <- length(r); m <- mean(r)
      s2 <- sum((r-m)^2)/n
      if (n < 3 || !is.finite(s2) || s2 <= 1e-20) NA else
        (sqrt(n*(n-1))/(n-2)) * (sum((r-m)^3)/n) / s2^1.5
    }
    krt  <- {
      n <- length(r); m <- mean(r)
      s2 <- sum((r-m)^2)/n
      if (!is.finite(s2) || s2 <= 1e-20) NA else
        ((sum((r-m)^4)/n) / s2^2) - 3
    }
    
    # === Drawdown ===
    mdd <- as.numeric(maxDrawdown(R[, col]))
    
    # === Tail-Risiko ===
    var95  <- as.numeric(quantile(r, 0.05))
    cvar95 <- mean(r[r <= var95])
    
    # === Risk-adjusted Returns ===
    sharpe  <- if (is.finite(ann_vol) && ann_vol > 0) {
      (ann_ret - rf_annual) / ann_vol
    } else {
      NA_real_
    }
    sortino <- if (is.finite(ann_semidev) && ann_semidev > 0) {
      (ann_ret - rf_annual) / ann_semidev
    } else {
      NA_real_
    }
    calmar  <- if (is.finite(mdd) && mdd > 0) ann_ret / mdd else NA_real_
    
    # === Hit Rate / Stabilität ===
    win_rate     <- mean(r > 0)
    best_day     <- max(r)
    worst_day    <- min(r)
    
    out[col, "AnnReturn"]   <- ann_ret
    out[col, "CumReturn"]   <- cum_ret
    out[col, "AnnVol"]      <- ann_vol
    out[col, "AnnVar"]      <- ann_var
    out[col, "AnnSemiVar"]  <- ann_semivar
    out[col, "AnnSemiDev"]  <- ann_semidev
    out[col, "Skewness"]    <- skw
    out[col, "ExcessKurt"]  <- krt
    out[col, "MaxDD"]       <- mdd
    out[col, "VaR_95"]      <- var95
    out[col, "CVaR_95"]     <- cvar95
    out[col, "Sharpe"]      <- sharpe
    out[col, "Sortino"]     <- sortino
    out[col, "Calmar"]      <- calmar
    out[col, "WinRate"]     <- win_rate
    out[col, "BestDay"]     <- best_day
    out[col, "WorstDay"]    <- worst_day
  }
  out
}

metrics_full <- calc_full_metrics(returns_xts)

# Teilperioden
get_subperiod <- function(R, start_d, end_d) {
  R[paste0(start_d, "/", end_d)]
}
returns_crisis <- get_subperiod(returns_xts, CRISIS_START, CRISIS_END)
returns_boom   <- get_subperiod(returns_xts, BOOM_START,   BOOM_END)

metrics_crisis <- if (nrow(returns_crisis) > 5) calc_full_metrics(returns_crisis)
metrics_boom   <- if (nrow(returns_boom)   > 5) calc_full_metrics(returns_boom)


# ==============================================================================
# 5. JAHRESWEISE PERFORMANCE & OUTPERFORMANCE
# ==============================================================================

cat("[3/6] Berechne jahresweise Performance ...\n")

# Jahresrenditen je Strategie (geometrische Aggregation der Tagesrenditen)
yearly_returns <- function(R) {
  yrs <- unique(format(index(R), "%Y"))
  out <- sapply(colnames(R), function(col) {
    sapply(yrs, function(y) {
      r <- as.numeric(R[format(index(R), "%Y") == y, col])
      r <- r[is.finite(r)]
      if (length(r) == 0) return(NA)
      prod(1 + r) - 1
    })
  })
  rownames(out) <- yrs
  as.data.frame(out)
}
ann_perf <- yearly_returns(returns_xts)

# Outperformance gegenüber SP500_TR (zwingend, EqualWeight ist nicht mehr Teil
# der Auswertung).
benchmark_col <- "SP500_TR"
if (!benchmark_col %in% colnames(returns_xts)) {
  warning("Benchmark SP500_TR nicht in den Daten - Outperformance-Sheet wird ",
          "übersprungen.")
  ann_outperf <- data.frame()
} else {
  ann_outperf <- ann_perf - ann_perf[[benchmark_col]]
  ann_outperf[[benchmark_col]] <- NULL
}


# ==============================================================================
# 6. KONZENTRATIONSMETRIKEN AUS GEWICHTEN
# ==============================================================================

cat("[4/6] Analysiere Portfolio-Konzentration ...\n")

# Pro Rebalancing: Anzahl aktive Positionen, HHI, max-Gewicht, Top-10-Anteil
calc_concentration <- function(W) {
  if (is.null(W)) return(NULL)
  out <- data.frame(
    Year = rownames(W),
    NumPositions = NA_integer_,
    HHI = NA_real_,
    EffN = NA_real_,
    MaxWeight = NA_real_,
    Top10Share = NA_real_
  )
  for (i in seq_len(nrow(W))) {
    w <- W[i, ]
    w <- w[!is.na(w) & w > 1e-6]
    if (length(w) == 0) next
    out$NumPositions[i] <- length(w)
    out$HHI[i]          <- sum(w^2)
    out$EffN[i]         <- 1 / sum(w^2)              # effektive Anzahl
    out$MaxWeight[i]    <- max(w)
    out$Top10Share[i]   <- sum(sort(w, decreasing = TRUE)[1:min(10, length(w))])
  }
  out
}
concentration_list <- lapply(weights_list, calc_concentration)

# Top-10-Positionen beim jüngsten Rebalancing je Strategie
top_holdings <- function(W, n = 10) {
  if (is.null(W) || nrow(W) == 0) return(NULL)
  last_year <- rownames(W)[nrow(W)]
  w <- W[nrow(W), ]
  w <- w[!is.na(w)]
  w <- sort(w, decreasing = TRUE)
  top <- head(w[w > 1e-6], n)
  data.frame(
    Year   = last_year,
    Rank   = seq_along(top),
    Ticker = names(top),
    Weight = as.numeric(top)
  )
}
top10_list <- lapply(weights_list, top_holdings, n = 10)




# ==============================================================================
# 7. GRAFIKEN
# ==============================================================================

cat("[5/6] Erstelle Grafiken ...\n")

# Long-Format der Renditen für ggplot
returns_long <- returns_df %>%
  pivot_longer(cols = -Date, names_to = "Strategy", values_to = "Return") %>%
  mutate(Date = as.Date(Date),
         Strategy = factor(Strategy, levels = strategies))

# 7.1 Equity Curves (Wachstum 1 EUR)
cat("    -> Plot 1: Equity Curves\n")
equity_long <- returns_long %>%
  group_by(Strategy) %>%
  arrange(Date) %>%
  mutate(Equity = cumprod(1 + replace_na(Return, 0))) %>%
  ungroup()

p_equity <- ggplot(equity_long, aes(x = Date, y = Equity,
                                    color = Strategy, linetype = Strategy)) +
  geom_line(linewidth = 0.7) +
  scale_color_manual(values = strategy_colors) +
  scale_linetype_manual(values = strategy_linetypes) +
  scale_x_date(breaks = "1 year", labels = date_format("%Y")) +
  scale_y_continuous(labels = label_number(accuracy = 0.1)) +
  labs(
    title    = "Out-of-Sample-Wertentwicklung der Portfolio-Strategien",
    subtitle = "Wachstum von 1 EUR seit Backtest-Start, kumulierte tägliche Renditen",
    x = NULL, y = "Portfolio-Wert (Index, Start = 1)",
    caption  = "Quelle: Eigene Berechnung. Diskrete Tagesrenditen aus S&P 500 Total-Return-Indizes."
  )
save_plot(p_equity, "01_Equity_Curves.png", width = 20, height = 12)

# 7.2 Drawdown-Verlauf
cat("    -> Plot 2: Drawdown\n")
calc_drawdown <- function(r) {
  r[is.na(r)] <- 0
  eq <- cumprod(1 + r)
  eq / cummax(eq) - 1
}

dd_long <- returns_long %>%
  group_by(Strategy) %>%
  arrange(Date) %>%
  mutate(Drawdown = calc_drawdown(Return)) %>%
  ungroup()

p_dd <- ggplot(dd_long, aes(x = Date, y = Drawdown,
                            color = Strategy, linetype = Strategy)) +
  geom_line(linewidth = 0.6) +
  scale_color_manual(values = strategy_colors) +
  scale_linetype_manual(values = strategy_linetypes) +
  scale_x_date(breaks = "1 year", labels = date_format("%Y")) +
  scale_y_continuous(labels = percent_format(accuracy = 1)) +
  labs(
    title    = "Drawdown-Verlauf der Strategien",
    subtitle = "Underwater-Plot: prozentualer Rückgang vom letzten Höchststand",
    x = NULL, y = "Drawdown",
    caption = "Quelle: Eigene Berechnung."
  )
save_plot(p_dd, "02_Drawdown.png", width = 20, height = 10)

# 7.3 Rolling 60-Tage-Volatilität (annualisiert)
cat("    -> Plot 3: Rolling Volatility\n")
roll_vol <- function(r, k = 60) {
  rollapply(r, width = k, FUN = sd, by.column = TRUE,
            align = "right", fill = NA) * sqrt(252)
}
vol_xts <- roll_vol(returns_xts, k = 60)
vol_df  <- data.frame(Date = index(vol_xts), coredata(vol_xts))
vol_long <- vol_df %>%
  pivot_longer(cols = -Date, names_to = "Strategy", values_to = "Vol") %>%
  mutate(Strategy = factor(Strategy, levels = strategies))

p_vol <- ggplot(vol_long, aes(x = Date, y = Vol,
                              color = Strategy, linetype = Strategy)) +
  geom_line(linewidth = 0.5, na.rm = TRUE) +
  scale_color_manual(values = strategy_colors) +
  scale_linetype_manual(values = strategy_linetypes) +
  scale_x_date(breaks = "1 year", labels = date_format("%Y")) +
  scale_y_continuous(labels = percent_format(accuracy = 1)) +
  labs(
    title    = "Rollierende 60-Tage-Volatilität (annualisiert)",
    subtitle = "Stichprobenstandardabweichung der Tagesrenditen, skaliert mit sqrt(252)",
    x = NULL, y = "Volatilität (annualisiert)",
    caption = "Quelle: Eigene Berechnung."
  )
save_plot(p_vol, "03_Rolling_Volatility.png", width = 20, height = 10)

# 7.4 Jahresrenditen als Bar-Plot
cat("    -> Plot 4: Annual Returns\n")
ann_long <- ann_perf %>%
  mutate(Year = rownames(ann_perf)) %>%
  pivot_longer(cols = -Year, names_to = "Strategy", values_to = "Return") %>%
  mutate(Strategy = factor(Strategy, levels = strategies))

p_annual <- ggplot(ann_long, aes(x = Year, y = Return, fill = Strategy)) +
  geom_col(position = position_dodge(width = 0.8), width = 0.75) +
  geom_hline(yintercept = 0, color = "black", linewidth = 0.3) +
  scale_fill_manual(values = strategy_colors) +
  scale_y_continuous(labels = percent_format(accuracy = 1)) +
  labs(
    title    = "Kalenderjährliche Renditen je Strategie",
    subtitle = "Geometrische Aggregation der täglichen OOS-Renditen pro Jahr",
    x = NULL, y = "Jahresrendite",
    caption = "Quelle: Eigene Berechnung."
  )
save_plot(p_annual, "04_Annual_Returns.png", width = 20, height = 11)

# 7.5 Performance-Metriken Heatmap (Z-Standardisierung pro Metrik)
cat("    -> Plot 5: Metrics Heatmap\n")
heatmap_metrics <- c("AnnReturn", "AnnVol", "Sharpe", "Sortino",
                     "MaxDD", "CVaR_95", "Skewness", "Calmar")
m_sub <- metrics_full[, heatmap_metrics, drop = FALSE]
m_z <- scale(m_sub)

neg_metrics <- c("AnnVol", "MaxDD")
m_z[, neg_metrics] <- -m_z[, neg_metrics]

m_z_df <- as.data.frame(m_z) %>%
  mutate(Strategy = rownames(m_z)) %>%
  pivot_longer(-Strategy, names_to = "Metric", values_to = "Z") %>%
  mutate(
    Metric = factor(Metric, levels = heatmap_metrics),
    Strategy = factor(Strategy, levels = strategies),
    # RawValue zeilenweise: für jede (Strategy, Metric)-Kombination genau
    # eine Zelle aus metrics_full. Direkte Matrix-Indizierung mit zwei
    # Vektoren liefert das Kreuzprodukt - daher mapply.
    RawValue = mapply(
      function(s, m) metrics_full[as.character(s), as.character(m)],
      Strategy, Metric
    )
  )

p_heat <- ggplot(m_z_df, aes(x = Metric, y = Strategy, fill = Z)) +
  geom_tile(color = "white", linewidth = 0.6) +
  geom_text(aes(label = sprintf("%.3f", RawValue)),
            color = "black", size = 3) +
  scale_fill_gradient2(low = "#C0392B", mid = "white", high = "#27AE60",
                       midpoint = 0,
                       name = "z-Score\n(höher = besser)") +
  labs(
    title    = "Performance-Metriken im Strategievergleich",
    subtitle = "Z-Standardisierung je Metrik (Volatilität und MDD mit invertiertem Vorzeichen; Schiefe als OOS-FOMO-Kennzahl)",
    x = NULL, y = NULL,
    caption  = "Quelle: Eigene Berechnung. Werte in Zellen sind Rohgrößen, Färbung ist relativ je Spalte."
  ) +
  theme(legend.position = "right",
        axis.text.x = element_text(angle = 30, hjust = 1))
save_plot(p_heat, "05_Metrics_Heatmap.png", width = 22, height = 10)

# 7.6 Teilperioden-Vergleich (Krise + Boom)
cat("    -> Plot 6: Krise vs Boom\n")
prep_subp <- function(R, label) {
  if (is.null(R) || nrow(R) == 0) return(NULL)
  eq <- apply(R, 2, function(x) {
    x[is.na(x)] <- 0
    cumprod(1 + x)
  })
  df <- data.frame(Date = index(R), eq) %>%
    pivot_longer(cols = -Date, names_to = "Strategy", values_to = "Equity") %>%
    mutate(Phase = label,
           Strategy = factor(Strategy, levels = strategies))
  df
}
df_crisis <- prep_subp(returns_crisis, paste0("Corona-Stressphase (",
                                              CRISIS_START, " - ",
                                              CRISIS_END, ")"))
df_boom   <- prep_subp(returns_boom,   paste0("Tech-Rallye (",
                                              BOOM_START, " - ",
                                              BOOM_END, ")"))
df_phases <- bind_rows(df_crisis, df_boom)

p_phases <- ggplot(df_phases, aes(x = Date, y = Equity,
                                  color = Strategy, linetype = Strategy)) +
  geom_line(linewidth = 0.7) +
  facet_wrap(~ Phase, scales = "free_x", ncol = 2) +
  scale_color_manual(values = strategy_colors) +
  scale_linetype_manual(values = strategy_linetypes) +
  scale_y_continuous(labels = label_number(accuracy = 0.01)) +
  labs(
    title    = "Performance in Stress- und Boom-Phasen",
    subtitle = "Wachstum von 1 EUR über den jeweiligen Teilzeitraum",
    x = NULL, y = "Wert (Index, Start = 1)",
    caption = "Quelle: Eigene Berechnung."
  )
save_plot(p_phases, "06_Krise_Boom.png", width = 22, height = 10)

# 7.7 Effektive Positionen über Zeit
cat("    -> Plot 7-8: Portfolio-Konzentration\n")
conc_df <- bind_rows(lapply(names(concentration_list), function(strat) {
  df <- concentration_list[[strat]]
  if (is.null(df)) return(NULL)
  df$Strategy <- strat
  df
}))
conc_df$Strategy <- factor(conc_df$Strategy,
                           levels = intersect(strategies, unique(conc_df$Strategy)))

p_effn <- ggplot(conc_df, aes(x = Year, y = EffN, color = Strategy,
                              group = Strategy, linetype = Strategy)) +
  geom_line(linewidth = 0.7) +
  geom_point(size = 2) +
  scale_color_manual(values = strategy_colors) +
  scale_linetype_manual(values = strategy_linetypes) +
  labs(
    title    = "Effektive Anzahl von Positionen im Zeitverlauf",
    subtitle = "EffN = 1 / HHI; höhere Werte bedeuten bessere Diversifikation",
    x = "Rebalancing-Jahr", y = "Effektive Positionen (1/HHI)",
    caption = "Quelle: Eigene Berechnung."
  )
save_plot(p_effn, "07_Konzentration_EffN.png", width = 18, height = 10)

# 7.8 Anzahl aktiver Positionen je Rebalancing
p_npos <- ggplot(conc_df, aes(x = Year, y = NumPositions, fill = Strategy)) +
  geom_col(position = position_dodge(width = 0.8), width = 0.75) +
  scale_fill_manual(values = strategy_colors) +
  labs(
    title    = "Anzahl aktiver Positionen je Rebalancing",
    subtitle = "Positionen mit Gewicht > 0,0001 % gezählt",
    x = "Rebalancing-Jahr", y = "Anzahl aktiver Titel",
    caption = "Quelle: Eigene Berechnung."
  )
save_plot(p_npos, "08_Konzentration_AnzPos.png", width = 18, height = 10)

# 7.9 Top-10-Positionen je Strategie (jüngstes Rebalancing)
cat("    -> Plot 9: Top-10 Holdings\n")
top10_df <- bind_rows(lapply(names(top10_list), function(strat) {
  df <- top10_list[[strat]]
  if (is.null(df)) return(NULL)
  df$Strategy <- strat
  df
}))

if (nrow(top10_df) > 0) {
  top10_df$Strategy <- factor(top10_df$Strategy,
                              levels = intersect(strategies,
                                                 unique(top10_df$Strategy)))
  p_top10 <- ggplot(top10_df,
                    aes(x = reorder(Ticker, Weight), y = Weight,
                        fill = Strategy)) +
    geom_col() +
    coord_flip() +
    facet_wrap(~ Strategy, scales = "free_y") +
    scale_fill_manual(values = strategy_colors, guide = "none") +
    scale_y_continuous(labels = percent_format(accuracy = 1)) +
    labs(
      title    = "Top-10-Positionen beim letzten Rebalancing",
      subtitle = paste("Stand:", unique(top10_df$Year)[1]),
      x = NULL, y = "Portfolio-Gewicht",
      caption = "Quelle: Eigene Berechnung."
    )
  save_plot(p_top10, "09_Top10_Holdings.png", width = 20, height = 12)
}

# 7.10 Schiefe + Semivarianz pro Strategie - Ex-post-Einordnung des
#      FOL/FOMO-Profils. Behavioral optimiert die Ratio in-sample aktiv,
#      Arnott dient als passiver Harvesting-Vergleich.
cat("    -> Plot 10: Schiefe vs. Semivarianz\n")
metrics_plot_df <- data.frame(
  Strategy = factor(rownames(metrics_full), levels = strategies),
  Skewness = metrics_full$Skewness,
  SemiDev  = metrics_full$AnnSemiDev,
  Sharpe   = metrics_full$Sharpe,
  AnnRet   = metrics_full$AnnReturn
)

p_skew_semi <- ggplot(metrics_plot_df,
                      aes(x = SemiDev, y = Skewness, color = Strategy)) +
  geom_hline(yintercept = 0, color = "grey60", linetype = "dashed") +
  geom_point(size = 5, alpha = 0.85) +
  geom_text(aes(label = Strategy),
            nudge_y = 0.05, size = 3.5, fontface = "bold", show.legend = FALSE) +
  scale_color_manual(values = strategy_colors, guide = "none") +
  scale_x_continuous(labels = percent_format(accuracy = 1),
                     limits = c(0, max(metrics_plot_df$SemiDev) * 1.15)) +
  labs(
    title    = "Risiko- und Schiefe-Trade-off: Schiefe vs. Semideviation",
    subtitle = "Ex-post-Profil: aktive FOMO/FOL-Ratio versus Arnott-Harvesting",
    x = "Annualisierte Semideviation (FOL-Risiko)",
    y = "Schiefe (FOMO-Indikator: höher = günstiger)",
    caption  = "Quelle: Eigene Berechnung. Behavioral optimiert Skew/SemiDev in-sample; Arnott bildet FOMO passiv über 1/N ab."
  )
save_plot(p_skew_semi, "10_Skew_Semi_Scatter.png", width = 18, height = 12)

# 7.11 Renditeverteilung je Strategie als Density-Plot
#      Zeigt die gesamte Verteilungsform, nicht nur Mittelwert und
#      Standardabweichung. Dadurch werden Tail-Eigenschaften und Schiefe sichtbar.
cat("    -> Plot 11: Renditeverteilungen\n")
density_long <- returns_long %>%
  filter(!is.na(Return), abs(Return) < 0.20)  # Extreme Ausreißer für Lesbarkeit abschneiden

p_density <- ggplot(density_long, aes(x = Return, fill = Strategy, color = Strategy)) +
  geom_density(alpha = 0.25, linewidth = 0.7) +
  geom_vline(xintercept = 0, color = "grey40", linetype = "dashed") +
  scale_fill_manual(values = strategy_colors) +
  scale_color_manual(values = strategy_colors) +
  scale_x_continuous(labels = percent_format(accuracy = 1)) +
  coord_cartesian(xlim = c(-0.05, 0.05)) +
  labs(
    title    = "Verteilung der täglichen Renditen",
    subtitle = "Density-Plot der Tagesrenditen mit Fokus auf Schiefe und Tail-Struktur",
    x = "Tagesrendite", y = "Dichte",
    caption = "Quelle: Eigene Berechnung. Ausreißer jenseits von ±5 % für bessere Lesbarkeit ausgeblendet."
  ) +
  theme(plot.margin = margin(14, 24, 18, 14))
save_plot(p_density, "11_Renditeverteilungen.png", width = 20, height = 11)

# 7.12 Sharpe vs Sortino Vergleich - klassisch vs Behavioral-adjustiert
cat("    -> Plot 12: Sharpe vs Sortino\n")
ratio_df <- data.frame(
  Strategy = factor(rownames(metrics_full), levels = strategies),
  Sharpe   = metrics_full$Sharpe,
  Sortino  = metrics_full$Sortino
) %>%
  pivot_longer(cols = c(Sharpe, Sortino), names_to = "Ratio", values_to = "Wert")

p_ratios <- ggplot(ratio_df, aes(x = Strategy, y = Wert, fill = Ratio)) +
  geom_col(position = position_dodge(width = 0.8), width = 0.7) +
  geom_text(aes(label = sprintf("%.2f", Wert)),
            position = position_dodge(width = 0.8),
            vjust = -0.5, size = 3.2) +
  scale_fill_manual(values = c("Sharpe" = "#34495E", "Sortino" = "#16A085")) +
  scale_y_continuous(
    limits = c(0, max(ratio_df$Wert, na.rm = TRUE) * 1.18),
    expand = expansion(mult = c(0, 0.02))
  ) +
  labs(
    title    = "Risikoadjustierte Renditen: Sharpe- vs. Sortino-Ratio",
    subtitle = "Die Sortino-Ratio berücksichtigt nur Downside-Risiko und ergänzt damit den Sharpe-Vergleich",
    x = NULL, y = "Ratio (annualisiert)",
    caption = "Quelle: Eigene Berechnung. Sortino verwendet Semideviation statt Standardabweichung im Nenner."
  ) +
  theme(plot.margin = margin(14, 24, 18, 14))
save_plot(p_ratios, "12_Sharpe_vs_Sortino.png", width = 18, height = 11)

# 7.13 Outperformance gegenüber S&P 500 als Cumulative-Difference-Plot
#      Zeigt sehr klar, wann jede Strategie besser/schlechter als der Markt ist
cat("    -> Plot 13: Outperformance vs. Benchmark\n")
if ("SP500_TR" %in% colnames(returns_xts)) {
  bench <- as.numeric(returns_xts[, "SP500_TR"])
  outperf_long <- returns_long %>%
    filter(Strategy != "SP500_TR") %>%
    group_by(Strategy) %>%
    arrange(Date) %>%
    mutate(
      Excess  = Return - bench[match(Date, index(returns_xts))],
      CumOutp = cumsum(replace_na(Excess, 0))
    ) %>%
    ungroup()
  
  p_outperf <- ggplot(outperf_long, aes(x = Date, y = CumOutp,
                                        color = Strategy, linetype = Strategy)) +
    geom_hline(yintercept = 0, color = "grey40", linewidth = 0.4) +
    geom_line(linewidth = 0.7) +
    scale_color_manual(values = strategy_colors) +
    scale_linetype_manual(values = strategy_linetypes) +
    scale_x_date(breaks = "1 year", labels = date_format("%Y")) +
    scale_y_continuous(labels = percent_format(accuracy = 1)) +
    labs(
      title    = "Kumulierte Outperformance gegenüber S&P 500 TR",
      subtitle = "Steigend = Strategie schlägt Benchmark, fallend = Underperformance",
      x = NULL, y = "Kumulierte Überrendite",
      caption = "Quelle: Eigene Berechnung."
    )
  save_plot(p_outperf, "13_Outperformance.png", width = 20, height = 11)
}

# 7.14 Korrelationsmatrix der Strategien als Heatmap
cat("    -> Plot 14: Korrelationsmatrix\n")
cor_mat <- cor(coredata(returns_xts), use = "pairwise.complete.obs")
cor_long <- as.data.frame(cor_mat) %>%
  mutate(Strategy_i = factor(rownames(cor_mat), levels = strategies)) %>%
  pivot_longer(-Strategy_i, names_to = "Strategy_j", values_to = "Corr") %>%
  mutate(Strategy_j = factor(Strategy_j, levels = strategies))

p_cor <- ggplot(cor_long, aes(x = Strategy_i, y = Strategy_j, fill = Corr)) +
  geom_tile(color = "white", linewidth = 0.6) +
  geom_text(aes(label = sprintf("%.2f", Corr)), color = "black", size = 4) +
  scale_fill_gradient2(low = "#3498DB", mid = "white", high = "#E74C3C",
                       midpoint = 0.7, limits = c(0, 1),
                       name = "Korrelation") +
  labs(
    title    = "Korrelationsmatrix der Strategie-Renditen",
    subtitle = "Rot = ähnliche Performance, Blau = geringere Korrelation",
    x = NULL, y = NULL,
    caption = "Quelle: Eigene Berechnung."
  ) +
  coord_fixed(clip = "off") +
  theme(
    legend.position = "right",
    plot.title.position = "plot",
    plot.caption.position = "plot",
    plot.subtitle = element_text(size = 9, margin = margin(b = 8)),
    plot.caption = element_text(size = 7, hjust = 0, margin = margin(t = 8)),
    plot.margin = margin(14, 24, 18, 14),
    axis.text.x = element_text(margin = margin(t = 6)),
    axis.text.y = element_text(margin = margin(r = 6))
  )
save_plot(p_cor, "14_Korrelationsmatrix.png", width = 18, height = 12)

# 7.15 Performance-Vergleich Krise vs Boom als Bar-Chart
cat("    -> Plot 15: Krise vs Boom Performance-Bars\n")
if (!is.null(metrics_crisis) && !is.null(metrics_boom)) {
  phase_perf <- data.frame(
    Strategy = factor(rep(rownames(metrics_crisis), 2), levels = strategies),
    Phase    = factor(rep(c("Corona-Stressphase 2020", "Tech-Rallye 2023"),
                          each = nrow(metrics_crisis)),
                      levels = c("Corona-Stressphase 2020", "Tech-Rallye 2023")),
    CumRet   = c(metrics_crisis$CumReturn, metrics_boom$CumReturn)
  )
  
  p_phase_perf <- ggplot(phase_perf,
                         aes(x = Strategy, y = CumRet, fill = Strategy)) +
    geom_col(width = 0.75) +
    geom_hline(yintercept = 0, color = "black", linewidth = 0.3) +
    geom_text(aes(label = sprintf("%.1f%%", CumRet*100)),
              vjust = ifelse(phase_perf$CumRet >= 0, -0.5, 1.3),
              size = 3.2, fontface = "bold") +
    facet_wrap(~ Phase, scales = "free_y") +
    scale_fill_manual(values = strategy_colors, guide = "none") +
    scale_y_continuous(
      labels = percent_format(accuracy = 1),
      expand = expansion(mult = c(0.18, 0.22))
    ) +
    coord_cartesian(clip = "off") +
    labs(
      title    = "Teilperioden-Performance: Krise vs. Boom",
      subtitle = "Kumulierte Renditen über den jeweiligen Teilzeitraum",
      x = NULL, y = "Kumulierte Rendite",
      caption = "Quelle: Eigene Berechnung."
    ) +
    theme(
      plot.margin = margin(16, 24, 18, 16),
      strip.text = element_text(margin = margin(t = 6, b = 6))
    )
  save_plot(p_phase_perf, "15_Krise_Boom_Bars.png", width = 22, height = 11)
}

# 7.16 Sammel-PDF mit allen Charts
cat("    -> Sammel-PDF\n")
pdf_device_open <- FALSE
tryCatch({
  pdf(COMBINED_PDF, width = 11, height = 7, onefile = TRUE)
  pdf_device_open <- TRUE
  print(p_equity)
  print(p_dd)
  print(p_vol)
  print(p_annual)
  print(p_heat)
  print(p_phases)
  print(p_effn)
  print(p_npos)
  if (exists("p_top10"))      print(p_top10)
  print(p_skew_semi)
  print(p_density)
  print(p_ratios)
  if (exists("p_outperf"))    print(p_outperf)
  print(p_cor)
  if (exists("p_phase_perf")) print(p_phase_perf)
}, error = function(e) {
  warning("Sammel-PDF konnte nicht geschrieben werden. ",
          "Bitte Backtest_Charts.pdf schließen und die Auswertung erneut starten. ",
          "Die einzelnen PNG-Grafiken wurden bereits erzeugt. Details: ",
          e$message)
}, finally = {
  if (pdf_device_open && grDevices::dev.cur() > 1) {
    invisible(grDevices::dev.off())
  }
})


# ==============================================================================
# 8. PRÄSENTATIONSFERTIGE EXCEL-DATEI MIT EINGEBETTETEN GRAFIKEN
# ==============================================================================

cat("[6/6] Schreibe präsentationsfertige Excel-Datei ...\n")

wb <- createWorkbook()

# ---- Konsistente Farbpalette und Styles --------------------------------------
COL_DARK    <- "#1F3A5F"   # primary dark blue (Cover, Titel)
COL_ACCENT  <- "#2E86AB"   # secondary blue (Section-Header)
COL_GOOD    <- "#27AE60"   # positiv
COL_BAD     <- "#C0392B"   # negativ
COL_NEUTRAL <- "#F4F4F4"   # zarter Hintergrund

# Title-Styles
style_main_title <- createStyle(textDecoration = "bold", fontSize = 22,
                                fontColour = "white", fgFill = COL_DARK,
                                halign = "center", valign = "center",
                                border = "TopBottomLeftRight",
                                borderColour = COL_DARK)
style_subtitle   <- createStyle(textDecoration = "italic", fontSize = 11,
                                fontColour = "#5D6D7E", halign = "center")
style_section    <- createStyle(textDecoration = "bold", fontSize = 13,
                                fontColour = "white", fgFill = COL_ACCENT,
                                halign = "left", valign = "center",
                                border = "TopBottomLeftRight",
                                borderColour = COL_ACCENT)

# Tabellen-Styles
style_th         <- createStyle(textDecoration = "bold", fontSize = 11,
                                fontColour = "white", fgFill = COL_DARK,
                                halign = "center", valign = "center",
                                border = "TopBottomLeftRight",
                                borderColour = "white", wrapText = TRUE)
style_row_label  <- createStyle(textDecoration = "bold",
                                fgFill = "#EAECEE",
                                border = "TopBottomLeftRight",
                                borderColour = "#D5DBDB")
style_pct        <- createStyle(numFmt = "0.00%",
                                border = "TopBottomLeftRight",
                                borderColour = "#D5DBDB", halign = "right")
style_num2       <- createStyle(numFmt = "0.00",
                                border = "TopBottomLeftRight",
                                borderColour = "#D5DBDB", halign = "right")
style_num4       <- createStyle(numFmt = "0.0000",
                                border = "TopBottomLeftRight",
                                borderColour = "#D5DBDB", halign = "right")
style_date       <- createStyle(numFmt = "yyyy-mm-dd",
                                border = "TopBottomLeftRight",
                                borderColour = "#D5DBDB")
style_text       <- createStyle(wrapText = TRUE, valign = "top",
                                border = "TopBottomLeftRight",
                                borderColour = "#D5DBDB",
                                fgFill = "#FBFCFC")
style_callout    <- createStyle(textDecoration = "bold",
                                fontColour = "white",
                                fgFill = COL_GOOD, halign = "center")

# Sicheres Image-Insert (fällt still aus, falls Datei fehlt)
insertImage_safe <- function(sheet, file, startCol, startRow,
                             width = 22, height = 12) {
  if (file.exists(file)) {
    insertImage(wb, sheet = sheet, file = file,
                startCol = startCol, startRow = startRow,
                width = width, height = height, units = "cm")
  }
}


# ==============================================================================
# SHEET 1 — COVER
# ==============================================================================
addWorksheet(wb, "01_Cover", gridLines = FALSE, tabColour = COL_DARK)

mergeCells(wb, "01_Cover", cols = 2:8, rows = 3:5)
writeData(wb, "01_Cover", "Behavioral Finance im Portfoliomanagement",
          startCol = 2, startRow = 3)
addStyle(wb, "01_Cover", style_main_title, rows = 3:5, cols = 2:8,
         gridExpand = TRUE)

mergeCells(wb, "01_Cover", cols = 2:8, rows = 6:6)
writeData(wb, "01_Cover",
          "Empirische Out-of-Sample-Evaluation: Mean-Variance vs. Behavioral vs. Arnott vs. Hybrid",
          startCol = 2, startRow = 6)
addStyle(wb, "01_Cover", style_subtitle, rows = 6, cols = 2:8,
         gridExpand = TRUE)

mergeCells(wb, "01_Cover", cols = 2:8, rows = 9:9)
writeData(wb, "01_Cover", "STUDIENDESIGN", startCol = 2, startRow = 9)
addStyle(wb, "01_Cover", style_section, rows = 9, cols = 2:8,
         gridExpand = TRUE)

cover_design <- data.frame(
  Aspekt = c("Anlageuniversum", "Backtest-Zeitraum", "Rebalancing",
             "Optimierungsverfahren", "Risikofreier Zinssatz",
             "Begründung Behavioral", "Limitation rf", "Benchmark"),
  Wert = c(paste("S&P-500-Index aus Backtest-Datei.", bt_survivorship),
           paste(format(start(returns_xts), "%d.%m.%Y"),
                 "bis", format(end(returns_xts), "%d.%m.%Y")),
           "Jährlich (31.12.) im Expanding-Window-Verfahren",
           "quadprog (MaxSharpe/FOL-MinSemiCov) + aktive FOMO/FOL-Ratio-Optimierung; Arnott als Robustheitsmodell",
           paste0(format(RF_ANNUAL*100, nsmall=1), " % p.a."),
           bt_behavioral_reason,
           bt_rf_limitation,
           "S&P-500-Total-Return-Index (^SP500TR via Yahoo Finance)"),
  stringsAsFactors = FALSE
)
writeData(wb, "01_Cover", cover_design, startCol = 2, startRow = 11,
          colNames = TRUE)
addStyle(wb, "01_Cover", style_th, rows = 11, cols = 2:3, gridExpand = TRUE)
addStyle(wb, "01_Cover", style_row_label,
         rows = 12:(11+nrow(cover_design)), cols = 2, gridExpand = TRUE)
addStyle(wb, "01_Cover", style_text,
         rows = 12:(11+nrow(cover_design)), cols = 3, gridExpand = TRUE)

mergeCells(wb, "01_Cover", cols = 2:8,
           rows = (11+nrow(cover_design)+2):(11+nrow(cover_design)+2))
writeData(wb, "01_Cover", "DIE STRATEGIEN IM VERGLEICH",
          startCol = 2, startRow = 11+nrow(cover_design)+2)
addStyle(wb, "01_Cover", style_section,
         rows = 11+nrow(cover_design)+2, cols = 2:8, gridExpand = TRUE)

strategies_explained <- data.frame(
  Strategie = c("Mean-Variance (MaxSharpe)",
                "Behavioral (aktive FOMO/FOL-Ratio)",
                "Arnott-Blend (Robustheit)",
                "Hybrid-Blend (MaxSharpe + Behavioral)",
                "S&P 500 TR (Benchmark)"),
  Zielfunktion = c("max (Rendite - rf) / Volatilität",
                   bt_model2,
                   bt_arnott,
                   bt_model3,
                   "S&P-500-Total-Return-Tagesrendite als Benchmark-Serie"),
  Theoretische_Grundlage = c("Rationaler Investor (Markowitz, Sharpe)",
                             "Positive Schiefe pro Einheit Downside-Risiko",
                             "Passives Schiefe-Harvesting",
                             "Erweiterte Präferenzen mit beiden Komponenten",
                             "Index-Benchmark als Tagesrendite-Serie; keine driftenden Einzeltitelgewichte"),
  stringsAsFactors = FALSE
)
writeData(wb, "01_Cover", strategies_explained, startCol = 2,
          startRow = 11+nrow(cover_design)+4, colNames = TRUE)
addStyle(wb, "01_Cover", style_th,
         rows = 11+nrow(cover_design)+4, cols = 2:4, gridExpand = TRUE)
addStyle(wb, "01_Cover", style_row_label,
         rows = (11+nrow(cover_design)+5):(11+nrow(cover_design)+4+nrow(strategies_explained)),
         cols = 2, gridExpand = TRUE)
addStyle(wb, "01_Cover", style_text,
         rows = (11+nrow(cover_design)+5):(11+nrow(cover_design)+4+nrow(strategies_explained)),
         cols = 3:4, gridExpand = TRUE)

mergeCells(wb, "01_Cover", cols = 2:8, rows = 30:30)
writeData(wb, "01_Cover",
          paste("Auswertung erstellt am:", format(Sys.time(), "%d.%m.%Y %H:%M")),
          startCol = 2, startRow = 30)
addStyle(wb, "01_Cover", style_subtitle, rows = 30, cols = 2:8,
         gridExpand = TRUE)

setColWidths(wb, "01_Cover", cols = 1, widths = 3)
setColWidths(wb, "01_Cover", cols = 2, widths = 32)
setColWidths(wb, "01_Cover", cols = 3, widths = 55)
setColWidths(wb, "01_Cover", cols = 4, widths = 50)
setColWidths(wb, "01_Cover", cols = 5:8, widths = 14)
setRowHeights(wb, "01_Cover", rows = 3, heights = 38)
setRowHeights(wb, "01_Cover", rows = 4:5, heights = 22)
setRowHeights(wb, "01_Cover", rows = 6, heights = 22)


# ==============================================================================
# SHEET 2 — EXECUTIVE SUMMARY
# ==============================================================================
addWorksheet(wb, "02_Executive_Summary", gridLines = FALSE,
             tabColour = COL_ACCENT)

mergeCells(wb, "02_Executive_Summary", cols = 2:9, rows = 2:2)
writeData(wb, "02_Executive_Summary", "Executive Summary",
          startCol = 2, startRow = 2)
addStyle(wb, "02_Executive_Summary", style_main_title,
         rows = 2, cols = 2:9, gridExpand = TRUE)
setRowHeights(wb, "02_Executive_Summary", rows = 2, heights = 38)

# Forschungsfrage
mergeCells(wb, "02_Executive_Summary", cols = 2:9, rows = 5:5)
writeData(wb, "02_Executive_Summary", "FORSCHUNGSFRAGE",
          startCol = 2, startRow = 5)
addStyle(wb, "02_Executive_Summary", style_section, rows = 5, cols = 2:9,
         gridExpand = TRUE)

mergeCells(wb, "02_Executive_Summary", cols = 2:9, rows = 6:7)
writeData(wb, "02_Executive_Summary",
          paste("Inwieweit können Behavioral-Finance-Maße wie Fear of Loss (Semivarianz) und Fear of Missing Out (Schiefe) im Portfoliomanagement 
klassische Risikomaße (Varianz, Beta) ersetzen oder ergänzen?"),
          startCol = 2, startRow = 6)
addStyle(wb, "02_Executive_Summary", style_text, rows = 6:7, cols = 2:9,
         gridExpand = TRUE)
setRowHeights(wb, "02_Executive_Summary", rows = 6:7, heights = 24)

# Kernergebnisse: Best/Worst je Metrik
mergeCells(wb, "02_Executive_Summary", cols = 2:9, rows = 10:10)
writeData(wb, "02_Executive_Summary", "KERNERGEBNISSE",
          startCol = 2, startRow = 10)
addStyle(wb, "02_Executive_Summary", style_section, rows = 10, cols = 2:9,
         gridExpand = TRUE)

strat_no_bench <- setdiff(strategies, "SP500_TR")
m_sub <- metrics_full[strat_no_bench, , drop = FALSE]

summary_metrics <- data.frame(
  Metrik = c("Höchste annualisierte Rendite",
             "Beste Sharpe-Ratio",
             "Beste Sortino-Ratio",
             "Geringster Maximum Drawdown",
             "Höchste Schiefe (FOMO)",
             "Geringste Semideviation (FOL)",
             "Beste Calmar-Ratio"),
  Strategie = c(strat_no_bench[which.max(m_sub$AnnReturn)],
                strat_no_bench[which.max(m_sub$Sharpe)],
                strat_no_bench[which.max(m_sub$Sortino)],
                strat_no_bench[which.min(m_sub$MaxDD)],
                strat_no_bench[which.max(m_sub$Skewness)],
                strat_no_bench[which.min(m_sub$AnnSemiDev)],
                strat_no_bench[which.max(m_sub$Calmar)]),
  Wert = c(sprintf("%.2f%%", 100 * max(m_sub$AnnReturn)),
           sprintf("%.3f", max(m_sub$Sharpe)),
           sprintf("%.3f", max(m_sub$Sortino)),
           sprintf("%.2f%%", 100 * min(m_sub$MaxDD)),
           sprintf("%.3f", max(m_sub$Skewness)),
           sprintf("%.2f%%", 100 * min(m_sub$AnnSemiDev)),
           sprintf("%.3f", max(m_sub$Calmar))),
  stringsAsFactors = FALSE
)
writeData(wb, "02_Executive_Summary", summary_metrics,
          startCol = 2, startRow = 12, colNames = TRUE)
addStyle(wb, "02_Executive_Summary", style_th,
         rows = 12, cols = 2:4, gridExpand = TRUE)
addStyle(wb, "02_Executive_Summary", style_row_label,
         rows = 13:(12+nrow(summary_metrics)), cols = 2, gridExpand = TRUE)
addStyle(wb, "02_Executive_Summary", style_callout,
         rows = 13:(12+nrow(summary_metrics)), cols = 3, gridExpand = TRUE)
addStyle(wb, "02_Executive_Summary", style_text,
         rows = 13:(12+nrow(summary_metrics)), cols = 4, gridExpand = TRUE)

# Zentrale Kennzahlen
key_table_row <- 12 + nrow(summary_metrics) + 3
mergeCells(wb, "02_Executive_Summary", cols = 2:9,
           rows = key_table_row:key_table_row)
writeData(wb, "02_Executive_Summary",
          "ZENTRALE KENNZAHLEN IM STRATEGIEVERGLEICH",
          startCol = 2, startRow = key_table_row)
addStyle(wb, "02_Executive_Summary", style_section,
         rows = key_table_row, cols = 2:9, gridExpand = TRUE)

key_metrics <- data.frame(
  Strategie  = rownames(metrics_full),
  Rendite    = metrics_full$AnnReturn,
  Volatil    = metrics_full$AnnVol,
  Sharpe     = metrics_full$Sharpe,
  Sortino    = metrics_full$Sortino,
  MaxDD      = metrics_full$MaxDD,
  Schiefe    = metrics_full$Skewness,
  CVaR_95    = metrics_full$CVaR_95,
  stringsAsFactors = FALSE
)
writeData(wb, "02_Executive_Summary", key_metrics,
          startCol = 2, startRow = key_table_row + 2, colNames = TRUE)
addStyle(wb, "02_Executive_Summary", style_th,
         rows = key_table_row + 2, cols = 2:9, gridExpand = TRUE)
addStyle(wb, "02_Executive_Summary", style_row_label,
         rows = (key_table_row + 3):(key_table_row + 2 + nrow(key_metrics)),
         cols = 2, gridExpand = TRUE)
for (cc in c(3, 4, 7, 9)) {
  addStyle(wb, "02_Executive_Summary", style_pct,
           rows = (key_table_row + 3):(key_table_row + 2 + nrow(key_metrics)),
           cols = cc, gridExpand = TRUE)
}
for (cc in c(5, 6, 8)) {
  addStyle(wb, "02_Executive_Summary", style_num2,
           rows = (key_table_row + 3):(key_table_row + 2 + nrow(key_metrics)),
           cols = cc, gridExpand = TRUE)
}
# Spalten-Bedeutung in key_metrics:
#   Col 3 = Rendite      
#   Col 4 = Volatilität 
#   Col 5 = Sharpe       
#   Col 6 = Sortino      
#   Col 7 = MaxDD        
#   Col 8 = Schiefe      
#   Col 9 = CVaR_95      
# Color-Scale: c(min_color, mid_color, max_color)
#   "höher = besser"  -> c(COL_BAD,  "white", COL_GOOD)
#   "niedriger = besser" -> c(COL_GOOD, "white", COL_BAD)
cf_row <- (key_table_row + 3):(key_table_row + 2 + nrow(key_metrics))
# Rendite (Col 3): höher besser
conditionalFormatting(wb, "02_Executive_Summary", cols = 3, rows = cf_row,
                      type = "colourScale",
                      style = c(COL_BAD, "white", COL_GOOD))
# Volatilität (Col 4): niedriger besser
conditionalFormatting(wb, "02_Executive_Summary", cols = 4, rows = cf_row,
                      type = "colourScale",
                      style = c(COL_GOOD, "white", COL_BAD))
# Sharpe (Col 5): höher besser
conditionalFormatting(wb, "02_Executive_Summary", cols = 5, rows = cf_row,
                      type = "colourScale",
                      style = c(COL_BAD, "white", COL_GOOD))
# Sortino (Col 6): höher besser
conditionalFormatting(wb, "02_Executive_Summary", cols = 6, rows = cf_row,
                      type = "colourScale",
                      style = c(COL_BAD, "white", COL_GOOD))
# MaxDD (Col 7): MaxDD ist POSITIV gespeichert (0.35 = 35% Drawdown).
#   Niedriger = weniger Drawdown = BESSER -> kleiner Wert bekommt grün!
conditionalFormatting(wb, "02_Executive_Summary", cols = 7, rows = cf_row,
                      type = "colourScale",
                      style = c(COL_GOOD, "white", COL_BAD))
# Schiefe (Col 8): höher = besser als OOS-FOMO-Kennzahl
conditionalFormatting(wb, "02_Executive_Summary", cols = 8, rows = cf_row,
                      type = "colourScale",
                      style = c(COL_BAD, "white", COL_GOOD))
# CVaR_95 (Col 9): höher = näher an 0 = besser (Werte sind negativ!)
conditionalFormatting(wb, "02_Executive_Summary", cols = 9, rows = cf_row,
                      type = "colourScale",
                      style = c(COL_BAD, "white", COL_GOOD))

# Modell-Setup aus dem Backtest-Workbook: wichtig, damit die Auswertung klar
# zwischen aktivem Behavioral, Arnott-Robustheit und Hybrid unterscheidet.
model_setup_row <- key_table_row + 3 + nrow(key_metrics) + 3
mergeCells(wb, "02_Executive_Summary", cols = 2:9,
           rows = model_setup_row:model_setup_row)
writeData(wb, "02_Executive_Summary",
          "MODELL- UND PARAMETERSETUP AUS DEM BACKTEST",
          startCol = 2, startRow = model_setup_row)
addStyle(wb, "02_Executive_Summary", style_section,
         rows = model_setup_row, cols = 2:9, gridExpand = TRUE)

model_setup <- data.frame(
  Aspekt = c("MaxSharpe", "Behavioral", "Arnott", "Hybrid",
             "Gemeinsame Constraints", "Behavioral-/Hybrid-Parameter",
             "Begründung Behavioral", "Diskussion Schiefe",
             "Limitation rf", "Datenbasis"),
  Wert = c(bt_model1, bt_model2, bt_arnott, bt_model3,
           bt_constraints, bt_behavioral_params,
           bt_behavioral_reason, bt_skew_discussion,
           bt_rf_limitation, bt_survivorship),
  stringsAsFactors = FALSE
)
writeData(wb, "02_Executive_Summary", model_setup,
          startCol = 2, startRow = model_setup_row + 2, colNames = TRUE)
addStyle(wb, "02_Executive_Summary", style_th,
         rows = model_setup_row + 2, cols = 2:3, gridExpand = TRUE)
addStyle(wb, "02_Executive_Summary", style_row_label,
         rows = (model_setup_row + 3):(model_setup_row + 2 + nrow(model_setup)),
         cols = 2, gridExpand = TRUE)
addStyle(wb, "02_Executive_Summary", style_text,
         rows = (model_setup_row + 3):(model_setup_row + 2 + nrow(model_setup)),
         cols = 3, gridExpand = TRUE)

setColWidths(wb, "02_Executive_Summary", cols = 1, widths = 3)
setColWidths(wb, "02_Executive_Summary", cols = 2, widths = 32)
setColWidths(wb, "02_Executive_Summary", cols = 3:9, widths = 15)


# ==============================================================================
# SHEET 3 — WERTENTWICKLUNG (Equity Curves) MIT EINGEBETTETEM CHART
# ==============================================================================
addWorksheet(wb, "03_Wertentwicklung", gridLines = FALSE,
             tabColour = COL_ACCENT)

mergeCells(wb, "03_Wertentwicklung", cols = 2:11, rows = 2:2)
writeData(wb, "03_Wertentwicklung", "Wertentwicklung im Vergleich",
          startCol = 2, startRow = 2)
addStyle(wb, "03_Wertentwicklung", style_main_title,
         rows = 2, cols = 2:11, gridExpand = TRUE)
setRowHeights(wb, "03_Wertentwicklung", rows = 2, heights = 38)

mergeCells(wb, "03_Wertentwicklung", cols = 2:11, rows = 3:3)
writeData(wb, "03_Wertentwicklung",
          "Wachstum von 1 EUR über den Out-of-Sample-Zeitraum",
          startCol = 2, startRow = 3)
addStyle(wb, "03_Wertentwicklung", style_subtitle, rows = 3, cols = 2:11,
         gridExpand = TRUE)

# Eingebetteter Chart prominent oben
insertImage_safe("03_Wertentwicklung",
                 file.path(PLOTS_DIR, "01_Equity_Curves.png"),
                 startCol = 1, startRow = 5, width = 26, height = 14)

# Tabelle MONATLICHE Equity-Werte unter dem Chart
returns_monthly_idx <- endpoints(returns_xts, on = "months")
returns_monthly <- period.apply(returns_xts, INDEX = returns_monthly_idx,
                                FUN = function(x) apply(x, 2, function(y) {
                                  y[is.na(y)] <- 0
                                  prod(1 + y) - 1
                                }))
equity_xts <- apply(returns_monthly, 2, function(x) cumprod(1 + x))
equity_df_excel <- data.frame(Datum = as.Date(index(returns_monthly)),
                              equity_xts, check.names = FALSE)

mergeCells(wb, "03_Wertentwicklung", cols = 2:11, rows = 35:35)
writeData(wb, "03_Wertentwicklung",
          "MONATLICHE EQUITY-WERTE (Index, Start = 1)",
          startCol = 2, startRow = 35)
addStyle(wb, "03_Wertentwicklung", style_section,
         rows = 35, cols = 2:11, gridExpand = TRUE)

writeData(wb, "03_Wertentwicklung", equity_df_excel,
          startCol = 2, startRow = 37, colNames = TRUE)
addStyle(wb, "03_Wertentwicklung", style_th,
         rows = 37, cols = 2:(1+ncol(equity_df_excel)), gridExpand = TRUE)
addStyle(wb, "03_Wertentwicklung", style_date,
         rows = 38:(37+nrow(equity_df_excel)), cols = 2, gridExpand = TRUE)
addStyle(wb, "03_Wertentwicklung", style_num4,
         rows = 38:(37+nrow(equity_df_excel)),
         cols = 3:(1+ncol(equity_df_excel)), gridExpand = TRUE)

setColWidths(wb, "03_Wertentwicklung", cols = 1, widths = 3)
setColWidths(wb, "03_Wertentwicklung", cols = 2, widths = 13)
setColWidths(wb, "03_Wertentwicklung", cols = 3:(1+ncol(equity_df_excel)),
             widths = 14)
# kein freezePane: Chart oben + Tabelle unten, freezePane würde Layout brechen


# ==============================================================================
# SHEET 4 — RISIKO-PROFIL (Drawdown, Trade-off, Verteilungen)
# ==============================================================================
addWorksheet(wb, "04_Risiko_Profil", gridLines = FALSE, tabColour = COL_ACCENT)

mergeCells(wb, "04_Risiko_Profil", cols = 2:11, rows = 2:2)
writeData(wb, "04_Risiko_Profil", "Risiko-Profil",
          startCol = 2, startRow = 2)
addStyle(wb, "04_Risiko_Profil", style_main_title,
         rows = 2, cols = 2:11, gridExpand = TRUE)
setRowHeights(wb, "04_Risiko_Profil", rows = 2, heights = 38)

mergeCells(wb, "04_Risiko_Profil", cols = 2:11, rows = 3:3)
writeData(wb, "04_Risiko_Profil",
          "Drawdown-Verlauf, Schiefe-Semideviation-Trade-off und Renditeverteilungen",
          startCol = 2, startRow = 3)
addStyle(wb, "04_Risiko_Profil", style_subtitle, rows = 3, cols = 2:11,
         gridExpand = TRUE)

mergeCells(wb, "04_Risiko_Profil", cols = 2:11, rows = 5:5)
writeData(wb, "04_Risiko_Profil",
          "DRAWDOWN-VERLAUF (prozentualer Rückgang vom Höchststand)",
          startCol = 2, startRow = 5)
addStyle(wb, "04_Risiko_Profil", style_section, rows = 5, cols = 2:11,
         gridExpand = TRUE)
insertImage_safe("04_Risiko_Profil",
                 file.path(PLOTS_DIR, "02_Drawdown.png"),
                 startCol = 1, startRow = 7, width = 26, height = 12)

mergeCells(wb, "04_Risiko_Profil", cols = 2:11, rows = 32:32)
writeData(wb, "04_Risiko_Profil",
          "BEHAVIORAL TRADE-OFF: SCHIEFE vs. SEMIDEVIATION",
          startCol = 2, startRow = 32)
addStyle(wb, "04_Risiko_Profil", style_section, rows = 32, cols = 2:11,
         gridExpand = TRUE)
insertImage_safe("04_Risiko_Profil",
                 file.path(PLOTS_DIR, "10_Skew_Semi_Scatter.png"),
                 startCol = 1, startRow = 34, width = 26, height = 14)

mergeCells(wb, "04_Risiko_Profil", cols = 2:11, rows = 64:64)
writeData(wb, "04_Risiko_Profil",
          "VERTEILUNG DER TAGESRENDITEN",
          startCol = 2, startRow = 64)
addStyle(wb, "04_Risiko_Profil", style_section, rows = 64, cols = 2:11,
         gridExpand = TRUE)
insertImage_safe("04_Risiko_Profil",
                 file.path(PLOTS_DIR, "11_Renditeverteilungen.png"),
                 startCol = 1, startRow = 66, width = 26, height = 12)

setColWidths(wb, "04_Risiko_Profil", cols = 1, widths = 3)


# ==============================================================================
# SHEET 5 — VOLLSTÄNDIGE METRIKEN MIT LEGENDE
# ==============================================================================
addWorksheet(wb, "05_Metriken_Detail", gridLines = FALSE)
detail_last_col <- 2 + ncol(metrics_full)

mergeCells(wb, "05_Metriken_Detail", cols = 2:detail_last_col, rows = 2:2)
writeData(wb, "05_Metriken_Detail", "Vollständige Performance-Metriken",
          startCol = 2, startRow = 2)
addStyle(wb, "05_Metriken_Detail", style_main_title,
         rows = 2, cols = 2:detail_last_col, gridExpand = TRUE)
setRowHeights(wb, "05_Metriken_Detail", rows = 2, heights = 38)

mergeCells(wb, "05_Metriken_Detail", cols = 2:detail_last_col, rows = 4:4)
writeData(wb, "05_Metriken_Detail", "LEGENDE DER KENNZAHLEN",
          startCol = 2, startRow = 4)
addStyle(wb, "05_Metriken_Detail", style_section, rows = 4, cols = 2:detail_last_col,
         gridExpand = TRUE)

metric_legend <- data.frame(
  Spalte = c("AnnReturn", "CumReturn", "AnnVol", "AnnVar", "AnnSemiVar",
             "AnnSemiDev", "Skewness", "ExcessKurt", "MaxDD",
             "VaR_95", "CVaR_95", "Sharpe", "Sortino", "Calmar",
             "WinRate", "BestDay", "WorstDay"),
  Bedeutung = c("Annualisierte Rendite (geometrisch, p.a.)",
                "Kumulierte Rendite über den Gesamtzeitraum",
                "Annualisierte Standardabweichung (klassisches Risikomaß)",
                "Annualisierte Varianz",
                "Annualisierte Target-Semivarianz (FOL: echte Verluste unter 0)",
                "Annualisierte Target-Semideviation (Wurzel der Target-Semivarianz)",
                "Schiefe der Renditeverteilung (OOS-FOMO-Kennzahl; höher = günstiger)",
                "Excess Kurtosis (Tail-Risiko)",
                "Maximaler Drawdown vom Höchststand",
                "Value-at-Risk 95 % (täglicher Verlust)",
                "Conditional Value-at-Risk 95 % (Tail-Erwartungswert)",
                "Sharpe-Ratio: (Rendite - rf) / Volatilität",
                "Sortino-Ratio: (Rendite - rf) / Semideviation",
                "Calmar-Ratio: Rendite / Maximum Drawdown",
                "Anteil Tage mit positiver Rendite",
                "Beste Tagesrendite", "Schlechteste Tagesrendite"),
  stringsAsFactors = FALSE
)
writeData(wb, "05_Metriken_Detail", metric_legend, startCol = 2, startRow = 5)
addStyle(wb, "05_Metriken_Detail", style_th,
         rows = 5, cols = 2:3, gridExpand = TRUE)
addStyle(wb, "05_Metriken_Detail", style_row_label,
         rows = 6:(5+nrow(metric_legend)), cols = 2, gridExpand = TRUE)
addStyle(wb, "05_Metriken_Detail", style_text,
         rows = 6:(5+nrow(metric_legend)), cols = 3, gridExpand = TRUE)

mtab_row <- 6 + nrow(metric_legend) + 2
# Banner-Zeile ohne mergeCells (das hat das Tabellen-Rendering blockiert)
writeData(wb, "05_Metriken_Detail", "VOLLSTÄNDIGE METRIKEN",
          startCol = 2, startRow = mtab_row)
addStyle(wb, "05_Metriken_Detail", style_section, rows = mtab_row,
         cols = 2:detail_last_col, gridExpand = TRUE)

mtab <- data.frame(Strategie = rownames(metrics_full), metrics_full,
                   row.names = NULL, check.names = FALSE)
writeData(wb, "05_Metriken_Detail", mtab,
          startCol = 2, startRow = mtab_row + 2)
addStyle(wb, "05_Metriken_Detail", style_th,
         rows = mtab_row + 2, cols = 2:(1+ncol(mtab)), gridExpand = TRUE)
addStyle(wb, "05_Metriken_Detail", style_row_label,
         rows = (mtab_row + 3):(mtab_row + 2 + nrow(mtab)),
         cols = 2, gridExpand = TRUE)
pct_cols_idx <- which(colnames(mtab) %in%
                        c("AnnReturn","CumReturn","AnnVol","AnnSemiDev",
                          "MaxDD","VaR_95","CVaR_95","WinRate","BestDay","WorstDay"))
num_cols_idx <- which(colnames(mtab) %in%
                        c("AnnVar","AnnSemiVar","Skewness","ExcessKurt",
                          "Sharpe","Sortino","Calmar"))
for (cc in pct_cols_idx) {
  addStyle(wb, "05_Metriken_Detail", style_pct,
           rows = (mtab_row + 3):(mtab_row + 2 + nrow(mtab)),
           cols = 1 + cc, gridExpand = TRUE)
}
for (cc in num_cols_idx) {
  addStyle(wb, "05_Metriken_Detail", style_num4,
           rows = (mtab_row + 3):(mtab_row + 2 + nrow(mtab)),
           cols = 1 + cc, gridExpand = TRUE)
}


cf_row_5 <- (mtab_row + 3):(mtab_row + 2 + nrow(mtab))
higher_better <- c("AnnReturn", "CumReturn", "Skewness", "Sharpe", "Sortino",
                   "Calmar", "WinRate", "BestDay",
                   "VaR_95", "CVaR_95", "WorstDay")   # neg. Werte -> höher besser
lower_better  <- c("AnnVol", "AnnVar", "AnnSemiVar", "AnnSemiDev", "ExcessKurt",
                   "MaxDD")                            # MaxDD ist POSITIV -> niedriger besser

for (col_name in higher_better) {
  idx <- which(colnames(mtab) == col_name)
  if (length(idx) > 0) {
    conditionalFormatting(wb, "05_Metriken_Detail", cols = 1 + idx,
                          rows = cf_row_5, type = "colourScale",
                          style = c(COL_BAD, "white", COL_GOOD))
  }
}
for (col_name in lower_better) {
  idx <- which(colnames(mtab) == col_name)
  if (length(idx) > 0) {
    conditionalFormatting(wb, "05_Metriken_Detail", cols = 1 + idx,
                          rows = cf_row_5, type = "colourScale",
                          style = c(COL_GOOD, "white", COL_BAD))
  }
}

# Heatmap-Visualisierung darunter einbetten
heatmap_row <- mtab_row + 3 + nrow(mtab) + 3
writeData(wb, "05_Metriken_Detail",
          "VISUELLER VERGLEICH (z-standardisiert; grün = besser)",
          startCol = 2, startRow = heatmap_row)
addStyle(wb, "05_Metriken_Detail", style_section, rows = heatmap_row,
         cols = 2:detail_last_col, gridExpand = TRUE)
insertImage_safe("05_Metriken_Detail",
                 file.path(PLOTS_DIR, "05_Metrics_Heatmap.png"),
                 startCol = 1, startRow = heatmap_row + 2,
                 width = 30, height = 12)

setColWidths(wb, "05_Metriken_Detail", cols = 1, widths = 3)
setColWidths(wb, "05_Metriken_Detail", cols = 2, widths = 18)
setColWidths(wb, "05_Metriken_Detail", cols = 3:detail_last_col, widths = 14)
# kein freezePane nötig: nur 4 Strategie-Zeilen, alles passt auf einen Bildschirm


# ==============================================================================
# SHEET 6 — JAHRESRENDITEN MIT CHART
# ==============================================================================
addWorksheet(wb, "06_Jahresrenditen", gridLines = FALSE)

mergeCells(wb, "06_Jahresrenditen", cols = 2:11, rows = 2:2)
writeData(wb, "06_Jahresrenditen",
          "Jahresrenditen und Outperformance",
          startCol = 2, startRow = 2)
addStyle(wb, "06_Jahresrenditen", style_main_title,
         rows = 2, cols = 2:11, gridExpand = TRUE)
setRowHeights(wb, "06_Jahresrenditen", rows = 2, heights = 38)

mergeCells(wb, "06_Jahresrenditen", cols = 2:11, rows = 4:4)
writeData(wb, "06_Jahresrenditen", "RENDITEN PRO KALENDERJAHR",
          startCol = 2, startRow = 4)
addStyle(wb, "06_Jahresrenditen", style_section, rows = 4, cols = 2:11,
         gridExpand = TRUE)

ann_perf_out <- data.frame(Jahr = rownames(ann_perf), ann_perf,
                           row.names = NULL, check.names = FALSE)
writeData(wb, "06_Jahresrenditen", ann_perf_out, startCol = 2, startRow = 6)
addStyle(wb, "06_Jahresrenditen", style_th,
         rows = 6, cols = 2:(1+ncol(ann_perf_out)), gridExpand = TRUE)
addStyle(wb, "06_Jahresrenditen", style_row_label,
         rows = 7:(6+nrow(ann_perf_out)), cols = 2, gridExpand = TRUE)
addStyle(wb, "06_Jahresrenditen", style_pct,
         rows = 7:(6+nrow(ann_perf_out)),
         cols = 3:(1+ncol(ann_perf_out)), gridExpand = TRUE)
conditionalFormatting(wb, "06_Jahresrenditen",
                      cols = 3:(1+ncol(ann_perf_out)),
                      rows = 7:(6+nrow(ann_perf_out)),
                      type = "colourScale",
                      style = c(COL_BAD, "white", COL_GOOD))

# Chart rechts daneben
insertImage_safe("06_Jahresrenditen",
                 file.path(PLOTS_DIR, "04_Annual_Returns.png"),
                 startCol = 2 + ncol(ann_perf_out) + 2, startRow = 6,
                 width = 26, height = 14)

# Outperformance-Tabelle unten
outperf_row <- 6 + nrow(ann_perf_out) + 4
mergeCells(wb, "06_Jahresrenditen", cols = 2:11,
           rows = outperf_row:outperf_row)
writeData(wb, "06_Jahresrenditen",
          paste("OUTPERFORMANCE GEGENÜBER", benchmark_col),
          startCol = 2, startRow = outperf_row)
addStyle(wb, "06_Jahresrenditen", style_section,
         rows = outperf_row, cols = 2:11, gridExpand = TRUE)

if (ncol(ann_outperf) > 0) {
  out_df <- data.frame(Jahr = rownames(ann_outperf), ann_outperf,
                       row.names = NULL, check.names = FALSE)
  writeData(wb, "06_Jahresrenditen", out_df,
            startCol = 2, startRow = outperf_row + 2)
  addStyle(wb, "06_Jahresrenditen", style_th,
           rows = outperf_row + 2, cols = 2:(1+ncol(out_df)),
           gridExpand = TRUE)
  addStyle(wb, "06_Jahresrenditen", style_row_label,
           rows = (outperf_row + 3):(outperf_row + 2 + nrow(out_df)),
           cols = 2, gridExpand = TRUE)
  addStyle(wb, "06_Jahresrenditen", style_pct,
           rows = (outperf_row + 3):(outperf_row + 2 + nrow(out_df)),
           cols = 3:(1+ncol(out_df)), gridExpand = TRUE)
  conditionalFormatting(wb, "06_Jahresrenditen",
                        cols = 3:(1+ncol(out_df)),
                        rows = (outperf_row + 3):(outperf_row + 2 + nrow(out_df)),
                        type = "colourScale",
                        style = c(COL_BAD, "white", COL_GOOD))
}

# Outperformance-Plot
op_chart_row <- outperf_row + 3 + nrow(ann_outperf) + 3
mergeCells(wb, "06_Jahresrenditen", cols = 2:11,
           rows = op_chart_row:op_chart_row)
writeData(wb, "06_Jahresrenditen",
          "KUMULIERTE OUTPERFORMANCE-VERLÄUFE",
          startCol = 2, startRow = op_chart_row)
addStyle(wb, "06_Jahresrenditen", style_section,
         rows = op_chart_row, cols = 2:11, gridExpand = TRUE)
insertImage_safe("06_Jahresrenditen",
                 file.path(PLOTS_DIR, "13_Outperformance.png"),
                 startCol = 1, startRow = op_chart_row + 2,
                 width = 28, height = 12)

setColWidths(wb, "06_Jahresrenditen", cols = 1, widths = 3)
setColWidths(wb, "06_Jahresrenditen", cols = 2, widths = 10)
setColWidths(wb, "06_Jahresrenditen", cols = 3:(1+ncol(ann_perf_out)),
             widths = 14)


# ==============================================================================
# SHEET 7 — KRISE & BOOM
# ==============================================================================
addWorksheet(wb, "07_Krise_Boom", gridLines = FALSE)

mergeCells(wb, "07_Krise_Boom", cols = 2:11, rows = 2:2)
writeData(wb, "07_Krise_Boom",
          "Performance in Stress- und Boom-Phasen",
          startCol = 2, startRow = 2)
addStyle(wb, "07_Krise_Boom", style_main_title,
         rows = 2, cols = 2:11, gridExpand = TRUE)
setRowHeights(wb, "07_Krise_Boom", rows = 2, heights = 38)

mergeCells(wb, "07_Krise_Boom", cols = 2:11, rows = 3:3)
writeData(wb, "07_Krise_Boom",
          paste0("Corona-Stressphase (", CRISIS_START, " bis ", CRISIS_END,
                 ") und Tech-Rallye 2023"),
          startCol = 2, startRow = 3)
addStyle(wb, "07_Krise_Boom", style_subtitle, rows = 3, cols = 2:11,
         gridExpand = TRUE)

current_row <- 5

write_subperiod_styled <- function(sheet, m, title, start_row) {
  if (is.null(m)) return(start_row)
  mergeCells(wb, sheet, cols = 2:11, rows = start_row:start_row)
  writeData(wb, sheet, title, startCol = 2, startRow = start_row)
  addStyle(wb, sheet, style_section, rows = start_row, cols = 2:11,
           gridExpand = TRUE)
  
  d <- data.frame(Strategie = rownames(m), m, row.names = NULL,
                  check.names = FALSE)
  writeData(wb, sheet, d, startCol = 2, startRow = start_row + 2)
  addStyle(wb, sheet, style_th,
           rows = start_row + 2, cols = 2:(1+ncol(d)), gridExpand = TRUE)
  addStyle(wb, sheet, style_row_label,
           rows = (start_row + 3):(start_row + 2 + nrow(d)),
           cols = 2, gridExpand = TRUE)
  pcols <- which(colnames(d) %in% c("AnnReturn","CumReturn","AnnVol",
                                    "AnnSemiDev","MaxDD","VaR_95","CVaR_95",
                                    "WinRate","BestDay","WorstDay"))
  for (cc in pcols) {
    addStyle(wb, sheet, style_pct,
             rows = (start_row + 3):(start_row + 2 + nrow(d)),
             cols = 1 + cc, gridExpand = TRUE)
  }
  ncols_idx <- which(colnames(d) %in% c("AnnVar","AnnSemiVar","Skewness",
                                        "ExcessKurt","Sharpe","Sortino",
                                        "Calmar"))
  for (cc in ncols_idx) {
    addStyle(wb, sheet, style_num4,
             rows = (start_row + 3):(start_row + 2 + nrow(d)),
             cols = 1 + cc, gridExpand = TRUE)
  }
  cum_idx <- which(colnames(d) == "CumReturn")
  if (length(cum_idx) > 0) {
    conditionalFormatting(wb, sheet, cols = 1 + cum_idx,
                          rows = (start_row + 3):(start_row + 2 + nrow(d)),
                          type = "colourScale",
                          style = c(COL_BAD, "white", COL_GOOD))
  }
  higher_better_sub <- c("AnnReturn", "CumReturn", "Skewness", "Sharpe",
                         "Sortino", "Calmar", "WinRate", "BestDay",
                         "VaR_95", "CVaR_95", "WorstDay")
  lower_better_sub  <- c("AnnVol", "AnnVar", "AnnSemiVar", "AnnSemiDev",
                         "ExcessKurt", "MaxDD")
  for (col_name in higher_better_sub) {
    idx <- which(colnames(d) == col_name)
    if (length(idx) > 0) {
      conditionalFormatting(wb, sheet, cols = 1 + idx,
                            rows = (start_row + 3):(start_row + 2 + nrow(d)),
                            type = "colourScale",
                            style = c(COL_BAD, "white", COL_GOOD))
    }
  }
  for (col_name in lower_better_sub) {
    idx <- which(colnames(d) == col_name)
    if (length(idx) > 0) {
      conditionalFormatting(wb, sheet, cols = 1 + idx,
                            rows = (start_row + 3):(start_row + 2 + nrow(d)),
                            type = "colourScale",
                            style = c(COL_GOOD, "white", COL_BAD))
    }
  }
  return(start_row + 3 + nrow(d) + 2)
}

current_row <- write_subperiod_styled("07_Krise_Boom", metrics_crisis,
                                      "CORONA-STRESSPHASE (Februar - Juni 2020)",
                                      current_row)
current_row <- write_subperiod_styled("07_Krise_Boom", metrics_boom,
                                      "TECH-RALLYE 2023",
                                      current_row)

# Charts
mergeCells(wb, "07_Krise_Boom", cols = 2:11,
           rows = (current_row + 1):(current_row + 1))
writeData(wb, "07_Krise_Boom",
          "WERTENTWICKLUNG UND KUMULIERTE RENDITEN",
          startCol = 2, startRow = current_row + 1)
addStyle(wb, "07_Krise_Boom", style_section,
         rows = current_row + 1, cols = 2:11, gridExpand = TRUE)

insertImage_safe("07_Krise_Boom",
                 file.path(PLOTS_DIR, "06_Krise_Boom.png"),
                 startCol = 1, startRow = current_row + 3,
                 width = 28, height = 13)
insertImage_safe("07_Krise_Boom",
                 file.path(PLOTS_DIR, "15_Krise_Boom_Bars.png"),
                 startCol = 1, startRow = current_row + 32,
                 width = 28, height = 13)

setColWidths(wb, "07_Krise_Boom", cols = 1, widths = 3)
setColWidths(wb, "07_Krise_Boom", cols = 2, widths = 18)
setColWidths(wb, "07_Krise_Boom", cols = 3:18, widths = 14)


# ==============================================================================
# SHEET 8 — RISIKOADJUSTIERTE RENDITEN (Sharpe vs Sortino + Korrelation)
# ==============================================================================
addWorksheet(wb, "08_Risiko_Adjustiert", gridLines = FALSE,
             tabColour = COL_ACCENT)

mergeCells(wb, "08_Risiko_Adjustiert", cols = 2:10, rows = 2:2)
writeData(wb, "08_Risiko_Adjustiert",
          "Risikoadjustierte Renditemaße",
          startCol = 2, startRow = 2)
addStyle(wb, "08_Risiko_Adjustiert", style_main_title,
         rows = 2, cols = 2:10, gridExpand = TRUE)
setRowHeights(wb, "08_Risiko_Adjustiert", rows = 2, heights = 38)

mergeCells(wb, "08_Risiko_Adjustiert", cols = 2:10, rows = 4:4)
writeData(wb, "08_Risiko_Adjustiert",
          "SHARPE vs. SORTINO IM VERGLEICH",
          startCol = 2, startRow = 4)
addStyle(wb, "08_Risiko_Adjustiert", style_section,
         rows = 4, cols = 2:10, gridExpand = TRUE)
insertImage_safe("08_Risiko_Adjustiert",
                 file.path(PLOTS_DIR, "12_Sharpe_vs_Sortino.png"),
                 startCol = 1, startRow = 6,
                 width = 26, height = 13)

mergeCells(wb, "08_Risiko_Adjustiert", cols = 2:10, rows = 32:32)
writeData(wb, "08_Risiko_Adjustiert",
          "ROLLIERENDE 60-TAGE-VOLATILITÄT",
          startCol = 2, startRow = 32)
addStyle(wb, "08_Risiko_Adjustiert", style_section,
         rows = 32, cols = 2:10, gridExpand = TRUE)
insertImage_safe("08_Risiko_Adjustiert",
                 file.path(PLOTS_DIR, "03_Rolling_Volatility.png"),
                 startCol = 1, startRow = 34,
                 width = 26, height = 12)

mergeCells(wb, "08_Risiko_Adjustiert", cols = 2:10, rows = 60:60)
writeData(wb, "08_Risiko_Adjustiert",
          "KORRELATIONSMATRIX DER STRATEGIE-RENDITEN",
          startCol = 2, startRow = 60)
addStyle(wb, "08_Risiko_Adjustiert", style_section,
         rows = 60, cols = 2:10, gridExpand = TRUE)
insertImage_safe("08_Risiko_Adjustiert",
                 file.path(PLOTS_DIR, "14_Korrelationsmatrix.png"),
                 startCol = 1, startRow = 62,
                 width = 22, height = 13)

setColWidths(wb, "08_Risiko_Adjustiert", cols = 1, widths = 3)


# ==============================================================================
# SHEET 9 — KONZENTRATIONSANALYSE
# ==============================================================================
addWorksheet(wb, "09_Konzentration", gridLines = FALSE)

mergeCells(wb, "09_Konzentration", cols = 2:10, rows = 2:2)
writeData(wb, "09_Konzentration", "Portfolio-Konzentration",
          startCol = 2, startRow = 2)
addStyle(wb, "09_Konzentration", style_main_title,
         rows = 2, cols = 2:10, gridExpand = TRUE)
setRowHeights(wb, "09_Konzentration", rows = 2, heights = 38)

mergeCells(wb, "09_Konzentration", cols = 2:10, rows = 3:3)
writeData(wb, "09_Konzentration",
          "HHI, effektive Positionen, Top-10-Anteil pro Rebalancing",
          startCol = 2, startRow = 3)
addStyle(wb, "09_Konzentration", style_subtitle,
         rows = 3, cols = 2:10, gridExpand = TRUE)

if (length(concentration_list) > 0) {
  conc_all <- bind_rows(lapply(names(concentration_list), function(strat) {
    df <- concentration_list[[strat]]
    if (is.null(df)) return(NULL)
    df$Strategy <- strat
    df[, c("Strategy", setdiff(colnames(df), "Strategy"))]
  }))
  
  mergeCells(wb, "09_Konzentration", cols = 2:10, rows = 5:5)
  writeData(wb, "09_Konzentration", "KONZENTRATIONSMETRIKEN",
            startCol = 2, startRow = 5)
  addStyle(wb, "09_Konzentration", style_section,
           rows = 5, cols = 2:10, gridExpand = TRUE)
  
  writeData(wb, "09_Konzentration", conc_all,
            startCol = 2, startRow = 7)
  addStyle(wb, "09_Konzentration", style_th,
           rows = 7, cols = 2:(1+ncol(conc_all)), gridExpand = TRUE)
  addStyle(wb, "09_Konzentration", style_row_label,
           rows = 8:(7+nrow(conc_all)), cols = 2:3, gridExpand = TRUE)
  pcols <- which(colnames(conc_all) %in% c("HHI","MaxWeight","Top10Share"))
  for (cc in pcols) {
    addStyle(wb, "09_Konzentration", style_pct,
             rows = 8:(7+nrow(conc_all)), cols = 1 + cc, gridExpand = TRUE)
  }
  ncols_conc <- which(colnames(conc_all) %in% c("EffN"))
  for (cc in ncols_conc) {
    addStyle(wb, "09_Konzentration", style_num2,
             rows = 8:(7+nrow(conc_all)), cols = 1 + cc, gridExpand = TRUE)
  }
}

# Fallback falls conc_all nicht existiert oder leer ist
if (!exists("conc_all") || is.null(conc_all) || nrow(conc_all) == 0) {
  conc_all <- data.frame(Strategy = character(0))
}

chart_start <- max(10, 7 + nrow(conc_all) + 3)
mergeCells(wb, "09_Konzentration", cols = 2:10,
           rows = chart_start:chart_start)
writeData(wb, "09_Konzentration", "EFFEKTIVE POSITIONEN ÜBER ZEIT",
          startCol = 2, startRow = chart_start)
addStyle(wb, "09_Konzentration", style_section,
         rows = chart_start, cols = 2:10, gridExpand = TRUE)
insertImage_safe("09_Konzentration",
                 file.path(PLOTS_DIR, "07_Konzentration_EffN.png"),
                 startCol = 1, startRow = chart_start + 2,
                 width = 24, height = 12)

mergeCells(wb, "09_Konzentration", cols = 2:10,
           rows = (chart_start + 28):(chart_start + 28))
writeData(wb, "09_Konzentration", "ANZAHL AKTIVER POSITIONEN PRO REBALANCING",
          startCol = 2, startRow = chart_start + 28)
addStyle(wb, "09_Konzentration", style_section,
         rows = chart_start + 28, cols = 2:10, gridExpand = TRUE)
insertImage_safe("09_Konzentration",
                 file.path(PLOTS_DIR, "08_Konzentration_AnzPos.png"),
                 startCol = 1, startRow = chart_start + 30,
                 width = 24, height = 12)

setColWidths(wb, "09_Konzentration", cols = 1, widths = 3)
setColWidths(wb, "09_Konzentration", cols = 2, widths = 16)
setColWidths(wb, "09_Konzentration", cols = 3:9, widths = 13)


# ==============================================================================
# SHEET 10 — TOP HOLDINGS
# ==============================================================================
addWorksheet(wb, "10_Top_Holdings", gridLines = FALSE)

mergeCells(wb, "10_Top_Holdings", cols = 2:10, rows = 2:2)
writeData(wb, "10_Top_Holdings",
          "Top-10-Positionen beim letzten Rebalancing",
          startCol = 2, startRow = 2)
addStyle(wb, "10_Top_Holdings", style_main_title,
         rows = 2, cols = 2:10, gridExpand = TRUE)
setRowHeights(wb, "10_Top_Holdings", rows = 2, heights = 38)

if (length(top10_list) > 0) {
  top10_all <- bind_rows(lapply(names(top10_list), function(strat) {
    df <- top10_list[[strat]]
    if (is.null(df)) return(NULL)
    df$Strategy <- strat
    df[, c("Strategy","Year","Rank","Ticker","Weight")]
  }))
  if (nrow(top10_all) > 0) {
    writeData(wb, "10_Top_Holdings", top10_all,
              startCol = 2, startRow = 5)
    addStyle(wb, "10_Top_Holdings", style_th,
             rows = 5, cols = 2:(1+ncol(top10_all)), gridExpand = TRUE)
    addStyle(wb, "10_Top_Holdings", style_row_label,
             rows = 6:(5+nrow(top10_all)), cols = 2, gridExpand = TRUE)
    addStyle(wb, "10_Top_Holdings", style_pct,
             rows = 6:(5+nrow(top10_all)),
             cols = 1 + which(colnames(top10_all) == "Weight"),
             gridExpand = TRUE)
    setColWidths(wb, "10_Top_Holdings", cols = 1, widths = 3)
    setColWidths(wb, "10_Top_Holdings", cols = 2:6, widths = 14)
  }
}

insertImage_safe("10_Top_Holdings",
                 file.path(PLOTS_DIR, "09_Top10_Holdings.png"),
                 startCol = 8, startRow = 5,
                 width = 22, height = 13)


# ==============================================================================
# SHEET 11 — CHART-GALERIE (alle Visualisierungen auf einen Blick)
# ==============================================================================
addWorksheet(wb, "11_Charts_Galerie", gridLines = FALSE,
             tabColour = COL_GOOD)

mergeCells(wb, "11_Charts_Galerie", cols = 2:10, rows = 2:2)
writeData(wb, "11_Charts_Galerie", "Galerie aller Visualisierungen",
          startCol = 2, startRow = 2)
addStyle(wb, "11_Charts_Galerie", style_main_title,
         rows = 2, cols = 2:10, gridExpand = TRUE)
setRowHeights(wb, "11_Charts_Galerie", rows = 2, heights = 38)

chart_files <- list(
  list(file = "01_Equity_Curves.png",       label = "Wertentwicklung"),
  list(file = "02_Drawdown.png",            label = "Drawdown-Verlauf"),
  list(file = "03_Rolling_Volatility.png",  label = "Rollierende Volatilität"),
  list(file = "04_Annual_Returns.png",      label = "Jahresrenditen"),
  list(file = "05_Metrics_Heatmap.png",     label = "Metriken-Heatmap"),
  list(file = "06_Krise_Boom.png",          label = "Krise vs. Boom (Equity)"),
  list(file = "07_Konzentration_EffN.png",  label = "Effektive Positionen"),
  list(file = "08_Konzentration_AnzPos.png",label = "Anzahl Positionen"),
  list(file = "09_Top10_Holdings.png",      label = "Top-10 Holdings"),
  list(file = "10_Skew_Semi_Scatter.png",   label = "Schiefe vs. Semideviation"),
  list(file = "11_Renditeverteilungen.png", label = "Renditeverteilungen"),
  list(file = "12_Sharpe_vs_Sortino.png",   label = "Sharpe vs. Sortino"),
  list(file = "13_Outperformance.png",      label = "Outperformance"),
  list(file = "14_Korrelationsmatrix.png",  label = "Korrelationsmatrix"),
  list(file = "15_Krise_Boom_Bars.png",     label = "Krise vs. Boom (Bars)")
)

chart_row <- 5
for (i in seq_along(chart_files)) {
  ch <- chart_files[[i]]
  mergeCells(wb, "11_Charts_Galerie", cols = 2:10,
             rows = chart_row:chart_row)
  writeData(wb, "11_Charts_Galerie",
            paste0("Chart ", i, ": ", ch$label),
            startCol = 2, startRow = chart_row)
  addStyle(wb, "11_Charts_Galerie", style_section,
           rows = chart_row, cols = 2:10, gridExpand = TRUE)
  insertImage_safe("11_Charts_Galerie",
                   file.path(PLOTS_DIR, ch$file),
                   startCol = 1, startRow = chart_row + 2,
                   width = 24, height = 13)
  chart_row <- chart_row + 30
}

setColWidths(wb, "11_Charts_Galerie", cols = 1, widths = 3)


# ==============================================================================
# SHEET 12 — ROHDATEN (OOS Returns)
# ==============================================================================
addWorksheet(wb, "12_Rohdaten_Returns", gridLines = TRUE,
             tabColour = "#95A5A6")

mergeCells(wb, "12_Rohdaten_Returns", cols = 1:ncol(returns_df), rows = 1:1)
writeData(wb, "12_Rohdaten_Returns",
          "Rohdaten: Out-of-Sample Tagesrenditen",
          startCol = 1, startRow = 1)
addStyle(wb, "12_Rohdaten_Returns", style_main_title,
         rows = 1, cols = 1:ncol(returns_df), gridExpand = TRUE)
setRowHeights(wb, "12_Rohdaten_Returns", rows = 1, heights = 30)

writeData(wb, "12_Rohdaten_Returns", returns_df, startRow = 3)
addStyle(wb, "12_Rohdaten_Returns", style_th,
         rows = 3, cols = 1:ncol(returns_df), gridExpand = TRUE)
addStyle(wb, "12_Rohdaten_Returns", style_date,
         rows = 4:(nrow(returns_df)+3), cols = 1, gridExpand = TRUE)
addStyle(wb, "12_Rohdaten_Returns", style_pct,
         rows = 4:(nrow(returns_df)+3),
         cols = 2:ncol(returns_df), gridExpand = TRUE)
setColWidths(wb, "12_Rohdaten_Returns", cols = 1:ncol(returns_df),
             widths = 13)
freezePane(wb, "12_Rohdaten_Returns", firstActiveRow = 4, firstActiveCol = 2)


# ==============================================================================
# SHEET 13-16 — GEWICHTE-MATRIZEN PRO STRATEGIE
# ==============================================================================

write_weights_matrix <- function(sheet_name, W, strategy_label) {
  if (is.null(W) || nrow(W) == 0) return(invisible())
  addWorksheet(wb, sheet_name, gridLines = TRUE, tabColour = "#95A5A6")
  
  n_cols_data <- ncol(W) + 1  # +1 für Datums-Spalte
  mergeCells(wb, sheet_name, cols = 1:min(n_cols_data, 20), rows = 1:1)
  writeData(wb, sheet_name,
            paste("Portfolio-Gewichte:", strategy_label),
            startCol = 1, startRow = 1)
  addStyle(wb, sheet_name, style_main_title, rows = 1,
           cols = 1:min(n_cols_data, 20), gridExpand = TRUE)
  setRowHeights(wb, sheet_name, rows = 1, heights = 30)
  
  # Subtitle
  mergeCells(wb, sheet_name, cols = 1:min(n_cols_data, 20), rows = 2:2)
  writeData(wb, sheet_name,
            paste0(nrow(W), " Rebalancings x ", ncol(W),
                   " Aktien (Werte in Prozent des Portfolios)"),
            startCol = 1, startRow = 2)
  addStyle(wb, sheet_name, style_subtitle, rows = 2,
           cols = 1:min(n_cols_data, 20), gridExpand = TRUE)
  
  # Header-Zeile + Daten zusammenbauen
  W_df <- data.frame(Rebalancing = rownames(W), W,
                     check.names = FALSE, stringsAsFactors = FALSE)
  writeData(wb, sheet_name, W_df, startRow = 4)
  addStyle(wb, sheet_name, style_th,
           rows = 4, cols = 1:n_cols_data, gridExpand = TRUE)
  addStyle(wb, sheet_name, style_row_label,
           rows = 5:(4+nrow(W)), cols = 1, gridExpand = TRUE)
  addStyle(wb, sheet_name, style_pct,
           rows = 5:(4+nrow(W)), cols = 2:n_cols_data, gridExpand = TRUE)
  
  setColWidths(wb, sheet_name, cols = 1, widths = 14)
  setColWidths(wb, sheet_name, cols = 2:n_cols_data, widths = 11)
  freezePane(wb, sheet_name, firstActiveRow = 5, firstActiveCol = 2)
}

tryCatch(
  write_weights_matrix("13_Gewichte_MaxSharpe", W_sharpe,
                       "Mean-Variance / MaxSharpe"),
  error = function(e) cat("  WARNUNG: Gewichte_MaxSharpe:", e$message, "\n")
)
tryCatch(
  write_weights_matrix("14_Gewichte_Behavioral", W_behav,
                       "Behavioral (aktive FOMO/FOL-Ratio)"),
  error = function(e) cat("  WARNUNG: Gewichte_Behavioral:", e$message, "\n")
)
tryCatch(
  write_weights_matrix("15_Gewichte_Arnott", W_arnott,
                       "Arnott-Blend (Robustheitsmodell)"),
  error = function(e) cat("  WARNUNG: Gewichte_Arnott:", e$message, "\n")
)
if (!is.null(W_hybrid)) {
  tryCatch(
    write_weights_matrix("16_Gewichte_Hybrid", W_hybrid,
                         "Hybrid-Blend (MaxSharpe + Behavioral-FOMO/FOL)"),
    error = function(e) cat("  WARNUNG: Gewichte_Hybrid:", e$message, "\n")
  )
}


# Sicherheits-Check: alle Sheets angelegt?
cat("  -> Sheets im Workbook:", paste(names(wb), collapse=", "), "\n")
cat("  -> Speichere Workbook ...\n")
saveWorkbook(wb, OUTPUT_XLSX, overwrite = TRUE)
cat("  -> Datei erfolgreich geschrieben!\n")


cat("\n================================================================\n")
cat("FERTIG.\n")
cat("Excel:  ", normalizePath(OUTPUT_XLSX,  mustWork = FALSE), "\n")
cat("Plots:  ", normalizePath(PLOTS_DIR,    mustWork = FALSE), "\n")
cat("PDF:    ", normalizePath(COMBINED_PDF, mustWork = FALSE), "\n")
cat("================================================================\n")
