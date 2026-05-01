# ═══════════════════════════════════════════════════════════════════════
#  graphic.R — Standalone Chart Generator (R version, procedural)
# ═══════════════════════════════════════════════════════════════════════
#
#  Reads the data_grafics_*.xlsx produced by 001_report.py
#  and recreates all 16 publication-ready charts.
#
#  NO function definitions — all code is inline / procedural.
#
#  Usage:
#    Rscript report/graphic.R
#
#  Output:
#    output/figures_custom_R/*.png   (16 charts, 300 DPI)
#
#  Customize freely: change colors, labels, sizes, themes, etc.
# ═══════════════════════════════════════════════════════════════════════


# ═══════════════════════════════════════════════════════════════════════
# Libraries
# ═══════════════════════════════════════════════════════════════════════

library(pacman)
p_load(tidyverse, readxl, scales, ggrepel, patchwork, fmsb)


# ═══════════════════════════════════════════════════════════════════════
# Configuration
# ═══════════════════════════════════════════════════════════════════════

MODEL_COLORS <- c(
  "gpt-4o"     = "#4472C4",
  "gpt_4o"     = "#4472C4",
  "gpt-5_2"    = "#ED7D31",
  "gpt_5_2"    = "#ED7D31",
  "gpt-5-mini" = "#A5A5A5",
  "gpt_5_mini" = "#A5A5A5",
  "gpt-5-nano" = "#FFC000",
  "gpt_5_nano" = "#FFC000"
)

DEFAULT_COLORS <- c("#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000", "#70AD47",
                    "#9B59B6", "#E74C3C", "#1ABC9C", "#34495E", "#F39C12")

DPI <- 300

# Inline theme elements (used directly in each chart instead of theme_chart())
BASE_THEME <- theme_minimal(base_size = 10) +
  theme(
    plot.title       = element_text(face = "bold", size = 11),
    axis.title       = element_text(size = 9),
    legend.position  = "bottom",
    panel.grid.minor = element_blank()
  )


# ═══════════════════════════════════════════════════════════════════════
# Paths
# ═══════════════════════════════════════════════════════════════════════

script_dir <- tryCatch({
  arg <- grep("--file=", commandArgs(FALSE), value = TRUE)[1]
  dirname(normalizePath(sub("--file=", "", arg)))
}, error = function(e) getwd())

project_dir <- dirname(script_dir)
output_dir  <- file.path(project_dir, "output")

xlsx_path <- sort(
  list.files(output_dir, pattern = "^data_grafics_.*\\.xlsx$", full.names = TRUE),
  decreasing = TRUE
)[1]

if (is.na(xlsx_path)) stop(paste0(
  "No data_grafics_*.xlsx found in ", output_dir,
  ". Run relatorio_unificado.py first."
))

out_dir <- file.path(output_dir, "figures_custom_R")
dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)

cat("  Script dir:", script_dir, "\n")
cat("  Reading:   ", basename(xlsx_path), "\n")
cat("  Output:    ", out_dir, "\n\n")


# ═══════════════════════════════════════════════════════════════════════
# Graphic 1 — Sensitivity per Model per Project
# ═══════════════════════════════════════════════════════════════════════

df_sensitivity <- read_excel(xlsx_path, sheet = "sensitivity_per_model")

agg_sensitivity <- df_sensitivity %>%
  group_by(Project, Model) %>%
  summarise(Mean = mean(Sensitivity, na.rm = TRUE) * 100, .groups = "drop")

# Build color scale inline for sensitivity models
sens_models <- unique(df_sensitivity$Model)
sens_colors <- setNames(
  sapply(seq_along(sens_models), function(i) {
    key <- gsub("-", "_", tolower(trimws(sens_models[i])))
    if (key %in% names(MODEL_COLORS)) MODEL_COLORS[[key]]
    else DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1]
  }),
  sens_models
)

g01 <- ggplot(agg_sensitivity, aes(x = Model, y = Mean, fill = Model)) +
  geom_col(alpha = 0.85, colour = "#333333", linewidth = 0.3) +
  geom_hline(yintercept = 95, linetype = "dashed", colour = "#27AE60", alpha = 0.6) +
  geom_hline(yintercept = 80, linetype = "dotted", colour = "#E74C3C", alpha = 0.5) +
  geom_point(
    data = df_sensitivity %>% mutate(Sens_pct = Sensitivity * 100),
    aes(x = Model, y = Sens_pct), inherit.aes = FALSE,
    size = 1.5, colour = "#222222", alpha = 0.7
  ) +
  geom_text(aes(label = sprintf("%.0f%%", Mean)), vjust = -0.5, size = 2.5, fontface = "bold") +
  facet_wrap(~ Project, scales = "free_x") +
  scale_fill_manual(values = sens_colors) +
  labs(title = "Sensitivity by Model per Project (vs Human TIAB)",
       y = "Sensitivity (%)", x = NULL) +
  ylim(0, 108) +
  BASE_THEME +
  theme(axis.text.x = element_text(angle = 25, hjust = 1, size = 7),
        legend.position = "none")

g01


# ═══════════════════════════════════════════════════════════════════════
# Graphic 2 — Listfinal Capture Heatmap
# ═══════════════════════════════════════════════════════════════════════

df_lf_capture <- read_excel(xlsx_path, sheet = "lf_capture_heatmap")

g02 <- ggplot(df_lf_capture, aes(x = Project, y = Model, fill = Capture_Rate_pct)) +
  geom_tile(colour = "white", linewidth = 1.5) +
  geom_text(
    aes(label = ifelse(is.na(Capture_Rate_pct), "\u2014", sprintf("%.1f%%", Capture_Rate_pct))),
    fontface = "bold", size = 3.5
  ) +
  scale_fill_gradient2(
    low = "#E74C3C", mid = "#FFC000", high = "#27AE60",
    midpoint = 85, na.value = "grey90",
    limits = c(60, 100), name = "Capture Rate (%)"
  ) +
  labs(title = "Listfinal Capture Rate by Model and Project\n(Average across Runs)",
       x = NULL, y = NULL) +
  BASE_THEME +
  theme(legend.position = "right")

g02


# ═══════════════════════════════════════════════════════════════════════
# Graphic 3 — Test-Retest Kappa
# ═══════════════════════════════════════════════════════════════════════

df_kappa <- read_excel(xlsx_path, sheet = "test_retest_kappa")
df_kappa$Label <- factor(df_kappa$Label, levels = df_kappa$Label)
df_kappa$ModelName <- sapply(as.character(df_kappa$Label),
                             function(x) strsplit(x, "\n")[[1]][1])

kappa_models <- unique(df_kappa$ModelName)
kappa_colors <- setNames(
  sapply(seq_along(kappa_models), function(i) {
    key <- gsub("-", "_", tolower(trimws(kappa_models[i])))
    if (key %in% names(MODEL_COLORS)) MODEL_COLORS[[key]]
    else DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1]
  }),
  kappa_models
)

g03 <- ggplot(df_kappa, aes(x = Label, y = Kappa, fill = ModelName)) +
  geom_col(alpha = 0.85, colour = "#333333", linewidth = 0.3) +
  geom_errorbar(aes(ymin = CI_lo, ymax = CI_hi), width = 0.25, linewidth = 0.8, colour = "#444444") +
  geom_hline(yintercept = 0.81, linetype = "dashed", colour = "#27AE60", alpha = 0.6) +
  geom_hline(yintercept = 0.61, linetype = "dashed", colour = "#F39C12", alpha = 0.5) +
  geom_text(
    aes(label = sprintf("%.3f", Kappa), y = pmin(CI_hi, 1.0) + 0.02),
    vjust = 0, size = 2.5, fontface = "bold"
  ) +
  scale_fill_manual(values = kappa_colors) +
  labs(title = "Test-Retest Reproducibility (Cohen's Kappa with 95% CI)",
       y = "Kappa", x = NULL) +
  ylim(0, 1.18) +
  BASE_THEME +
  theme(axis.text.x = element_text(size = 7), legend.position = "none")

g03


# ═══════════════════════════════════════════════════════════════════════
# Graphic 4 — Radar Chart (Model Comparison)
# ═══════════════════════════════════════════════════════════════════════

df_radar <- read_excel(xlsx_path, sheet = "model_comparison_radar")

metric_cols <- setdiff(names(df_radar), "Model")
max_row     <- rep(1.05, length(metric_cols))
min_row     <- rep(0,    length(metric_cols))
radar_data  <- rbind(max_row, min_row, as.data.frame(df_radar[, metric_cols]))
colnames(radar_data) <- metric_cols

radar_colors <- sapply(seq_len(nrow(df_radar)), function(i) {
  key <- gsub("-", "_", tolower(trimws(df_radar$Model[i])))
  if (key %in% names(MODEL_COLORS)) MODEL_COLORS[[key]]
  else DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1]
})


# ═══════════════════════════════════════════════════════════════════════
# Graphic 5 — Cost vs Sensitivity (Bubble)
# ═══════════════════════════════════════════════════════════════════════

df_cost_sens <- read_excel(xlsx_path, sheet = "cost_vs_sensitivity")

cost_models <- unique(df_cost_sens$Model)
cost_colors <- setNames(
  sapply(seq_along(cost_models), function(i) {
    key <- gsub("-", "_", tolower(trimws(cost_models[i])))
    if (key %in% names(MODEL_COLORS)) MODEL_COLORS[[key]]
    else DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1]
  }),
  cost_models
)

g05 <- ggplot(df_cost_sens, aes(x = Avg_Cost_USD, y = Avg_Sensitivity_pct)) +
  geom_point(aes(size = Avg_F1, colour = Model), alpha = 0.75, stroke = 1) +
  geom_hline(yintercept = 95, linetype = "dashed", colour = "#27AE60", alpha = 0.5) +
  geom_text_repel(aes(label = Model), size = 3, fontface = "bold", segment.colour = "grey60") +
  scale_colour_manual(values = cost_colors) +
  scale_size_continuous(range = c(4, 15), name = "F1 Score") +
  labs(title = "Cost vs Sensitivity (Bubble Size = F1 Score)",
       x = "Average Cost (USD)", y = "Average Sensitivity (%)") +
  ylim(0, 108) +
  BASE_THEME

g05


# ═══════════════════════════════════════════════════════════════════════
# Graphic 6 — Workload Reduction
# ═══════════════════════════════════════════════════════════════════════

df_workload <- read_excel(xlsx_path, sheet = "workload_reduction")

df_work_long <- df_workload %>%
  select(Model, Human_Hours, AI_Hours) %>%
  pivot_longer(cols = c(Human_Hours, AI_Hours), names_to = "Type", values_to = "Hours") %>%
  mutate(
    Type = ifelse(Type == "Human_Hours", "Human", "AI"),
    Type = factor(Type, levels = c("Human", "AI"))
  )

g06a <- ggplot(df_work_long, aes(x = Hours, y = reorder(Model, Hours), fill = Type)) +
  geom_col(position = position_dodge(width = 0.7), width = 0.6,
           alpha = 0.85, colour = "#333333", linewidth = 0.3) +
  scale_fill_manual(values = c("Human" = "#E74C3C", "AI" = "#2ECC71")) +
  labs(title = "Average Screening Time: Human vs AI",
       x = "Time (hours)", y = NULL) +
  BASE_THEME

work_models <- unique(df_workload$Model)
work_colors <- setNames(
  sapply(seq_along(work_models), function(i) {
    key <- gsub("-", "_", tolower(trimws(work_models[i])))
    if (key %in% names(MODEL_COLORS)) MODEL_COLORS[[key]]
    else DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1]
  }),
  work_models
)

g06b <- ggplot(df_workload, aes(x = Speed_Factor, y = reorder(Model, Speed_Factor))) +
  geom_col(aes(fill = Model), alpha = 0.85, colour = "#333333", linewidth = 0.3,
           width = 0.5, show.legend = FALSE) +
  scale_fill_manual(values = work_colors) +
  geom_text(aes(label = sprintf("%.0f\u00d7", Speed_Factor)),
            hjust = -0.3, fontface = "bold", size = 3.5) +
  labs(title = "Speed Factor (Human \u00f7 AI)", x = "Speed Factor (\u00d7)", y = NULL) +
  BASE_THEME

g06 <- g06a + g06b + plot_layout(widths = c(2, 1))

g06


# ═══════════════════════════════════════════════════════════════════════
# Graphic 7 — Efficiency Frontier Grid (Individual Runs)
# ═══════════════════════════════════════════════════════════════════════

df_frontier <- read_excel(xlsx_path, sheet = "eff_frontier_runs")
df_frontier$Test <- factor(df_frontier$Test)

front_models <- unique(df_frontier$Model)
front_colors <- setNames(
  sapply(seq_along(front_models), function(i) {
    key <- gsub("-", "_", tolower(trimws(front_models[i])))
    if (key %in% names(MODEL_COLORS)) MODEL_COLORS[[key]]
    else DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1]
  }),
  front_models
)

g07 <- ggplot(df_frontier, aes(x = AI_Positive_Rate_pct, y = LF_Capture_pct,
                                colour = Model, shape = Test)) +
  annotate("rect", xmin = -Inf, xmax = Inf, ymin = 95, ymax = Inf,
           fill = "#D5F5E3", alpha = 0.15) +
  geom_point(size = 3, alpha = 0.9, stroke = 0.8) +
  geom_hline(yintercept = 95, linetype = "dashed", colour = "#27AE60", alpha = 0.5) +
  facet_wrap(~ Project) +
  scale_colour_manual(values = front_colors) +
  scale_shape_manual(values = c("1" = 16, "2" = 15, "3" = 17, "4" = 18)) +
  labs(title = "Efficiency Frontier by Project (Individual Runs)",
       x = "AI Positive Rate (%)", y = "LF Capture (%)") +
  ylim(50, 106) + xlim(-2, 102) +
  BASE_THEME

g07


# ═══════════════════════════════════════════════════════════════════════
# Graphic 8 — Efficiency Frontier Average (per Project)
# ═══════════════════════════════════════════════════════════════════════

agg_frontier <- df_frontier %>%
  group_by(Project, Model) %>%
  summarise(
    mx = mean(AI_Positive_Rate_pct, na.rm = TRUE),
    my = mean(LF_Capture_pct, na.rm = TRUE),
    sx = replace_na(sd(AI_Positive_Rate_pct, na.rm = TRUE), 0),
    sy = replace_na(sd(LF_Capture_pct, na.rm = TRUE), 0),
    .groups = "drop"
  )

g08 <- ggplot(agg_frontier, aes(x = mx, y = my, colour = Model)) +
  geom_errorbar(aes(ymin = my - sy, ymax = my + sy), width = 1, alpha = 0.5) +
  geom_errorbar(aes(xmin = mx - sx, xmax = mx + sx), width = 1, alpha = 0.5,
                orientation = "y") +
  geom_point(size = 5, stroke = 1.5) +
  geom_point(data = df_frontier, aes(x = AI_Positive_Rate_pct, y = LF_Capture_pct),
             size = 1, alpha = 0.3) +
  geom_hline(yintercept = 95, linetype = "dashed", colour = "#27AE60", alpha = 0.5) +
  facet_wrap(~ Project) +
  scale_colour_manual(values = front_colors) +
  labs(title = "Efficiency Frontier by Project (Mean of Runs per Model)",
       x = "AI Positive Rate (%)", y = "LF Capture (%)") +
  ylim(50, 106) + xlim(-2, 102) +
  BASE_THEME

g08


# ═══════════════════════════════════════════════════════════════════════
# Graphic 9 — Efficiency Frontier Overall
# ═══════════════════════════════════════════════════════════════════════

agg_overall <- df_frontier %>%
  group_by(Model) %>%
  summarise(
    mx  = mean(AI_Positive_Rate_pct, na.rm = TRUE),
    my  = mean(LF_Capture_pct, na.rm = TRUE),
    sx  = replace_na(sd(AI_Positive_Rate_pct, na.rm = TRUE), 0),
    sy  = replace_na(sd(LF_Capture_pct, na.rm = TRUE), 0),
    eff = mean(Efficiency_Score, na.rm = TRUE),
    .groups = "drop"
  )

g09 <- ggplot(agg_overall, aes(x = mx, y = my, colour = Model)) +
  annotate("rect", xmin = -Inf, xmax = 50, ymin = 95, ymax = Inf,
           fill = "#D5F5E3", alpha = 0.25) +
  geom_errorbar(aes(ymin = my - sy, ymax = my + sy), width = 1.5, alpha = 0.5, linewidth = 0.8) +
  geom_errorbar(aes(xmin = mx - sx, xmax = mx + sx), width = 1.5, alpha = 0.5, linewidth = 0.8,
                orientation = "y") +
  geom_point(size = 7, stroke = 2) +
  geom_hline(yintercept = 95, linetype = "dashed", colour = "#27AE60", alpha = 0.6) +
  annotate("text", x = 5, y = 102, label = "IDEAL ZONE",
           fontface = "bold", colour = "#27AE60", alpha = 0.5, size = 3.5) +
  scale_colour_manual(values = front_colors) +
  labs(title = "Efficiency Frontier: Overall Mean per Model (\u00b1 1 SD)",
       x = "AI Positive Rate (%) \u2014 lower is more selective",
       y = "Listfinal Capture Rate (%)", colour = NULL) +
  ylim(50, 106) + xlim(-2, 102) +
  BASE_THEME

g09


# ═══════════════════════════════════════════════════════════════════════
# Graphic 10 — Efficiency Score by Project
# ═══════════════════════════════════════════════════════════════════════

df_eff_proj <- read_excel(xlsx_path, sheet = "eff_score_by_project")
df_eff_proj$Label <- paste0(df_eff_proj$Model, "\nT", df_eff_proj$Test)

eff_proj_models <- unique(df_eff_proj$Model)
eff_proj_colors <- setNames(
  sapply(seq_along(eff_proj_models), function(i) {
    key <- gsub("-", "_", tolower(trimws(eff_proj_models[i])))
    if (key %in% names(MODEL_COLORS)) MODEL_COLORS[[key]]
    else DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1]
  }),
  eff_proj_models
)

g10 <- ggplot(df_eff_proj, aes(x = Label, y = Efficiency_Score, fill = Model)) +
  geom_col(alpha = 0.85, colour = "#333333", linewidth = 0.3) +
  geom_text(aes(label = sprintf("%.3f", Efficiency_Score)),
            vjust = -0.4, size = 2.5, fontface = "bold") +
  facet_wrap(~ Project, scales = "free_x") +
  scale_fill_manual(values = eff_proj_colors) +
  labs(title = "Efficiency Score by Model per Project (Individual Runs)",
       y = "Efficiency Score", x = NULL) +
  ylim(0, 1.0) +
  BASE_THEME +
  theme(axis.text.x = element_text(angle = 25, hjust = 1, size = 6.5),
        legend.position = "none")

g10


# ═══════════════════════════════════════════════════════════════════════
# Graphic 11 — Efficiency Score Aggregated
# ═══════════════════════════════════════════════════════════════════════

df_eff_agg <- read_excel(xlsx_path, sheet = "eff_score_aggregated") %>%
  arrange(desc(Mean_Efficiency_Score)) %>%
  mutate(
    Model = factor(Model, levels = Model),
    Rank  = row_number()
  )

eff_agg_models <- as.character(df_eff_agg$Model)
eff_agg_colors <- setNames(
  sapply(seq_along(eff_agg_models), function(i) {
    key <- gsub("-", "_", tolower(trimws(eff_agg_models[i])))
    if (key %in% names(MODEL_COLORS)) MODEL_COLORS[[key]]
    else DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1]
  }),
  eff_agg_models
)

g11 <- ggplot(df_eff_agg, aes(x = Model, y = Mean_Efficiency_Score, fill = Model)) +
  geom_col(alpha = 0.85, colour = "#333333", linewidth = 0.3) +
  geom_errorbar(aes(ymin = Mean_Efficiency_Score - SD,
                    ymax = Mean_Efficiency_Score + SD),
                width = 0.25, linewidth = 0.7, colour = "#555555") +
  geom_text(aes(label = sprintf("%.3f", Mean_Efficiency_Score)),
            vjust = -0.5, size = 3.5, fontface = "bold") +
  geom_text(aes(label = paste0("#", Rank), y = 0.01),
            colour = "white", fontface = "bold", size = 3) +
  scale_fill_manual(values = eff_agg_colors) +
  labs(title = "Efficiency Score by Model (Mean \u00b1 SD, All Projects & Runs)",
       y = "Efficiency Score", x = NULL) +
  ylim(0, max(df_eff_agg$Mean_Efficiency_Score) * 1.25 + 0.05) +
  BASE_THEME +
  theme(legend.position = "none")

g11


# ═══════════════════════════════════════════════════════════════════════
# Graphic 12 — Sensitivity & Specificity Dual Gold Standard
# ═══════════════════════════════════════════════════════════════════════

df_dual <- read_excel(xlsx_path, sheet = "sens_spec_dual_gold")

df_dual_long <- df_dual %>%
  pivot_longer(cols = -Model, names_to = "Metric", values_to = "Value") %>%
  mutate(
    Gold = ifelse(grepl("TIAB", Metric), "TIAB", "Listfinal"),
    Type = ifelse(grepl("Sens", Metric), "Sensitivity", "Specificity")
  )

g12 <- ggplot(df_dual_long, aes(x = Model, y = Value, fill = interaction(Type, Gold))) +
  geom_col(position = position_dodge(width = 0.8), width = 0.18,
           alpha = 0.85, colour = "#333333", linewidth = 0.3) +
  geom_hline(yintercept = 95, linetype = "dashed", colour = "#27AE60", alpha = 0.4) +
  geom_text(aes(label = sprintf("%.0f", Value)),
            position = position_dodge(width = 0.8), vjust = -0.4, size = 2) +
  scale_fill_manual(
    values = c(
      "Sensitivity.TIAB"      = "#4472C4",
      "Sensitivity.Listfinal" = "#27AE60",
      "Specificity.TIAB"      = "#ED7D31",
      "Specificity.Listfinal" = "#FFC000"
    ),
    labels = c("Sens (TIAB)", "Sens (Listfinal)", "Spec (TIAB)", "Spec (Listfinal)")
  ) +
  labs(title = "Sensitivity & Specificity: TIAB vs Listfinal Gold Standard",
       y = "Percentage (%)", x = NULL, fill = NULL) +
  ylim(0, 110) +
  BASE_THEME +
  theme(axis.text.x = element_text(face = "bold", size = 8))

g12


# ═══════════════════════════════════════════════════════════════════════
# Graphic 13 — Aggregated Model Performance
# ═══════════════════════════════════════════════════════════════════════

df_agg_perf <- read_excel(xlsx_path, sheet = "aggregated_performance")

mean_cols <- grep("_mean$", names(df_agg_perf), value = TRUE)
sd_cols   <- grep("_sd$",   names(df_agg_perf), value = TRUE)

long_mean <- df_agg_perf %>%
  select(Model, all_of(mean_cols)) %>%
  pivot_longer(-Model, names_to = "Metric_raw", values_to = "Mean") %>%
  mutate(Metric = gsub("_mean$", "", Metric_raw))

long_sd <- df_agg_perf %>%
  select(Model, all_of(sd_cols)) %>%
  pivot_longer(-Model, names_to = "Metric_raw", values_to = "SD") %>%
  mutate(Metric = gsub("_sd$", "", Metric_raw))

df_perf_long <- inner_join(
  long_mean %>% select(-Metric_raw),
  long_sd   %>% select(-Metric_raw),
  by = c("Model", "Metric")
)

g13 <- ggplot(df_perf_long, aes(x = Model, y = Mean)) +
  geom_col(fill = "#5B9BD5", alpha = 0.7, colour = "#333333", linewidth = 0.3, width = 0.6) +
  geom_errorbar(aes(ymin = Mean - SD, ymax = Mean + SD), width = 0.2,
                linewidth = 0.5, colour = "#555555") +
  geom_hline(yintercept = 0.95, linetype = "dashed", colour = "green", alpha = 0.4) +
  facet_wrap(~ Metric, scales = "free_y", ncol = 3) +
  labs(title = "Aggregated Model Performance (Mean \u00b1 SD across Projects)",
       y = NULL, x = NULL) +
  ylim(0, 1.08) +
  BASE_THEME +
  theme(axis.text.x = element_text(angle = 30, hjust = 1, size = 7))

g13


# ═══════════════════════════════════════════════════════════════════════
# Graphic 14 — F1 Score vs Cost
# ═══════════════════════════════════════════════════════════════════════

df_f1_cost <- read_excel(xlsx_path, sheet = "f1_vs_cost")

g14 <- ggplot(df_f1_cost, aes(x = Avg_Cost_USD, y = Avg_F1_LF)) +
  geom_point(size = 4, colour = "#5B9BD5", stroke = 1) +
  geom_hline(yintercept = 0.95, linetype = "dashed", colour = "green",
             alpha = 0.5, linewidth = 0.8) +
  geom_text_repel(aes(label = Model), size = 3, fontface = "bold", segment.colour = "grey60") +
  labs(title = "F1 Score (Full-Text Gold Standard) vs Cost per Model",
       x = "Average Cost (USD)", y = "Average F1 Score (vs Listfinal)") +
  ylim(0, 1.05) +
  BASE_THEME

g14


# ═══════════════════════════════════════════════════════════════════════
# Graphic 15 — Sensitivity vs Specificity Trade-Off
# ═══════════════════════════════════════════════════════════════════════

df_tradeoff <- read_excel(xlsx_path, sheet = "sens_spec_tradeoff")

trade_models <- unique(df_tradeoff$Model)
trade_colors <- setNames(
  sapply(seq_along(trade_models), function(i) {
    key <- gsub("-", "_", tolower(trimws(trade_models[i])))
    if (key %in% names(MODEL_COLORS)) MODEL_COLORS[[key]]
    else DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1]
  }),
  trade_models
)

g15 <- ggplot(df_tradeoff, aes(x = Specificity_pct, y = Sensitivity_pct,
                                colour = Model, shape = Project)) +
  geom_point(size = 3, alpha = 0.8, stroke = 0.6) +
  geom_hline(yintercept = 95, linetype = "dashed", colour = "#27AE60", alpha = 0.5) +
  geom_vline(xintercept = 50, linetype = "dotted", colour = "#E74C3C", alpha = 0.4) +
  facet_wrap(~ Gold_Standard) +
  scale_colour_manual(values = trade_colors) +
  scale_shape_manual(values = c("mino" = 16, "NMDA" = 15, "zebra" = 17)) +
  labs(title = "Sensitivity vs Specificity under Two Gold Standards",
       x = "Specificity (%)", y = "Sensitivity (%)") +
  xlim(0, 105) + ylim(0, 105) +
  BASE_THEME

g15


# ═══════════════════════════════════════════════════════════════════════
# Graphic 16 — Model Ranking Heatmap
# ═══════════════════════════════════════════════════════════════════════

df_ranking <- read_excel(xlsx_path, sheet = "model_ranking_heatmap")

rank_metric_cols <- setdiff(names(df_ranking), c("Model", "Overall_Score"))

df_rank_long <- df_ranking %>%
  mutate(Rank = row_number()) %>%
  pivot_longer(cols = all_of(rank_metric_cols), names_to = "Metric", values_to = "Value") %>%
  group_by(Metric) %>%
  mutate(Norm = if (n() < 2 | diff(range(Value, na.rm = TRUE)) < 1e-9) 0.5
                else (Value - min(Value, na.rm = TRUE)) /
                     (max(Value, na.rm = TRUE) - min(Value, na.rm = TRUE))) %>%
  ungroup() %>%
  mutate(
    Metric = factor(Metric, levels = rank_metric_cols),
    Model  = factor(Model, levels = df_ranking$Model),
    Label  = ifelse(is.na(Value), "",
                    ifelse(Value > 1.5, sprintf("%.1f%%", Value), sprintf("%.3f", Value)))
  )

g16 <- ggplot(df_rank_long, aes(x = Metric, y = Model, fill = Norm)) +
  geom_tile(colour = "white", linewidth = 1.5) +
  geom_text(aes(label = Label), fontface = "bold", size = 3) +
  scale_fill_gradient2(
    low = "#E74C3C", mid = "#FFFF00", high = "#27AE60",
    midpoint = 0.5, limits = c(0, 1), name = "Normalised\nScore"
  ) +
  scale_y_discrete(
    limits = rev(levels(df_rank_long$Model)),
    labels = function(x) paste0("#", rev(seq_along(x)), " ", rev(x))
  ) +
  labs(title = "Model Ranking Summary (Mean Across All Projects & Runs)",
       x = NULL, y = NULL) +
  theme_minimal(base_size = 10) +
  theme(
    plot.title        = element_text(face = "bold", size = 12),
    axis.text.x       = element_text(size = 8, hjust = 0.5),
    axis.text.y       = element_text(size = 9, face = "bold"),
    panel.grid        = element_blank(),
    legend.position   = "right",
    axis.ticks.length = unit(0, "pt")
  )

g16


# ═══════════════════════════════════════════════════════════════════════
# Save All Charts
# ═══════════════════════════════════════════════════════════════════════

n_proj <- length(unique(df_sensitivity$Project))

ggsave(file.path(out_dir, "01_sensitivity_per_model_per_project.png"),
       g01, width = 4.5 * n_proj, height = 4.5, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "02_listfinal_capture_heatmap.png"),
       g02, width = 3 + length(unique(df_lf_capture$Project)) * 1.5,
       height = 1.5 + length(unique(df_lf_capture$Model)) * 0.8,
       dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "03_test_retest_kappa.png"),
       g03, width = 10, height = 5, dpi = DPI, bg = "white")

# Graphic 4 — Radar (base R)
png(file.path(out_dir, "04_model_comparison_radar.png"),
    width = 7, height = 7, units = "in", res = DPI, bg = "white")
par(mar = c(1, 1, 3, 1))
fmsb::radarchart(radar_data,
                 axistype = 1,
                 pcol = radar_colors,
                 pfcol = adjustcolor(radar_colors, alpha.f = 0.1),
                 plwd = 2, plty = 1,
                 cglcol = "grey70", cglty = 1, cglwd = 0.5,
                 axislabcol = "grey40", vlcex = 0.8,
                 title = "Model Comparison \u2014 Key Metrics Radar")
legend("topright", legend = df_radar$Model, col = radar_colors, lwd = 2, bty = "n", cex = 0.8)
dev.off()

ggsave(file.path(out_dir, "05_cost_vs_sensitivity_bubble.png"),
       g05, width = 9, height = 5.5, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "06_workload_reduction.png"),
       g06, width = 12, height = 4.5, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "07_efficiency_frontier_grid.png"),
       g07, width = 5.5 * min(n_proj, 3), height = 5, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "08_efficiency_frontier_by_project_avg.png"),
       g08, width = 5.5 * n_proj, height = 5.5, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "09_efficiency_frontier_averaged.png"),
       g09, width = 9, height = 6, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "10_efficiency_score_by_project.png"),
       g10, width = 5 * n_proj, height = 5, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "11_efficiency_score_aggregated.png"),
       g11, width = 8, height = 5, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "12_sensitivity_specificity_dual_gold.png"),
       g12, width = 10, height = 5, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "13_aggregated_performance_metrics.png"),
       g13, width = 14, height = 8, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "14_f1_score_vs_cost.png"),
       g14, width = 8, height = 5, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "15_sensitivity_specificity_tradeoff.png"),
       g15, width = 12, height = 5.5, dpi = DPI, bg = "white")

ggsave(file.path(out_dir, "16_model_ranking_heatmap.png"),
       g16, width = 10, height = max(3.5, nrow(df_ranking) * 0.9),
       dpi = DPI, bg = "white")

cat("\n  Done — 16 charts saved to", out_dir, "\n")
