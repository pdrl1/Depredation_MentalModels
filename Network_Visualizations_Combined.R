# ============================================================
# COMBINED NETWORK VISUALIZATIONS
# Australia (AU) · Alabama (AL) · Galveston (GA)
# ============================================================
#
# Produces 7 publication-ready plots, each a 1-row × 3-column
# facet (one panel per region):
#
#   P1  Degree centrality       — full models, top N concepts
#   P2  Betweenness centrality  — full models, top N concepts
#   P3  Eigenvector centrality  — full models, top N concepts
#   P4  Concept type composition — all models (full + sub-regions)
#   P5  Top-5 by degree per sub-region — 3 side-by-side panels
#   P6  Indegree vs Outdegree   — full models, coloured by Type
#   P7  Validation metrics      — R/T ratio, type counts per model
#
# HOW TO RUN:
#   Set the three OUT_DIR paths in Section 1, then:
#   source("Network_Visualizations_Combined.R")
# ============================================================


# ── SECTION 0 — PACKAGES ─────────────────────────────────────

pkgs <- c("dplyr", "tidyr", "ggplot2", "scales",
          "stringr", "purrr", "forcats", "readr",
          "RColorBrewer", "ggrepel", "patchwork", "tidytext")
for (p in pkgs) {
  if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
  library(p, character.only = TRUE)
}


# ── SECTION 1 — PARAMETERS ───────────────────────────────────

# Folders where each region's *_centrality.csv files live
# (the output folders from each Network_Analysis_*.R script)
AU_DIR <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Australia/Network_Outputs"
AL_DIR <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Alabama/Network_Outputs"
GA_DIR <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Network_Outputs"

# Where to save the combined plots
OUT_DIR    <- "~/path/to/Combined_Plots"
SAVE_PLOTS <- TRUE

# Top N concepts shown in bar-chart plots (P1, P2, P3)
TOP_N <- 20

# Colour palette per region (used consistently across all plots)
REGION_COLOURS <- c(
  "Australia" = "#1B9E77",
  "Alabama"   = "#D95F02",
  "Galveston" = "#7570B3"
)

# Concept type colours (used consistently)
TYPE_COLOURS <- c(
  "Transmitter" = "#f89880",
  "Ordinary"    = "#74add1",
  "Receiver"    = "#009e60"
)
TYPE_ORDER <- c("Transmitter", "Ordinary", "Receiver")


# ── SECTION 2 — LOAD DATA ────────────────────────────────────

load_region <- function(dir_path, region_label, prefix) {
  csv_files <- list.files(dir_path,
                          pattern = paste0("^", prefix, ".*_centrality\\.csv$"),
                          full.names = TRUE)
  if (length(csv_files) == 0)
    stop(sprintf("No centrality CSVs found for %s in: %s",
                 region_label, dir_path))
  
  map_dfr(csv_files, function(f) {
    base     <- tools::file_path_sans_ext(basename(f))
    base     <- sub("_centrality$", "", base)
    model    <- sub(paste0("^", prefix, "_"), "", base)
    model    <- str_replace_all(model, "_", " ")
    scope    <- if (tolower(model) == "full") "full" else "sub"
    model_id <- paste0(prefix, "_", str_replace_all(model, " ", "_"))
    
    read_csv(f, show_col_types = FALSE) %>%
      mutate(
        Region   = region_label,
        Model    = model,
        Scope    = scope,
        Model_ID = model_id
      )
  })
}

if (!dir.exists(OUT_DIR)) dir.create(OUT_DIR, recursive = TRUE)

all_df <- bind_rows(
  load_region(AU_DIR, "Australia", "AU"),
  load_region(AL_DIR, "Alabama",   "AL"),
  load_region(GA_DIR, "Galveston", "GA")
) %>%
  mutate(
    Region = factor(Region, levels = c("Australia", "Alabama", "Galveston")),
    Type   = factor(Type,   levels = TYPE_ORDER)
  )

full_df <- filter(all_df, Scope == "full")
sub_df  <- filter(all_df, Scope == "sub")

cat(sprintf("Loaded %d models | %d concept-model rows\n",
            n_distinct(all_df$Model_ID), nrow(all_df)))


# ── SHARED HELPER FUNCTIONS ───────────────────────────────────

# Shared plot theme
theme_fcm <- function(base_size = 10) {
  theme_bw(base_size = base_size) %+replace%
    theme(
      strip.text = element_text(
        face = "bold",
        size = base_size + 1,
        margin = margin(b = 6)),
      strip.background = element_rect(colour = NA),
      axis.text.y      = element_text(size = base_size - 2),
      axis.text.x      = element_text(size = base_size - 2),
      panel.grid.minor = element_blank(),
      legend.position  = "bottom",
      legend.title     = element_text(size = base_size - 1),
      plot.title       = element_text(face = "bold", size = base_size + 2, margin = margin(b = 4)),
      plot.subtitle    = element_text(size = base_size - 1, margin=margin(b=8))
    )
}

# Per-facet bar chart (horizontal): top N concepts by a given metric
# Uses tidytext::reorder_within so each panel is sorted independently
bar_facet <- function(df, metric_col, fill_col = "Type",
                      top_n = TOP_N,
                      xlab = metric_col) {
  plot_df <- df %>%
    group_by(Region) %>%
    slice_max(order_by = .data[[metric_col]], n = top_n,
              with_ties = FALSE) %>%
    ungroup() %>%
    mutate(
      Concept_ordered = reorder_within(
        str_wrap(Concept, 28), .data[[metric_col]], Region
      )
    )
  
  ggplot(plot_df,
         aes(x     = Concept_ordered,
             y     = .data[[metric_col]],
             fill  = .data[[fill_col]])) +
    geom_col(width = 0.75) +
    geom_text(aes(label = round(.data[[metric_col]], 2)),
              hjust = -0.15, size = 2.4, colour = "grey30") +
    coord_flip() +
    scale_x_reordered() +
    scale_fill_manual(values = TYPE_COLOURS, drop = FALSE) +
    facet_wrap(~ Region, nrow = 1, scales = "free") +
    expand_limits(y = max(plot_df[[metric_col]], na.rm = TRUE) * 1.18) +
    labs(x = NULL, y = xlab, fill = "Concept type") +
    theme_fcm()
}

save_plot <- function(p, filename, width = 18, height = 8) {
  if (SAVE_PLOTS)
    ggsave(file.path(OUT_DIR, filename), p,
           width = width, height = height)
  print(p)
  invisible(p)
}

# ================================================================
# PLOT 1 — DEGREE CENTRALITY (full models, top N)
# ================================================================

cat("\nPlot 1: Degree centrality...\n")

p1 <- bar_facet(full_df, "Degree",
                xlab = "Degree centrality (sum |incoming| + |outgoing| weights)") +
  labs(
    title    = "Degree Centrality — Full Models",
    subtitle = paste0("Top ", TOP_N,
                      " concepts per region")
  )


save_plot(p1, "P1_degree_centrality.pdf", width = 18, height = 9)

png(filename = "Degree_centrality.png", 
    width = 17, height = 6, units = "in", # Set size in inches
    res = 600)  
p1
dev.off()


# ================================================================
# PLOT 1b — INDEGREE CENTRALITY (full models, top N)
# ================================================================

cat("Plot 1b: Indegree centrality...\n")

p1b <- bar_facet(full_df, "Indegree",
                 xlab = "Indegree (sum of incoming |weights|)") +
  labs(
    title    = "Indegree Centrality — Full Models",
    subtitle = paste0("Top ", TOP_N,
                      " most influenced concepts per region | fill = concept type")
  )

save_plot(p1b, "P1b_indegree_centrality.pdf", width = 18, height = 9)

png(filename = "P1b_indegree_centrality.png", 
    width = 17, height = 6, units = "in", # Set size in inches
    res = 600)  
p1b
dev.off()

# ================================================================
# PLOT 1c — OUTDEGREE CENTRALITY (full models, top N)
# ================================================================

cat("Plot 1c: Outdegree centrality...\n")

p1c <- bar_facet(full_df, "Outdegree",
                 xlab = "Outdegree (sum of outgoing |weights|)") +
  labs(
    title    = "Outdegree Centrality — Full Models",
    subtitle = paste0("Top ", TOP_N,
                      " most influential (driver) concepts per region | fill = concept type")
  )

save_plot(p1c, "P1c_outdegree_centrality.pdf", width = 18, height = 9)

png(filename = "P1c_outdegree_centrality.png", 
    width = 17, height = 6, units = "in", # Set size in inches
    res = 600)  
p1c
dev.off()

# ================================================================
# PLOT 2 — BETWEENNESS CENTRALITY (full models, top N)
# ================================================================

cat("Plot 2: Betweenness centrality...\n")

p2 <- bar_facet(full_df, "Betweenness",
                xlab = "Betweenness centrality (normalised)") +
  labs(
    title    = "Betweenness Centrality — Full Models",
    subtitle = paste0("Top ", TOP_N,
                      " concepts | high betweenness = key bridge / bottleneck")
  )

save_plot(p2, "P2_betweenness_centrality.pdf", width = 18, height = 9)

png(filename = "betweenneess.png", 
    width = 17, height = 6, units = "in", # Set size in inches
    res = 600)  
p2
dev.off()

# ================================================================
# PLOT 3 — EIGENVECTOR CENTRALITY (full models, top N)
# ================================================================

cat("Plot 3: Eigenvector centrality...\n")

p3 <- bar_facet(full_df, "Eigenvector",
                xlab = "Eigenvector centrality (normalised 0–1)") +
  labs(
    title    = "Eigenvector Centrality — Full Models",
    subtitle = paste0("Top ", TOP_N,
                      " concepts | high eigenvector = well-connected to other central concepts")
  )

save_plot(p3, "P3_eigenvector_centrality.pdf", width = 18, height = 9)

png(filename = "eigenvector_centrality.png", 
    width = 17, height = 6, units = "in", # Set size in inches
    res = 600)  
p3
dev.off()

# ================================================================
# PLOT 4 — CONCEPT TYPE COMPOSITION (all models per region)
# ================================================================
# One panel per region; x-axis = each model (full + sub-regions);
# bars show proportion of Transmitters / Ordinary / Receivers.

cat("Plot 4: Concept type composition...\n")

type_df <- all_df %>%
  count(Region, Model_ID, Model, Scope, Type, .drop = FALSE) %>%
  group_by(Region, Model_ID) %>%
  mutate(
    pct        = n / sum(n),
    Model_wrap = str_wrap(Model, 12),
    # Order x-axis: full model first, then subs alphabetically
    Model_wrap = fct_reorder(Model_wrap, Scope == "full",
                             .desc = TRUE)
  ) %>%
  ungroup()

p4 <- ggplot(type_df,
             aes(x = Model_wrap, y = pct, fill = Type)) +
  geom_col(position = "fill", width = 0.8) +
  geom_text(
    aes(label = ifelse(pct > 0.07 & n > 0, n, "")),
    position = position_fill(vjust = 0.5),
    size = 2.6, colour = "white", fontface = "bold"
  ) +
  scale_y_continuous(labels = percent_format(accuracy = 1)) +
  scale_fill_manual(values = TYPE_COLOURS, drop = FALSE) +
  facet_wrap(~ Region, nrow = 1, scales = "free_x") +
  labs(
    title    = "Concept Type Composition — All Models",
    subtitle = "Transmitters (blue) = external drivers  |  Receivers (red) = outcomes\nNumbers = concept count per type",
    x = NULL, y = "Proportion", fill = "Concept type"
  ) +
  theme_fcm() +
  theme(axis.text.x = element_text(angle = 35, hjust = 1, size = 8))

save_plot(p4, "P4_type_composition.pdf", width = 16, height = 7)


# ================================================================
# PLOT 5 — TOP 5 BY DEGREE PER SUB-REGION
# ================================================================
# Each region gets its own patchwork panel (sub-regions as
# x-axis groups, top-5 concepts as labelled bars).
# The three panels are then assembled 1 × 3.

cat("Plot 5: Top-5 concepts per sub-region...\n")

make_subregion_panel <- function(region_label) {
  df <- sub_df %>%
    filter(Region == region_label) %>%
    group_by(Model) %>%
    slice_max(Degree, n = 5, with_ties = FALSE) %>%
    mutate(Rank = row_number()) %>%
    ungroup() %>%
    mutate(
      Concept_wrap = str_wrap(Concept, 20),
      Model        = str_wrap(Model, 12)
    )
  
  ggplot(df, aes(x = factor(Rank), y = Degree,
                 fill = Region, label = Concept_wrap)) +
    geom_col(show.legend = FALSE,
             fill = REGION_COLOURS[[region_label]], alpha = 0.85) +
    geom_text(hjust = -0.08, size = 2.3, colour = "grey20") +
    coord_flip() +
    expand_limits(y = max(df$Degree) * 1.55) +
    scale_x_discrete(labels = paste0("#", 1:5)) +
    facet_wrap(~ Model, nrow = 1, scales = "free_x") +
    labs(
      title = region_label,
      x = NULL, y = "Degree"
    ) +
    theme_fcm(base_size = 8) +
    theme(
      plot.title    = element_text(colour = REGION_COLOURS[[region_label]],
                                   face = "bold", size = 10),
      axis.text.x   = element_text(size = 7),
      strip.text    = element_text(size = 7)
    )
}

p5_au <- make_subregion_panel("Australia")
p5_al <- make_subregion_panel("Alabama")
p5_ga <- make_subregion_panel("Galveston")

p5 <- p5_au / p5_al / p5_ga +
  plot_annotation(
    title    = "Top 5 Concepts by Degree — Sub-Regions",
    subtitle = "Each row = one region; each column within = one sub-region",
    theme    = theme(
      plot.title    = element_text(face = "bold", size = 13),
      plot.subtitle = element_text(size = 10, colour = "grey40")
    )
  )

save_plot(p5, "P5_subregion_top5.pdf", width = 18, height = 14)


png(filename = "top5_degree_subregion.png", 
    width = 17, height = 10, units = "in", # Set size in inches
    res = 600)  
p5
dev.off()

# ================================================================
# PLOT 6 — INDEGREE vs OUTDEGREE (full models)
# ================================================================
# Scatter plot: x = Indegree, y = Outdegree
# Diagonal = balanced concept
# Above diagonal = more outgoing (driver)
# Below diagonal = more incoming (receiver)
# Points labelled for top-degree concepts

cat("Plot 6: Indegree vs Outdegree...\n")

# Label top concepts (top 12 by Degree per region)
label_df <- full_df %>%
  group_by(Region) %>%
  slice_max(Degree, n = 12, with_ties = FALSE) %>%
  ungroup()

# Axis limit: same scale within each facet (free), reference line needs manual
p6 <- ggplot(full_df,
             aes(x = Indegree, y = Outdegree,
                 colour = Type, size = Degree)) +
  # Reference diagonal (balanced)
  geom_abline(slope = 1, intercept = 0,
              linetype = "dashed", colour = "grey60", linewidth = 0.5) +
  geom_point(alpha = 0.65) +
  geom_label_repel(
    data          = label_df,
    aes(label     = str_wrap(Concept, 20)),
    size          = 2.3,
    label.padding = unit(0.12, "lines"),
    max.overlaps  = 12,
    show.legend   = FALSE,
    colour        = "grey20",
    fill          = alpha("white", 0.8)
  ) +
  scale_colour_manual(values = TYPE_COLOURS, drop = FALSE) +
  scale_size_continuous(range = c(1.5, 6), guide = "none") +
  facet_wrap(~ Region, nrow = 1, scales = "free") +
  labs(
    title    = "Indegree vs Outdegree — Full Models",
    subtitle = "Above diagonal = more outgoing (driver)  |  Below = more incoming (receiver)  |  Size = Degree",
    x = "Indegree (sum of incoming |weights|)",
    y = "Outdegree (sum of outgoing |weights|)",
    colour = "Concept type"
  ) +
  theme_fcm()

save_plot(p6, "P6_indegree_vs_outdegree.pdf", width = 18, height = 7)

png(filename = "P6_indegree_vs_outdegree.png", 
    width = 17, height = 8, units = "in", # Set size in inches
    res = 600)  
p6
dev.off()

# ================================================================
# PLOT 6b — CLOSENESS vs BETWEENNESS (full models)
# ================================================================
# Closeness: how quickly a concept can reach all others
#            (high = central broadcaster)
# Betweenness: how often a concept sits on shortest paths
#              between other pairs (high = bottleneck / bridge)
#
# Four quadrants:
#   High closeness + high betweenness → key hub (fast & a bridge)
#   High closeness + low betweenness  → efficient broadcaster, not a bridge
#   Low closeness  + high betweenness → bottleneck connecting distant parts
#   Low closeness  + low betweenness  → peripheral concept

cat("Plot 6b: Closeness vs Betweenness...\n")

label_df_cb <- full_df %>%
  group_by(Region) %>%
  slice_max(Degree, n = 12, with_ties = FALSE) %>%
  ungroup()

# Compute per-region median lines for reference quadrant annotation
quad_lines <- full_df %>%
  group_by(Region) %>%
  summarise(
    med_closeness    = median(Closeness,    na.rm = TRUE),
    med_betweenness  = median(Betweenness,  na.rm = TRUE),
    .groups = "drop"
  )

p6b <- ggplot(full_df,
              aes(x = Closeness, y = Betweenness,
                  colour = Type, size = Degree)) +
  # Median reference lines per facet
  geom_vline(data = quad_lines,
             aes(xintercept = med_closeness),
             linetype = "dashed", colour = "grey70", linewidth = 0.4) +
  geom_hline(data = quad_lines,
             aes(yintercept = med_betweenness),
             linetype = "dashed", colour = "grey70", linewidth = 0.4) +
  geom_point(alpha = 0.65) +
  geom_label_repel(
    data          = label_df_cb,
    aes(label     = str_wrap(Concept, 20)),
    size          = 2.3,
    label.padding = unit(0.12, "lines"),
    max.overlaps  = 12,
    show.legend   = FALSE,
    colour        = "grey20",
    fill          = alpha("white", 0.8)
  ) +
  scale_colour_manual(values = TYPE_COLOURS, drop = FALSE) +
  scale_size_continuous(range = c(1.5, 6), guide = "none") +
  facet_wrap(~ Region, nrow = 1, scales = "free") +
  labs(
    title    = "Closeness vs Betweenness Centrality — Full Models",
    subtitle = "Top-right = fast broadcaster AND bridge  |  Top-left = bottleneck only  |  Bottom-right = broadcaster only\nDashed lines = per-region medians  |  Size = Degree",
    x = "Closeness (harmonic mean reach; higher = faster broadcaster)",
    y = "Betweenness (normalised; higher = more of a bridge/bottleneck)",
    colour = "Concept type"
  ) +
  theme_fcm()

save_plot(p6b, "P6b_closeness_vs_betweenness.pdf", width = 18, height = 7)


png(filename = "P6b_closeness_vs_betweenness.png", 
    width = 17, height = 8, units = "in", # Set size in inches
    res = 600)  
p6b
dev.off()


# ================================================================
# PLOT 8 — Katz Centrality
# ================================================================

p_katz <- bar_facet(full_df, "Katz",
                    xlab = "Katz centrality (normalised 0–1)") +
  labs(
    title    = "Katz Centrality — Full Models",
    subtitle = paste0("Top ", TOP_N,
                      " concepts per region | Katz is robust to concepts with no feedback loops | fill = concept type")
  )

save_plot(p_katz, "P_katz_centrality.pdf", width = 18, height = 9)

png(filename = "P_katz_centrality.png", 
    width = 17, height = 8, units = "in", # Set size in inches
    res = 600)  
p_katz
dev.off()


# ================================================================
# PLOT 7 — VALIDATION METRICS (computed from Type column)
# ================================================================
# Metrics that can be derived from the centrality CSVs:
#   - N concepts, N Transmitters, N Receivers, N Ordinary
#   - R/T ratio
#   - Mean degree, max degree
# Path-based metrics (clustering, diameter) need the raw graph
# and cannot be recovered from the centrality CSVs alone.

cat("Plot 7: Validation metrics...\n")

val_df <- all_df %>%
  group_by(Region, Model_ID, Model, Scope) %>%
  summarise(
    N_concepts    = n(),
    N_transmitter = sum(Type == "Transmitter", na.rm = TRUE),
    N_receiver    = sum(Type == "Receiver",    na.rm = TRUE),
    N_ordinary    = sum(Type == "Ordinary",    na.rm = TRUE),
    N_isolated    = sum(Type == "Isolated",    na.rm = TRUE),
    RT_ratio      = ifelse(N_transmitter > 0,
                           round(N_receiver / N_transmitter, 2), NA_real_),
    Mean_degree   = round(mean(Degree, na.rm = TRUE), 2),
    Max_degree    = round(max(Degree,  na.rm = TRUE), 2),
    .groups = "drop"
  ) %>%
  mutate(
    Model_wrap = str_wrap(Model, 12),
    Scope_label = ifelse(Scope == "full", "Full model", "Sub-region")
  )

# Reshape to long for faceting by metric
val_long <- val_df %>%
  select(Region, Model_wrap, Scope_label,
         `N concepts` = N_concepts,
         `N transmitters` = N_transmitter,
         `N receivers`    = N_receiver,
         `R/T ratio`      = RT_ratio,
         `Mean degree`    = Mean_degree) %>%
  pivot_longer(
    cols      = -c(Region, Model_wrap, Scope_label),
    names_to  = "Metric",
    values_to = "Value"
  ) %>%
  mutate(
    Metric = factor(Metric,
                    levels = c("N concepts", "N transmitters",
                               "N receivers", "R/T ratio",
                               "Mean degree"))
  )

p7 <- ggplot(val_long,
             aes(x = Model_wrap, y = Value,
                 fill = Scope_label)) +
  geom_col(width = 0.7, position = "dodge") +
  geom_text(aes(label = ifelse(!is.na(Value),
                               as.character(round(Value, 1)), "")),
            position = position_dodge(0.7),
            vjust = -0.3, size = 2.3, colour = "grey30") +
  scale_fill_manual(
    values = c("Full model"  = "#1f78b4",
               "Sub-region"  = "#a6cee3"),
    name = NULL
  ) +
  facet_grid(Metric ~ Region, scales = "free") +
  labs(
    title    = "Validation Metrics — All Models",
    subtitle = "Computed from concept type classification | Path-based metrics not shown (need raw graph)",
    x = NULL, y = "Value"
  ) +
  theme_fcm(base_size = 9) +
  theme(
    axis.text.x  = element_text(angle = 35, hjust = 1, size = 7),
    strip.text.y = element_text(size = 8)
  )

save_plot(p7, "P7_validation_metrics.pdf", width = 16, height = 14)




# ── DONE ──────────────────────────────────────────────────────

cat("\n============================================================\n")
cat("ALL PLOTS SAVED\n")
cat(sprintf("Directory: %s\n\n", OUT_DIR))
cat("  P1_degree_centrality.pdf       — Top ", TOP_N, " by Degree (full models)\n")
cat("  P2_betweenness_centrality.pdf  — Top ", TOP_N, " by Betweenness (full models)\n")
cat("  P3_eigenvector_centrality.pdf  — Top ", TOP_N, " by Eigenvector (full models)\n")
cat("  P4_type_composition.pdf        — Type composition, all models\n")
cat("  P5_subregion_top5.pdf          — Top 5 per sub-region (3 stacked panels)\n")
cat("  P6_indegree_vs_outdegree.pdf   — Scatter Indegree × Outdegree (full models)\n")
cat("  P7_validation_metrics.pdf      — N concepts, R/T ratio, mean degree\n")
cat("============================================================\n")