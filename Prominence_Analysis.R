# ============================================================
# PROMINENCE ANALYSIS — Hoffman et al. (2014)
# Australia · Alabama · Galveston
# ============================================================
#
# Prominence = average of Occurrence and normalised Centrality
#
# Occurrence   = number of sub-region/sector models a concept
#                appears in ÷ total models in that community
#                (already 0–1 by definition)
#
# Centrality   = average Degree centrality across all models
#                the concept appears in, then min-max
#                normalised to 0–1 within its community
#
# Prominence   = (Occurrence + Norm_Centrality) / 2
#
# Community definitions (sub-region/sector models only;
# full aggregated models are NOT counted as community models):
#   Australia → 4 models: NSW, North Australia,
#                          Queensland, Western Australia
#   Alabama   → 5 models: Alabama, Florida, Louisiana,
#                          Mississippi, Texas
#   Galveston → 3 models: Charter, Commercial, Recreational
#
# HOW TO RUN:
#   Set CSV_DIRS in Section 1, then:
#   source("Prominence_Analysis.R")
# ============================================================


# ── SECTION 0 — PACKAGES ─────────────────────────────────────

pkgs <- c("dplyr", "tidyr", "readr", "stringr",
          "ggplot2", "scales", "forcats", "ggrepel")
for (p in pkgs) {
  if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
  library(p, character.only = TRUE)
}


# ── SECTION 1 — PARAMETERS ───────────────────────────────────

# Folders containing the *_centrality.csv files
CSV_DIRS <- c(
  "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Australia/Network_Outputs",
  "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Alabama/Network_Outputs",
  "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Network_Outputs"
)

OUT_DIR    <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Prominence_Outputs"
SAVE_CSV   <- TRUE
SAVE_PLOTS <- TRUE

# Top N concepts shown in prominence plots
TOP_N <- 20


# ── SECTION 2 — LOAD SUB-REGION CSVs ─────────────────────────
# Only sub-region/sector models are loaded here.
# Full models (*_full_centrality.csv) are excluded because
# prominence is computed from the individual community models,
# not the aggregate.

if (!dir.exists(OUT_DIR)) dir.create(OUT_DIR, recursive = TRUE)

csv_files <- unlist(lapply(CSV_DIRS, function(d) {
  list.files(d, pattern = "_centrality\\.csv$", full.names = TRUE)
}))

# Parse filename → Region, Model, Scope
parse_filename <- function(path) {
  base   <- tools::file_path_sans_ext(basename(path))
  base   <- sub("_centrality$", "", base)
  parts  <- str_split(base, "_", n = 2)[[1]]
  region <- parts[1]
  model  <- if (length(parts) == 2) str_replace_all(parts[2], "_", " ")
  else "full"
  scope  <- if (tolower(model) == "full") "full" else "sub"
  list(region = region, model = model, scope = scope)
}

# Map prefix to community label
community_labels <- c(
  "AU" = "Australia",
  "AL" = "Alabama",
  "GA" = "Galveston"
)

all_sub_df <- purrr::map_dfr(csv_files, function(f) {
  info <- parse_filename(f)
  if (info$scope != "sub") return(NULL)   # skip full models
  read_csv(f, show_col_types = FALSE) %>%
    mutate(
      Community = community_labels[info$region],
      Model     = info$model
    )
})

# Count total models per community (for occurrence denominator)
total_models <- all_sub_df %>%
  distinct(Community, Model) %>%
  count(Community, name = "Total_models")

cat("Community model counts:\n")
print(total_models)


# ── SECTION 3 — COMPUTE PROMINENCE ───────────────────────────

prominence_df <- all_sub_df %>%
  
  # Step 1: per concept × community, count appearances and
  #         average Degree across the models it appears in
  group_by(Community, Concept) %>%
  summarise(
    N_models     = n_distinct(Model),       # appearances
    Avg_Degree   = mean(Degree, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  
  #N_models is concept-specific — it is how many of those models that particular concept actually appears in. 

  # Step 2: attach total model count
  left_join(total_models, by = "Community") %>%
  
  #Total_models is fixed for each community — it is the total number of 
  #sub-region models that exist in that community, regardless of whether a concept appears in them or not.
  
  # Step 3: Occurrence = N_models / Total_models  (already 0–1)
  mutate(Occurrence = N_models / Total_models) %>%
  
  #So for example, in Australia there are 4 sub-region models (New South Wales, Queensland, North Australia, Western Australia). 
  # If the concept "Shark Abundance" appears in all 4 of them, its occurrence is 4/4 = 1.0. 
  #If it only appears in 2 of them, its occurrence is 2/4 = 0.5. A concept that appears in just one model gets 0.25.

  # Step 4: min-max normalise Avg_Degree within each community
  #         to produce Norm_Centrality on a 0–1 scale
  
  # Normalize Centrality to 0–1. Because degree centrality is not naturally bounded between 0 and 1 (unlike occurrence), 
  # the code applies min-max normalization within each community: Norm_Centrality = (Avg_Degree - min) / (max - min). 
  # This rescales the lowest-degree concept to 0 and the highest to 1, making it directly comparable to occurrence. 
  # This is exactly what the paper describes when it says centrality is normalized to create a 0–1 scale.

  group_by(Community) %>%
  mutate(
    Norm_Centrality = {
      mn <- min(Avg_Degree, na.rm = TRUE)
      mx <- max(Avg_Degree, na.rm = TRUE)
      if (mx == mn) rep(0, n())
      else (Avg_Degree - mn) / (mx - mn)
    }
  ) %>%
  ungroup() %>%
  
  # Step 5: Prominence = average of Occurrence & Norm_Centrality
  mutate(
    Prominence = (Occurrence + Norm_Centrality) / 2
  ) %>%
  
  #Step 4 — Average the two components. 
  # Finally, Prominence = (Occurrence + Norm_Centrality) / 2. 
  # This is a straight average of the two equally-weighted 0–1 components, which is precisely the formula Hoffman et al. propose. 
  # A concept scores high on prominence only if it is both widespread across models (high occurrence) 
  # and strongly connected within the models it appears in (high normalized centrality).
  
  # Round for readability
  mutate(across(where(is.numeric), ~ round(.x, 4))) %>%
  
  # Sort by Community then Prominence descending
  arrange(Community, desc(Prominence))


# ── SECTION 4 — PRINT & SAVE ─────────────────────────────────

cat("\n=== PROMINENCE RESULTS ===\n")
for (comm in c("Australia", "Alabama", "Galveston")) {
  cat(sprintf("\n--- %s ---\n", comm))
  sub <- filter(prominence_df, Community == comm)
  print(
    sub %>%
      select(Concept, N_models, Total_models,
             Occurrence, Avg_Degree,
             Norm_Centrality, Prominence),
    n = Inf
  )
}

if (SAVE_CSV)
  write_csv(prominence_df,
            file.path(OUT_DIR, "Prominence_all_communities.csv"))

# Separate CSV per community
if (SAVE_CSV) {
  for (comm in unique(prominence_df$Community)) {
    safe <- str_replace_all(comm, " ", "_")
    write_csv(
      filter(prominence_df, Community == comm),
      file.path(OUT_DIR, sprintf("Prominence_%s.csv", safe))
    )
  }
}

cat(sprintf("\nSaved to: %s\n", OUT_DIR))


# ── SECTION 5 — PLOTS ────────────────────────────────────────

community_colours <- c(
  "Australia" = "#1B9E77",
  "Alabama"   = "#D95F02",
  "Galveston" = "#7570B3"
)

# ── Plot 1: Prominence bar chart — top N per community ────────
plot_df <- prominence_df %>%
  group_by(Community) %>%
  slice_max(Prominence, n = TOP_N, with_ties = FALSE) %>%
  ungroup() %>%
  mutate(
    Concept = tidytext::reorder_within(
      str_wrap(Concept, 30), Prominence, Community
    )
  )


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

p1 <- ggplot(plot_df,
             aes(x = Concept, y = Prominence,
                 fill = Community)) +
  geom_col(width = 0.75, show.legend = FALSE) +
  geom_text(aes(label = sprintf("%.2f", Prominence)),
            hjust = -0.1, size = 2.6, colour = "grey30") +
  coord_flip() +
  tidytext::scale_x_reordered() +
  scale_fill_manual(values = community_colours) +
  scale_y_continuous(limits = c(0, 1.12)) +
  facet_wrap(~ Community, nrow = 1, scales = "free_y") +
  labs(
    title    = "Concept Prominence",
    subtitle = paste0(
      "Top ", TOP_N, " concepts per community | ",
      "Prominence = (Occurrence + Normalised Degree Centrality) / 2"
    ),
    x = NULL, y = "Prominence (0–1)"
  ) +
  theme_fcm()

if (SAVE_PLOTS)
  ggsave(file.path(OUT_DIR, "P1_prominence_top_concepts.pdf"),
         p1, width = 18, height = 9)
print(p1)

png(filename = "P1_prominence_top_concepts.png", 
    width = 17, height = 8, units = "in", # Set size in inches
    res = 600)  
p1
dev.off()


# ── Plot 2: Occurrence vs Centrality scatter ──────────────────
# Shows which concepts are prominent because they are universal
# (high occurrence) vs. because they are central (high centrality)

p2 <- ggplot(prominence_df,
             aes(x = Occurrence, y = Norm_Centrality,
                 colour = Community, size = Prominence)) +
  geom_vline(xintercept = 0.5, linetype = "dashed",
             colour = "grey70", linewidth = 0.4) +
  geom_hline(yintercept = 0.5, linetype = "dashed",
             colour = "grey70", linewidth = 0.4) +
  geom_point(alpha = 0.7) +
  ggrepel::geom_label_repel(
    data = prominence_df %>%
      group_by(Community) %>%
      slice_max(Prominence, n = 8, with_ties = FALSE) %>%
      ungroup(),
    aes(label = str_wrap(Concept, 20)),
    size = 2.3, label.padding = unit(0.1, "lines"),
    max.overlaps = 15, show.legend = FALSE,
    colour = "grey20", fill = alpha("white", 0.8)
  ) +
  scale_colour_manual(values = community_colours) +
  scale_size_continuous(range = c(1.5, 6), guide = "none") +
  scale_x_continuous(labels = percent_format(accuracy = 1),
                     limits = c(0, 1)) +
  scale_y_continuous(limits = c(0, 1)) +
  facet_wrap(~ Community, nrow = 1) +
  annotate("text", x = 0.15, y = 0.92,
           label = "High centrality\nlow occurrence",
           size = 2.6, colour = "grey50", fontface = "italic") +
  annotate("text", x = 0.80, y = 0.08,
           label = "High occurrence\nlow centrality",
           size = 2.6, colour = "grey50", fontface = "italic") +
  annotate("text", x = 0.80, y = 0.92,
           label = "High prominence\n(both)",
           size = 2.6, colour = "grey30", fontface = "bold") +
  labs(
    title    = "Occurrence vs. Normalised Centrality",
    subtitle = "Top-right = high prominence on both dimensions | Size = Prominence score",
    x = "Occurrence (proportion of community models)",
    y = "Normalised degree centrality (0–1)",
    colour = "Community"
  ) +
  theme_fcm()

if (SAVE_PLOTS)
  ggsave(file.path(OUT_DIR, "P2_occurrence_vs_centrality.pdf"),
         p2, width = 16, height = 7)
print(p2)


png(filename = "P2_occurrence_vs_centrality.png", 
    width = 17, height = 8, units = "in", # Set size in inches
    res = 600)  
p2
dev.off()


cat("\nProminence analysis complete.\n")
