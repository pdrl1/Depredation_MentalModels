# ============================================================
# CROSS-MODEL CENTRALITY COMPARISON ANALYSIS
# Regions: Australia (AU) · Alabama (AL) · Galveston (GA)
# ============================================================
#
# Loads all *_centrality.csv files, computes ranks and
# z-scores within each model, and answers six analytical
# questions about how concepts shift across models:
#
#   Q1  Which concepts change role (Transmitter / Receiver /
#         Ordinary) across models?
#   Q2  How do centrality ranks change across models?
#   Q3  Do concepts shift type of importance
#         (driver vs. receiver profile)?
#   Q4  Which concepts are drivers in sub-regions but
#         receivers in the regional full model?
#   Q5  How does aggregation change perceived system
#         structure / eigenvector dominance?
#   Q6  Which concepts are consistently central everywhere?
#
#   WF  Workflow: stable / shifting / emerging concepts
#         + interpretive questions
#
# HOW TO RUN:
#   1. Set CSV_DIR to the folder containing all
#      *_centrality.csv files (AU_full, AL_Texas, etc.)
#   2. Set OUT_DIR for outputs
#   3. source("Cross_Model_Centrality_Comparison.R")
# ============================================================


# ── SECTION 0 — PACKAGES ─────────────────────────────────────

pkgs <- c("dplyr", "tidyr", "ggplot2", "scales",
          "stringr", "purrr", "forcats", "tibble",
          "RColorBrewer", "ggrepel", "readr")
for (p in pkgs) {
  if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
  library(p, character.only = TRUE)
}


# ── SECTION 1 — PARAMETERS ───────────────────────────────────

# List every folder that contains *_centrality.csv files
CSV_DIRS <- c(
  "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Australia/Network_Outputs",
  "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Alabama/Network_Outputs",
  "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Network_Outputs"
)

# Output folder
OUT_DIR <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Centrality_Outputs"

SAVE_CSV   <- TRUE
SAVE_PLOTS <- TRUE

# "Central" = top X% of Degree rank within a model
CENTRAL_PCT <- 0.25

# Top N concepts shown in rank heatmaps
TOP_N <- 20

# Minimum number of models a concept must appear in
# to be included in cross-model comparisons
MIN_MODELS <- 2

# ── Optional: harmonize concept names across regions ─────────
# Add pairs here if the same real-world concept has different
# names in different regions. The RIGHT-hand value becomes the
# canonical name used in cross-region plots.
# Example: GA uses "Shark Depredation"; AU and AL use "Depredation"
#
HARMONIZE <- c(
  "Shark Depredation" = "Depredation"
  # "Old name" = "Canonical name"
)


# ── SECTION 2 — LOAD & PARSE ALL CSVs ────────────────────────

if (!dir.exists(OUT_DIR)) dir.create(OUT_DIR, recursive = TRUE)

csv_files <- unlist(lapply(CSV_DIRS, function(d) {
  list.files(d, pattern = "_centrality\\.csv$", full.names = TRUE)
}))

if (length(csv_files) == 0)
  stop("No *_centrality.csv files found in CSV_DIR. Check path.")

# Parse filename → Region, Model, Scope
parse_filename <- function(path) {
  base  <- tools::file_path_sans_ext(basename(path))
  base  <- sub("_centrality$", "", base)
  parts <- str_split(base, "_", n = 2)[[1]]
  region <- parts[1]
  model  <- if (length(parts) == 2) str_replace_all(parts[2], "_", " ")
  else "full"
  scope  <- if (tolower(model) == "full") "full" else "sub"
  list(region = region, model = model, scope = scope)
}

master_df <- map_dfr(csv_files, function(f) {
  info <- parse_filename(f)
  df   <- read_csv(f, show_col_types = FALSE) %>%
    rename_with(str_to_title) %>%          # normalise column case
    rename(any_of(c(                       # tolerate minor name variants
      Concept     = "Label",
      Betweenness = "Betweenness_centrality",
      Eigenvector = "Eigenvector_centrality"
    )))
  df %>%
    mutate(
      Region   = info$region,
      Model    = info$model,
      Scope    = info$scope,
      Model_ID = paste0(info$region, "_",
                        str_replace_all(info$model, " ", "_"))
    )
}) %>%
  # Apply concept name harmonization
  mutate(Concept = recode(Concept, !!!HARMONIZE))

cat(sprintf("Loaded %d models, %d concept-model rows.\n",
            n_distinct(master_df$Model_ID),
            nrow(master_df)))
cat("Models found:\n")
master_df %>% distinct(Region, Model, Scope) %>%
  arrange(Region, Scope, Model) %>%
  print(n = Inf)


# ── SECTION 3 — RANK & Z-SCORE WITHIN EACH MODEL ─────────────
#
# rank_pct = percentile rank (0–1): rank / n_concepts
# Percentile rank is used for fair cross-model comparison
# because models differ in size (11 to 92 concepts).
# Rank 1 = highest value; rank_pct near 0 = top of model.

metrics <- c("Indegree", "Outdegree", "Degree",
             "Betweenness", "Closeness", "Eigenvector", "Katz")

rank_zscore <- function(df, col) {
  df %>%
    group_by(Model_ID) %>%
    mutate(
      !!paste0(col, "_rank")    := rank(-get(col), ties.method = "min"),
      !!paste0(col, "_rank_pct"):= get(paste0(col, "_rank")) /
        n_distinct(Concept),
      !!paste0(col, "_z")       := {
        v <- get(col)
        s <- sd(v, na.rm = TRUE)
        if (is.na(s) || s == 0) rep(0, n()) else (v - mean(v, na.rm=TRUE)) / s
      }
    ) %>%
    ungroup()
}

for (m in metrics) {
  master_df <- rank_zscore(master_df, m)
}

# Concept "driver profile" — is this concept more driver or receiver?
# driver_score > 0  → outdegree dominates (driver-like)
# driver_score < 0  → indegree dominates  (receiver-like)
master_df <- master_df %>%
  mutate(
    Driver_score  = Outdegree - Indegree,
    Driver_profile = case_when(
      Driver_score >  0.1 * Degree ~ "Driver",
      Driver_score < -0.1 * Degree ~ "Receiver-like",
      TRUE                          ~ "Balanced"
    )
  )

if (SAVE_CSV)
  write_csv(master_df, file.path(OUT_DIR, "00_master_centrality.csv"))

cat("\nMaster table saved: 00_master_centrality.csv\n")


# ── SECTION 4 — Q1: ROLE SHIFTS ──────────────────────────────
#
# "Which concepts change role (Transmitter / Receiver /
#  Ordinary) across models?"
#
# Strategy:
# (a) Within-region: compare each sub-region Type with full model Type
# (b) Cross-region: compare full model Types across AU / AL / GA

cat("\n================================================================\n")
cat("Q1 — ROLE SHIFTS\n")
cat("================================================================\n")

# (a) Within-region role shifts --------------------------------
within_role <- master_df %>%
  group_by(Region, Concept) %>%
  filter(n_distinct(Model_ID) >= 2) %>%
  summarise(
    Types_observed = paste(Model_ID, Type, sep = "=", collapse = " | "),
    N_types        = n_distinct(Type),
    .groups = "drop"
  ) %>%
  filter(N_types > 1) %>%
  arrange(Region, desc(N_types))

cat("\nConcepts that CHANGE role within their region",
    "(sub-regions vs. full model):\n")
print(within_role, n = Inf)

# (b) Cross-region role shifts (full models only) ---------------
cross_role <- master_df %>%
  filter(Scope == "full") %>%
  group_by(Concept) %>%
  filter(n_distinct(Region) >= 2) %>%
  summarise(
    Regions_present  = paste(sort(unique(Region)), collapse = " | "),
    Types_observed   = paste(Region, Type, sep = "=", collapse = " | "),
    N_types          = n_distinct(Type),
    .groups = "drop"
  ) %>%
  filter(N_types > 1) %>%
  arrange(desc(N_types))

cat("\nConcepts that CHANGE role across regions (full models):\n")
print(cross_role, n = Inf)

if (SAVE_CSV) {
  write_csv(within_role, file.path(OUT_DIR, "Q1a_within_region_role_shifts.csv"))
  write_csv(cross_role,  file.path(OUT_DIR, "Q1b_cross_region_role_shifts.csv"))
}


# ── SECTION 5 — Q2: RANK DIFFERENCES ─────────────────────────
#
# "How do centrality ranks change across models?"
#
# For each concept present in ≥ MIN_MODELS models, compute:
#   - Degree rank in each model (percentile)
#   - Mean, min, max rank_pct
#   - Rank range (max_pct - min_pct) = volatility
#   - Biggest single jump: max - min rank_pct

cat("\n================================================================\n")
cat("Q2 — RANK DIFFERENCES\n")
cat("================================================================\n")

# Wide rank table: concepts × models
rank_wide <- master_df %>%
  select(Concept, Model_ID, Degree_rank_pct) %>%
  pivot_wider(names_from = Model_ID,
              values_from = Degree_rank_pct) %>%
  arrange(Concept)

# Summary statistics on rank volatility
rank_summary <- master_df %>%
  group_by(Concept) %>%
  filter(n() >= MIN_MODELS) %>%
  summarise(
    N_models      = n(),
    Regions       = paste(sort(unique(Region)), collapse = "|"),
    Mean_rank_pct = round(mean(Degree_rank_pct),  3),
    SD_rank_pct   = round(sd(Degree_rank_pct),    3),
    Min_rank_pct  = round(min(Degree_rank_pct),   3),
    Max_rank_pct  = round(max(Degree_rank_pct),   3),
    Rank_range    = round(Max_rank_pct - Min_rank_pct, 3),
    Best_model    = Model_ID[which.min(Degree_rank_pct)],
    Worst_model   = Model_ID[which.max(Degree_rank_pct)],
    .groups = "drop"
  ) %>%
  arrange(desc(Rank_range))

cat("\nTop 20 concepts by rank VOLATILITY (Degree, percentile):\n")
print(head(rank_summary, 20))

cat("\nTop 20 most STABLE concepts (lowest rank range):\n")
print(head(arrange(rank_summary, Rank_range), 20))

if (SAVE_CSV) {
  write_csv(rank_wide,    file.path(OUT_DIR, "Q2a_rank_wide_table.csv"))
  write_csv(rank_summary, file.path(OUT_DIR, "Q2b_rank_volatility.csv"))
}


# ── SECTION 6 — Q3: TYPE OF IMPORTANCE SHIFTS ────────────────
#
# "Do concepts change type of importance across regions?"
# A concept might be a strong DRIVER (outdegree > indegree) in
# one model but a strong RECEIVER (indegree > outdegree) in another.

cat("\n================================================================\n")
cat("Q3 — TYPE OF IMPORTANCE SHIFTS\n")
cat("================================================================\n")

importance_profile <- master_df %>%
  group_by(Concept) %>%
  filter(n() >= MIN_MODELS) %>%
  summarise(
    N_models          = n(),
    Regions           = paste(sort(unique(Region)), collapse = "|"),
    Driver_profiles   = paste(Model_ID, Driver_profile,
                              sep = "=", collapse = " | "),
    N_profiles        = n_distinct(Driver_profile),
    Profile_shift     = N_profiles > 1,
    # Indegree vs Outdegree rank dominance per model
    IndegRank_mean    = round(mean(Indegree_rank_pct),  3),
    OutdegRank_mean   = round(mean(Outdegree_rank_pct), 3),
    # Where does indegree dominate?
    Indeg_dominant_in = paste(Model_ID[Indegree_rank_pct <
                                         Outdegree_rank_pct],
                              collapse = " | "),
    # Where does outdegree dominate?
    Outdeg_dominant_in= paste(Model_ID[Outdegree_rank_pct <
                                         Indegree_rank_pct],
                              collapse = " | "),
    .groups = "drop"
  ) %>%
  filter(Profile_shift) %>%
  arrange(desc(N_profiles), desc(N_models))

cat("\nConcepts that SHIFT type of importance",
    "(driver ↔ receiver profile):\n")
print(importance_profile, n = 30)

if (SAVE_CSV)
  write_csv(importance_profile,
            file.path(OUT_DIR, "Q3_importance_profile_shifts.csv"))


# ── SECTION 7 — Q4: DRIVERS → RECEIVERS ──────────────────────
#
# "Which key concepts shift from drivers in sub-regional FCMs
#  to receivers in the regional model?"
#
# Focus: within-region comparison only.
# A "driver" in a sub-region = Transmitter OR Driver_profile=="Driver"
# A "receiver" in the full model = Receiver OR Driver_profile=="Receiver-like"

cat("\n================================================================\n")
cat("Q4 — SUB-REGION DRIVERS → FULL-MODEL RECEIVERS\n")
cat("================================================================\n")

# Sub-region rows where concept is driver
sub_drivers <- master_df %>%
  filter(Scope == "sub") %>%
  filter(Type == "Transmitter" | Driver_profile == "Driver") %>%
  select(Region, Concept, Model_ID,
         Type_sub    = Type,
         Profile_sub = Driver_profile,
         Outdegree_sub = Outdegree,
         Indegree_sub  = Indegree,
         Degree_rank_pct_sub = Degree_rank_pct)

# Full-model rows where concept is receiver-like
full_receivers <- master_df %>%
  filter(Scope == "full") %>%
  filter(Type == "Receiver" | Driver_profile == "Receiver-like") %>%
  select(Region, Concept,
         Type_full    = Type,
         Profile_full = Driver_profile,
         Outdegree_full = Outdegree,
         Indegree_full  = Indegree,
         Degree_rank_pct_full = Degree_rank_pct)

driver_to_receiver <- inner_join(sub_drivers, full_receivers,
                                 by = c("Region", "Concept")) %>%
  arrange(Region, Concept)

cat("\nConcepts that are DRIVERS in sub-regions but",
    "RECEIVERS in full model:\n")
print(driver_to_receiver %>%
        select(Region, Concept, Model_ID,
               Type_sub, Type_full,
               Outdegree_sub, Indegree_sub,
               Outdegree_full, Indegree_full),
      n = Inf)

cat("\nInterpretation hint:\n")
cat("  These concepts are perceived as CAUSES at the local scale\n")
cat("  but as OUTCOMES at the regional scale — suggesting that\n")
cat("  their causal power is not consistently recognized across\n")
cat("  participant groups, or that aggregation dilutes their\n")
cat("  outgoing signal relative to incoming connections.\n")

if (SAVE_CSV)
  write_csv(driver_to_receiver,
            file.path(OUT_DIR, "Q4_driver_to_receiver_shifts.csv"))


# ── SECTION 8 — Q5: AGGREGATION & EIGENVECTOR DOMINANCE ──────
#
# "How does aggregation change perceived system structure?"
# "Do some sub-regions dominate eigenvector centrality?"
#
# Approach:
# (a) Compare Eigenvector rank_pct in sub-region vs. full model
# (b) For each concept, note which sub-region gives it its
#     highest eigenvector rank, and whether that high rank
#     persists in the full model

cat("\n================================================================\n")
cat("Q5 — AGGREGATION & EIGENVECTOR DOMINANCE\n")
cat("================================================================\n")

# (a) Within-region eigenvector shift sub → full -----------------
eig_shift <- master_df %>%
  group_by(Region, Concept) %>%
  filter(any(Scope == "full") & any(Scope == "sub")) %>%
  summarise(
    Eig_full        = Eigenvector[Scope == "full"][1],
    Eig_rank_pct_full = Eigenvector_rank_pct[Scope == "full"][1],
    Eig_rank_pct_sub_mean = mean(Eigenvector_rank_pct[Scope == "sub"],
                                 na.rm = TRUE),
    Eig_rank_pct_sub_best = min(Eigenvector_rank_pct[Scope == "sub"],
                                na.rm = TRUE),   # lowest pct = best rank
    Best_sub_model  = {
      subs <- Model_ID[Scope == "sub"]
      eigs <- Eigenvector_rank_pct[Scope == "sub"]
      if (length(eigs) == 0) NA_character_
      else subs[which.min(eigs)]
    },
    Eig_shift       = round(Eig_rank_pct_full - Eig_rank_pct_sub_best, 3),
    # Positive = concept LOST eigenvector rank going to full model
    # Negative = concept GAINED eigenvector rank in full model
    .groups = "drop"
  ) %>%
  arrange(Region, desc(abs(Eig_shift)))

cat("\nTop concepts with biggest Eigenvector rank SHIFT",
    "(sub-region → full model):\n")
cat("  Positive Eig_shift = concept lost rank in aggregated model\n")
cat("  Negative Eig_shift = concept gained rank in aggregated model\n\n")
print(head(eig_shift, 30))

# (b) Which sub-regions dominate eigenvector in the full model? --
# For each concept that is top CENTRAL_PCT in full model eigenvector,
# which sub-region gave it its highest eigenvector?
top_eig_full <- master_df %>%
  filter(Scope == "full",
         Eigenvector_rank_pct <= CENTRAL_PCT) %>%
  select(Region, Concept, Eig_full = Eigenvector,
         Eig_rank_pct_full = Eigenvector_rank_pct)

eig_dominance <- eig_shift %>%
  inner_join(top_eig_full, by = c("Region", "Concept")) %>%
  count(Region, Best_sub_model, name = "N_top_concepts") %>%
  arrange(Region, desc(N_top_concepts))

cat("\nWhich sub-regions contribute most top-eigenvector",
    "concepts to the full model?\n")
print(eig_dominance)

# (c) Cross-region: do some REGIONS dominate eigenvector? --------
cross_eig <- master_df %>%
  filter(Scope == "full") %>%
  group_by(Concept) %>%
  filter(n_distinct(Region) >= 2) %>%
  summarise(
    Best_eig_region = Region[which.min(Eigenvector_rank_pct)],
    Eig_rank_pct    = paste(Region, round(Eigenvector_rank_pct,2),
                            sep="=", collapse=" | "),
    .groups = "drop"
  ) %>%
  count(Best_eig_region, name = "N_concepts") %>%
  arrange(desc(N_concepts))

cat("\nAcross shared concepts (full models), which region",
    "gives highest Eigenvector rank?\n")
print(cross_eig)

if (SAVE_CSV) {
  write_csv(eig_shift,     file.path(OUT_DIR, "Q5a_eigenvector_shift.csv"))
  write_csv(eig_dominance, file.path(OUT_DIR, "Q5b_eigenvector_dominance.csv"))
}


# ── SECTION 9 — Q6: CONSISTENTLY CENTRAL CONCEPTS ────────────
#
# "Which concepts are consistently central everywhere?"
#
# A concept is "consistently central" in a given model if its
# Degree rank_pct ≤ CENTRAL_PCT (e.g., top 25%).
# Stability score = N models where central / N models appeared in.

cat("\n================================================================\n")
cat("Q6 — CONSISTENTLY CENTRAL CONCEPTS\n")
cat("================================================================\n")

consistency <- master_df %>%
  group_by(Concept) %>%
  filter(n() >= MIN_MODELS) %>%
  summarise(
    N_models         = n(),
    Regions          = paste(sort(unique(Region)), collapse = "|"),
    Scopes           = paste(sort(unique(Scope)),  collapse = "|"),
    N_central        = sum(Degree_rank_pct <= CENTRAL_PCT, na.rm = TRUE),
    Stability_score  = round(N_central / N(),  3),
    Mean_degree_rank = round(mean(Degree_rank_pct), 3),
    Mean_degree_z    = round(mean(Degree_z,    na.rm = TRUE), 3),
    .groups = "drop"
  ) %>%
  arrange(desc(Stability_score), Mean_degree_rank)

cat(sprintf("\nConcepts central (top %d%%) in ALL models they appear in:\n",
            round(CENTRAL_PCT * 100)))
perfectly_stable <- filter(consistency, Stability_score == 1, N_models >= MIN_MODELS)
print(perfectly_stable, n = Inf)

cat("\nTop 20 most consistently central concepts overall:\n")
print(head(consistency, 20))

# Cross-region stable core (appear in ≥ 2 regions, always central)
cross_stable <- consistency %>%
  filter(str_detect(Regions, "\\|"),   # ≥ 2 regions
         Stability_score >= 0.75)

cat("\nCross-region STABLE CORE",
    "(appear in ≥ 2 regions, central ≥ 75% of models):\n")
print(cross_stable, n = Inf)

if (SAVE_CSV) {
  write_csv(consistency,   file.path(OUT_DIR, "Q6a_consistency_scores.csv"))
  write_csv(cross_stable,  file.path(OUT_DIR, "Q6b_cross_region_stable_core.csv"))
}


# ── SECTION 10 — WORKFLOW ────────────────────────────────────
#
# Step-by-step output as described in the research workflow:
# 1. Centrality per model         → 00_master_centrality.csv
# 2. Ranks within model           → (columns *_rank, *_rank_pct)
# 3. Centrality comparison table  → WF_centrality_comparison.csv
# 4. Rank change + z-score diff   → WF_rank_change.csv
# 5. Stable / shifting / emerging → WF_stable / _shifting / _emerging.csv
# 6. Interpretive questions       → WF_interpretive_questions.csv

cat("\n================================================================\n")
cat("WORKFLOW OUTPUTS\n")
cat("================================================================\n")

# ── Step 3: Centrality comparison table ───────────────────────
# Pivot: one row per concept, columns = Degree rank_pct per model
comp_table <- master_df %>%
  select(Concept, Region, Model_ID, Scope,
         Type, Degree_rank_pct, Degree_z,
         Indegree_rank_pct, Outdegree_rank_pct,
         Eigenvector_rank_pct, Betweenness_rank_pct) %>%
  pivot_wider(
    id_cols     = Concept,
    names_from  = Model_ID,
    values_from = c(Type, Degree_rank_pct, Degree_z,
                    Indegree_rank_pct, Outdegree_rank_pct,
                    Eigenvector_rank_pct)
  )

if (SAVE_CSV)
  write_csv(comp_table, file.path(OUT_DIR, "WF_centrality_comparison_wide.csv"))

# ── Step 4: Rank change (sub-region → full model) ─────────────
rank_change <- master_df %>%
  filter(Scope %in% c("full", "sub")) %>%
  group_by(Region, Concept) %>%
  filter(any(Scope == "full"), any(Scope == "sub")) %>%
  summarise(
    Full_degree_rank_pct  = Degree_rank_pct[Scope == "full"][1],
    Sub_degree_rank_pct   = mean(Degree_rank_pct[Scope == "sub"]),
    Rank_change           = round(Full_degree_rank_pct -
                                    Sub_degree_rank_pct, 3),
    # Positive = dropped in rank going to full model
    # Negative = rose in rank going to full model
    Full_z    = Degree_z[Scope == "full"][1],
    Sub_z_avg = mean(Degree_z[Scope == "sub"]),
    Z_diff    = round(Full_z - Sub_z_avg, 3),
    Full_type = Type[Scope == "full"][1],
    Sub_types = paste(unique(Type[Scope == "sub"]), collapse="|"),
    .groups = "drop"
  ) %>%
  arrange(Region, desc(abs(Rank_change)))

if (SAVE_CSV)
  write_csv(rank_change, file.path(OUT_DIR, "WF_rank_change.csv"))

cat("\nTop 15 concepts by rank change (sub-region → full model):\n")
cat("  Positive = concept dropped in rank going to full model\n")
cat("  Negative = concept rose in rank going to full model\n\n")
print(head(arrange(rank_change, desc(abs(Rank_change))), 15))

# ── Step 5: Stable / Shifting / Emerging ──────────────────────

# STABLE: Stability_score = 1.0 AND appears in ≥ MIN_MODELS models
stable <- consistency %>%
  filter(Stability_score == 1, N_models >= MIN_MODELS) %>%
  arrange(Mean_degree_rank)

# SHIFTING: high rank range (volatile)
shifting <- rank_summary %>%
  filter(Rank_range >= quantile(Rank_range, 0.75, na.rm = TRUE)) %>%
  arrange(desc(Rank_range))

# EMERGING: concepts present in only 1 region's full model
#           but top CENTRAL_PCT in that model
emerging <- master_df %>%
  filter(Scope == "full") %>%
  group_by(Concept) %>%
  filter(n_distinct(Region) == 1,
         Degree_rank_pct <= CENTRAL_PCT) %>%
  summarise(
    Region         = unique(Region),
    Model_ID       = unique(Model_ID),
    Type           = unique(Type),
    Degree_rank_pct = round(mean(Degree_rank_pct), 3),
    .groups = "drop"
  ) %>%
  arrange(Region, Degree_rank_pct)

# DISAPPEARING: concepts in sub-region top 25% but bottom 50% in full model
disappearing <- rank_change %>%
  filter(Sub_degree_rank_pct <= CENTRAL_PCT,
         Full_degree_rank_pct > 0.50) %>%
  arrange(Region, desc(Rank_change))

cat("\nSTABLE concepts (top 25% in every model they appear in):\n")
print(stable %>% select(Concept, Regions, N_models, Mean_degree_rank),
      n = Inf)

cat("\nSHIFTING concepts (highest rank volatility):\n")
print(head(shifting %>% select(Concept, Regions, N_models,
                               Mean_rank_pct, Rank_range,
                               Best_model, Worst_model), 20))

cat("\nEMERGING concepts (prominent in ONE region's full model only):\n")
print(head(emerging, 20))

cat("\nDISAPPEARING concepts (top 25% in sub-region,",
    "bottom 50% in full model):\n")
print(disappearing %>%
        select(Region, Concept,
               Sub_degree_rank_pct, Full_degree_rank_pct,
               Rank_change, Full_type, Sub_types),
      n = Inf)

if (SAVE_CSV) {
  write_csv(stable,       file.path(OUT_DIR, "WF_stable_concepts.csv"))
  write_csv(shifting,     file.path(OUT_DIR, "WF_shifting_concepts.csv"))
  write_csv(emerging,     file.path(OUT_DIR, "WF_emerging_concepts.csv"))
  write_csv(disappearing, file.path(OUT_DIR, "WF_disappearing_concepts.csv"))
}

# ── Step 6: Interpretive questions ────────────────────────────
# Auto-generate questions from data patterns

questions <- c()

# From stable core
if (nrow(cross_stable) > 0) {
  top_stable <- paste(head(cross_stable$Concept, 3), collapse = ", ")
  questions <- c(questions,
                 paste0("Why are [", top_stable, "] consistently central across",
                        " all regions and sub-regions? Are they fundamental",
                        " system hubs or artefacts of a shared interview instrument?")
  )
}

# From shifting concepts
if (nrow(shifting) > 0) {
  top_shifting <- paste(head(shifting$Concept, 2), collapse = " and ")
  questions <- c(questions,
                 paste0("What explains why [", top_shifting, "] are central",
                        " in some models but peripheral in others?",
                        " Is this a genuine regional difference in perceived",
                        " importance, or a data collection artefact?")
  )
}

# From driver→receiver shift
if (nrow(driver_to_receiver) > 0) {
  dr <- head(driver_to_receiver$Concept, 2)
  reg <- head(driver_to_receiver$Region, 2)
  questions <- c(questions,
                 paste0("In region [", paste(unique(reg), collapse="|"), "],",
                        " concept(s) [", paste(unique(dr), collapse=", "), "]",
                        " act as drivers at the local scale but receivers at",
                        " the regional scale. Does this reflect genuine",
                        " scale-dependency, or is it because local participants",
                        " have more perceived control than regional ones?")
  )
}

# From role shifts
if (nrow(within_role) > 0) {
  top_role <- paste(head(within_role$Concept, 2), collapse=" and ")
  questions <- c(questions,
                 paste0("Concept(s) [", top_role, "] change role",
                        " (Transmitter/Ordinary/Receiver) across sub-regions.",
                        " Is this because different participant groups have",
                        " different mental models, or because the concept",
                        " genuinely plays a different structural role",
                        " at different spatial scales?")
  )
}

# From eigenvector dominance
if (nrow(eig_dominance) > 0) {
  dom_sub <- eig_dominance %>% slice(1)
  questions <- c(questions,
                 paste0("Sub-region [", dom_sub$Best_sub_model[1],
                        "] contributes the most top-eigenvector concepts",
                        " to the full regional model. Does this sub-region",
                        " disproportionately shape the perceived system",
                        " structure, and if so, why?")
  )
}

# From emerging concepts
if (nrow(emerging) > 0) {
  em_concepts <- paste(head(emerging$Concept, 3), collapse=", ")
  questions <- c(questions,
                 paste0("Concept(s) [", em_concepts, "] are central",
                        " in one region but absent or peripheral elsewhere.",
                        " Are these region-specific phenomena, or are they",
                        " under-reported in other regions?")
  )
}

q_df <- data.frame(
  N        = seq_along(questions),
  Question = questions,
  stringsAsFactors = FALSE
)

cat("\n================================================================\n")
cat("INTERPRETIVE QUESTIONS GENERATED FROM DATA PATTERNS\n")
cat("================================================================\n")
for (i in seq_along(questions)) {
  cat(sprintf("\n[Q%d] %s\n", i, questions[i]))
}

if (SAVE_CSV)
  write_csv(q_df, file.path(OUT_DIR, "WF_interpretive_questions.csv"))


# ── SECTION 11 — PLOTS ───────────────────────────────────────

cat("\n--- Generating plots ---\n")

region_colors <- c(
  "AU" = "#1B9E77",
  "AL" = "#D95F02",
  "GA" = "#7570B3"
)

scope_shapes <- c("full" = 16, "sub" = 21)

# ── Plot 1: Rank heatmap — top concepts across all models ──────
# Select top TOP_N concepts by mean Degree z-score

top_concepts <- master_df %>%
  group_by(Concept) %>%
  summarise(mean_z = mean(Degree_z, na.rm = TRUE),
            n_mods = n(), .groups = "drop") %>%
  filter(n_mods >= MIN_MODELS) %>%
  slice_max(mean_z, n = TOP_N) %>%
  pull(Concept)

heat_df <- master_df %>%
  filter(Concept %in% top_concepts) %>%
  mutate(
    Concept  = factor(Concept, levels = rev(top_concepts)),
    Model_ID = factor(Model_ID,
                      levels = sort(unique(Model_ID)))
  )

p1 <- ggplot(heat_df,
             aes(x = Model_ID, y = Concept,
                 fill = Degree_rank_pct)) +
  geom_tile(color = "white", linewidth = 0.4) +
  geom_text(aes(label = sprintf("%.0f%%",
                                Degree_rank_pct * 100)),
            size = 2.2, color = "grey20") +
  scale_fill_gradient2(
    low = "#1a9641", mid = "#ffffbf", high = "#d7191c",
    midpoint = 0.5,
    labels = percent,
    name = "Degree rank\n(0% = top)"
  ) +
  facet_grid(. ~ Region,
             scales = "free_x", space = "free_x") +
  labs(
    title    = "Degree Centrality Rank Across All Models",
    subtitle = paste0("Top ", TOP_N, " concepts by mean z-score | ",
                      "Cell = percentile rank within model (green = top)"),
    x = NULL, y = NULL
  ) +
  theme_bw(base_size = 9) +
  theme(
    axis.text.x  = element_text(angle = 40, hjust = 1, size = 7),
    axis.text.y  = element_text(size = 7),
    strip.text   = element_text(face = "bold"),
    legend.position = "right"
  )

if (SAVE_PLOTS)
  ggsave(file.path(OUT_DIR, "P1_rank_heatmap_all_models.pdf"),
         p1, width = max(14, n_distinct(heat_df$Model_ID) * 0.8),
         height = 9)
print(p1)


# ── Plot 2: Role shifts — stacked bar per model ────────────────
type_order <- c("Transmitter", "Ordinary", "Isolated", "Receiver")
type_colors <- c(
  "Transmitter" = "#4575b4",
  "Ordinary"    = "#74add1",
  "Isolated"    = "#e0e0e0",
  "Receiver"    = "#d73027"
)

type_df <- master_df %>%
  mutate(Type = factor(Type, levels = type_order)) %>%
  count(Region, Model_ID, Scope, Type) %>%
  group_by(Model_ID) %>%
  mutate(pct = n / sum(n))

p2 <- ggplot(type_df, aes(x = Model_ID, y = pct,
                          fill = Type)) +
  geom_col(position = "fill") +
  geom_text(aes(label = ifelse(pct > 0.05,
                               paste0(n), "")),
            position = position_fill(vjust = 0.5),
            size = 2.8, color = "white", fontface = "bold") +
  scale_y_continuous(labels = percent) +
  scale_fill_manual(values = type_colors) +
  facet_grid(. ~ Region, scales = "free_x", space = "free_x") +
  labs(
    title    = "Q1 — Role Composition per Model",
    subtitle = "Transmitters (blue) = drivers | Receivers (red) = outcomes",
    x = NULL, y = "Proportion"
  ) +
  theme_bw(base_size = 9) +
  theme(
    axis.text.x  = element_text(angle = 40, hjust = 1, size = 7),
    strip.text   = element_text(face = "bold"),
    legend.position = "bottom"
  )

if (SAVE_PLOTS)
  ggsave(file.path(OUT_DIR, "P2_role_composition.pdf"),
         p2, width = 12, height = 5)
print(p2)


# ── Plot 3: Driver profile shift — Indegree vs Outdegree ───────
# Scatter: x = Indegree_rank_pct, y = Outdegree_rank_pct
# Points below diagonal = Outdegree dominates (driver)
# Points above diagonal = Indegree dominates (receiver)
# Connect same concept across models with line

profile_df <- master_df %>%
  group_by(Concept) %>%
  filter(n() >= MIN_MODELS) %>%
  ungroup()

# Label only the top-shifting concepts
label_concepts <- importance_profile %>%
  head(min(10, nrow(importance_profile))) %>%
  pull(Concept)

p3 <- ggplot(profile_df,
             aes(x = Indegree_rank_pct,
                 y = Outdegree_rank_pct,
                 color = Region,
                 shape = Scope)) +
  geom_abline(slope = 1, intercept = 0,
              linetype = "dashed", color = "grey60") +
  geom_point(alpha = 0.6, size = 2) +
  geom_line(aes(group = Concept), alpha = 0.15, color = "grey50") +
  geom_label_repel(
    data = filter(profile_df, Concept %in% label_concepts),
    aes(label = Concept),
    size = 2.4, max.overlaps = 15, label.padding = 0.15
  ) +
  annotate("text", x = 0.05, y = 0.95, label = "Receiver-like",
           size = 3, color = "grey40", fontface = "italic") +
  annotate("text", x = 0.95, y = 0.05, label = "Driver-like",
           size = 3, color = "grey40", fontface = "italic") +
  scale_color_manual(values = region_colors) +
  scale_shape_manual(values = scope_shapes) +
  scale_x_continuous(labels = percent, limits = c(0, 1)) +
  scale_y_continuous(labels = percent, limits = c(0, 1)) +
  labs(
    title    = "Q3 — Type of Importance: Driver vs. Receiver Profile",
    subtitle = "Below diagonal = Outdegree dominates (driver) | Above = Indegree dominates (receiver)\nLines connect same concept across models",
    x = "Indegree rank percentile (0% = highest indegree)",
    y = "Outdegree rank percentile (0% = highest outdegree)"
  ) +
  theme_bw(base_size = 9) +
  theme(legend.position = "right")

if (SAVE_PLOTS)
  ggsave(file.path(OUT_DIR, "P3_driver_receiver_profile.pdf"),
         p3, width = 10, height = 8)
print(p3)


# ── Plot 4: Rank change sub-region → full model ────────────────
rank_change_plot <- rank_change %>%
  group_by(Region) %>%
  slice_max(abs(Rank_change), n = 15) %>%
  ungroup() %>%
  mutate(
    Direction = ifelse(Rank_change > 0, "Dropped in full", "Rose in full"),
    Concept_label = str_wrap(Concept, 25)
  )

p4 <- ggplot(rank_change_plot,
             aes(x = reorder(Concept_label, -Rank_change),
                 y = Rank_change,
                 fill = Direction)) +
  geom_col() +
  geom_hline(yintercept = 0, color = "grey30") +
  coord_flip() +
  scale_fill_manual(values = c("Dropped in full" = "#d73027",
                               "Rose in full"   = "#4575b4")) +
  scale_y_continuous(labels = function(x) paste0(round(x*100), "pp")) +
  facet_wrap(~Region, scales = "free_y", ncol = 1) +
  labs(
    title    = "Q2 & Q4 — Rank Change: Sub-Region Average → Full Model",
    subtitle = "Positive = concept dropped in rank | Negative = rose in rank\n(percentage-point change in percentile rank)",
    x = NULL, y = "Rank percentile change",
    fill = NULL
  ) +
  theme_bw(base_size = 9) +
  theme(legend.position = "top",
        strip.text = element_text(face = "bold"))

if (SAVE_PLOTS)
  ggsave(file.path(OUT_DIR, "P4_rank_change_sub_to_full.pdf"),
         p4, width = 10,
         height = 4 + n_distinct(rank_change_plot$Region) * 4)
print(p4)


# ── Plot 5: Eigenvector comparison sub-region vs. full ─────────
eig_plot_df <- eig_shift %>%
  group_by(Region) %>%
  slice_max(abs(Eig_shift), n = 12) %>%
  ungroup() %>%
  mutate(
    Direction = ifelse(Eig_shift > 0,
                       "Lost eig. rank in full",
                       "Gained eig. rank in full"),
    Concept_label = str_wrap(Concept, 25)
  )

p5 <- ggplot(eig_plot_df,
             aes(x = reorder(Concept_label, -Eig_shift),
                 y = Eig_shift,
                 fill = Direction)) +
  geom_col() +
  geom_hline(yintercept = 0, color = "grey30") +
  coord_flip() +
  scale_fill_manual(values = c(
    "Lost eig. rank in full"   = "#d73027",
    "Gained eig. rank in full" = "#4575b4"
  )) +
  scale_y_continuous(labels = function(x) paste0(round(x*100), "pp")) +
  facet_wrap(~Region, scales = "free_y", ncol = 1) +
  labs(
    title    = "Q5 — Eigenvector Rank Shift: Sub-Region → Full Model",
    subtitle = "Positive = concept lost influence position in aggregated model\nNegative = concept emerged as influential through aggregation",
    x = NULL, y = "Eigenvector rank percentile change", fill = NULL
  ) +
  theme_bw(base_size = 9) +
  theme(legend.position = "top",
        strip.text = element_text(face = "bold"))

if (SAVE_PLOTS)
  ggsave(file.path(OUT_DIR, "P5_eigenvector_shift.pdf"),
         p5, width = 10,
         height = 4 + n_distinct(eig_plot_df$Region) * 4)
print(p5)


# ── Plot 6: Stability score — consistently central concepts ────
p6_df <- consistency %>%
  filter(N_models >= MIN_MODELS) %>%
  slice_max(Stability_score, n = 30) %>%
  mutate(
    Regions_label = Regions,
    Concept = str_wrap(Concept, 30),
    Bar_color = case_when(
      Stability_score == 1 ~ "Always central",
      Stability_score >= 0.67 ~ "Often central",
      TRUE ~ "Sometimes central"
    )
  )

p6 <- ggplot(p6_df,
             aes(x = reorder(Concept, Stability_score),
                 y = Stability_score,
                 fill = Bar_color)) +
  geom_col() +
  geom_text(aes(label = sprintf("%d/%d", N_central, N_models)),
            hjust = -0.1, size = 2.8) +
  coord_flip() +
  scale_y_continuous(labels = percent, limits = c(0, 1.15)) +
  scale_fill_manual(values = c(
    "Always central"    = "#1a9641",
    "Often central"     = "#74add1",
    "Sometimes central" = "#fdae61"
  )) +
  labs(
    title    = "Q6 — Concept Stability Across Models",
    subtitle = paste0("Proportion of models (n ≥ ", MIN_MODELS,
                      ") where concept is in top ", CENTRAL_PCT*100,
                      "% by Degree | label = central/appeared"),
    x = NULL, y = "Stability score", fill = NULL
  ) +
  theme_bw(base_size = 9) +
  theme(legend.position = "top",
        axis.text.y = element_text(size = 7))

if (SAVE_PLOTS)
  ggsave(file.path(OUT_DIR, "P6_stability_scores.pdf"),
         p6, width = 10, height = 10)
print(p6)


# ── DONE ──────────────────────────────────────────────────────

cat("\n================================================================\n")
cat("ALL OUTPUTS SAVED\n")
cat("================================================================\n")
cat(sprintf("Directory: %s\n\n", OUT_DIR))
cat("CSVs:\n")
cat("  00_master_centrality.csv          — All concepts, all models, all metrics\n")
cat("  Q1a_within_region_role_shifts.csv — Role changes within regions\n")
cat("  Q1b_cross_region_role_shifts.csv  — Role changes across regions\n")
cat("  Q2a_rank_wide_table.csv           — Degree rank pivot table\n")
cat("  Q2b_rank_volatility.csv           — Rank range per concept\n")
cat("  Q3_importance_profile_shifts.csv  — Driver/receiver profile shifts\n")
cat("  Q4_driver_to_receiver_shifts.csv  — Sub-region drivers → full receivers\n")
cat("  Q5a_eigenvector_shift.csv         — Eigenvector rank shifts\n")
cat("  Q5b_eigenvector_dominance.csv     — Which sub-regions dominate eigenvector\n")
cat("  Q6a_consistency_scores.csv        — Stability scores per concept\n")
cat("  Q6b_cross_region_stable_core.csv  — Cross-region stable core\n")
cat("  WF_centrality_comparison_wide.csv — Wide comparison table (workflow step 3)\n")
cat("  WF_rank_change.csv                — Rank change sub → full (workflow step 4)\n")
cat("  WF_stable_concepts.csv            — Always-central concepts\n")
cat("  WF_shifting_concepts.csv          — Highly volatile concepts\n")
cat("  WF_emerging_concepts.csv          — Central in one region only\n")
cat("  WF_disappearing_concepts.csv      — Top sub-region, peripheral full model\n")
cat("  WF_interpretive_questions.csv     — Auto-generated research questions\n\n")
cat("Plots (PDF):\n")
cat("  P1_rank_heatmap_all_models.pdf    — Rank heatmap, top concepts × all models\n")
cat("  P2_role_composition.pdf           — Role (T/O/R) composition per model\n")
cat("  P3_driver_receiver_profile.pdf    — Indegree vs outdegree profile shifts\n")
cat("  P4_rank_change_sub_to_full.pdf    — Rank change sub-region → full model\n")
cat("  P5_eigenvector_shift.pdf          — Eigenvector shift through aggregation\n")
cat("  P6_stability_scores.pdf           — Concept stability scores\n\n")
cat("Analysis complete.\n")