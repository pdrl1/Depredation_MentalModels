# #############################################################################
#
#   GALVESTON SHARK-DEPREDATION FCM — MASTER MANUSCRIPT PIPELINE (self-contained)
#   ===========================================================================
#   One file that reproduces every analysis and figure reported in the
#   manuscript. The FCM engine is defined once at the top; each subsequent
#   SECTION is a self-contained stage (it loads its own models and defines its
#   own settings), so the script runs cleanly top to bottom.
#
#   ---------------------------------------------------------------------------
#   MANUSCRIPT FLOW  (section -> original script it was built from)
#   ---------------------------------------------------------------------------
#     ENGINE ................. FCM inference (Modified Kosko + sigmoid)      [00]
#     SECTION 1  Kumu -> CSV .. build the 61-concept aggregated matrix       [01]
#     SECTION 2  Shark Depredation scenario
#         2a  data + baseline/scenario/change tables                        [03]
#         2b  manuscript figures (lollipop, dumbbell, heatmap, ego-nets)    [09]
#         2c  activation heatmap (change-from-baseline)                     [10]
#     SECTION 3  Thematic scenarios (Ecological / Management / Behaviour)
#         3a  individual + additive scenarios + comparison                  [07]
#         3b  cross-model synthesis figures (depredation + shared outcomes) [08]
#         3c  thematic heatmaps (additive headline + per-model detail)      [11]
#
#   ---------------------------------------------------------------------------
#   HOW TO RUN
#   ---------------------------------------------------------------------------
#     Set the working directory to this "R_scripts" folder, then:
#         setwd(".../Review Analysis Mahi Claude/R_scripts")
#         source("00_MASTER_manuscript_pipeline.R")
#     (paths are relative to this folder: ../data for inputs, ../results for
#      outputs). All required packages are installed on first use if missing:
#      readxl, openxlsx, ggplot2, scales, patchwork.
# #############################################################################

# ── Sanity check: inputs must be reachable relative to this folder ────────────
if (!dir.exists(file.path("..", "data"))) {
  stop("Please setwd() to the 'R_scripts' folder before running (so that ",
       "'../data' and '../results' resolve correctly).")
}

# ── Install / load the packages used across the pipeline ──────────────────────
.pkgs <- c("readxl", "openxlsx", "ggplot2", "scales", "patchwork")
for (.p in .pkgs) if (!requireNamespace(.p, quietly = TRUE))
  install.packages(.p, repos = "https://cloud.r-project.org")
library(ggplot2); library(scales)


# #############################################################################
#   ENGINE  — FCM Simulation Engine (Modified Kosko inference + sigmoid)   [00]
#   Defined ONCE here; every section below uses these functions.
# #############################################################################

# Sigmoid activation function.
sigmoid <- function(x, lam = 1.0) {
  1.0 / (1.0 + exp(-lam * x))
}

# Load a MentalModeler-format CSV adjacency matrix.
# Returns: list(W = numeric matrix, concepts = character vector)
load_adjacency_matrix <- function(filepath) {
  df <- read.csv(filepath, row.names = 1, check.names = FALSE,
                 colClasses = "character", na.strings = "")
  concepts <- rownames(df)
  W <- suppressWarnings(as.matrix(sapply(df, as.numeric)))
  W[is.na(W)] <- 0.0
  W <- matrix(as.numeric(W), nrow = length(concepts),
              ncol = ncol(df), dimnames = list(concepts, colnames(df)))
  list(W = W, concepts = concepts)
}

# Run FCM simulation using Modified Kosko inference.
run_fcm <- function(W, concepts, clamped = NULL, initial_state = 0.5,
                    lam = 1.0, threshold = 0.001, max_iter = 100) {
  n <- length(concepts)
  concept_idx <- setNames(seq_len(n), concepts)
  
  if (is.list(initial_state) || (!is.null(names(initial_state)) && length(initial_state) > 1)) {
    A <- vapply(concepts, function(c) {
      v <- initial_state[[c]]; if (is.null(v)) 0.5 else as.numeric(v)
    }, numeric(1))
  } else {
    A <- rep(as.numeric(initial_state), n)
  }
  names(A) <- concepts
  
  if (!is.null(clamped)) for (c in names(clamped))
    if (!is.na(concept_idx[c]) && c %in% concepts) A[concept_idx[c]] <- clamped[[c]]
  
  history <- list(A)
  for (iteration in seq_len(max_iter)) {
    A_new <- sigmoid(as.vector(A %*% W) + A, lam = lam)   # A @ W + A  (row-vector)
    names(A_new) <- concepts
    if (!is.null(clamped)) for (c in names(clamped))
      if (!is.na(concept_idx[c]) && c %in% concepts) A_new[concept_idx[c]] <- clamped[[c]]
    history[[length(history) + 1]] <- A_new
    if (max(abs(A_new - A)) < threshold) {
      A <- A_new
      return(list(steady_state = as.list(setNames(A, concepts)), history = history,
                  iterations = iteration, converged = TRUE))
    }
    A <- A_new
  }
  list(steady_state = as.list(setNames(A, concepts)), history = history,
       iterations = max_iter, converged = FALSE)
}

# Run a scenario by clamping one or more concepts; returns baseline + change.
run_scenario <- function(W, concepts, activated_concepts, clamp_value = 1.0,
                         initial_state = 0.5, lam = 1.0, threshold = 0.001, max_iter = 100) {
  baseline <- run_fcm(W, concepts, clamped = NULL, initial_state = initial_state,
                      lam = lam, threshold = threshold, max_iter = max_iter)
  clamped <- setNames(as.list(rep(clamp_value, length(activated_concepts))), activated_concepts)
  scenario <- run_fcm(W, concepts, clamped = clamped, initial_state = initial_state,
                      lam = lam, threshold = threshold, max_iter = max_iter)
  change <- list()
  for (c in concepts) change[[c]] <- scenario$steady_state[[c]] - baseline$steady_state[[c]]
  scenario$baseline <- baseline$steady_state
  scenario$change   <- change
  scenario
}

# Convert scenario results to a data.frame sorted by absolute change.
results_to_dataframe <- function(scenario_result, concepts) {
  df <- data.frame(
    Concept  = concepts,
    Baseline = vapply(concepts, function(c) scenario_result$baseline[[c]], numeric(1)),
    Scenario = vapply(concepts, function(c) scenario_result$steady_state[[c]], numeric(1)),
    Change   = vapply(concepts, function(c) scenario_result$change[[c]], numeric(1)),
    stringsAsFactors = FALSE)
  df$Abs_Change <- abs(df$Change)
  df <- df[order(-df$Abs_Change), ]
  rownames(df) <- NULL
  df
}

cat("Engine loaded. Starting pipeline...\n")


# #############################################################################
#   SECTION 1 — KUMU -> CSV : build the 61-concept aggregated matrix       [01]
#   Reads the Kumu workbook, expands multi-sector edges, averages shared edges,
#   and writes ../data/Aggregated_Full_61.csv (the full community model).
# #############################################################################
cat("\n\n=====================  SECTION 1  Kumu -> CSV (01)  =====================\n")

library(readxl)

DATA_DIR    <- file.path("..", "data")
KUMU_FILE   <- file.path(DATA_DIR, "Kumu_Galveston_AggregatedModdel_All.xlsx")
OUTPUT_FILE <- file.path(DATA_DIR, "Aggregated_Full_61.csv")

cat(sprintf("Input:  %s\n", KUMU_FILE))
cat(sprintf("Output: %s\n", OUTPUT_FILE))

elements    <- as.data.frame(read_excel(KUMU_FILE, sheet = "Elements"))
connections <- as.data.frame(read_excel(KUMU_FILE, sheet = "Connections"))

cat(sprintf("Elements:    %d concepts\n", nrow(elements)))
cat(sprintf("Connections: %d edges\n", nrow(connections)))

concepts <- as.character(elements[["Label"]])
connections$weight <- suppressWarnings(as.numeric(connections[["Original_MM_Weight"]]))

# Expand multi-tag rows into per-sector weights, then average per (From, To).
expanded_rows <- list()
for (i in seq_len(nrow(connections))) {
  row  <- connections[i, ]
  tags <- strsplit(as.character(row[["Tags"]]), "\\|")[[1]]
  for (tag in tags) {
    expanded_rows[[length(expanded_rows) + 1]] <- data.frame(
      From = as.character(row[["From"]]), To = as.character(row[["To"]]),
      weight = row[["weight"]], Tag = trimws(tag), stringsAsFactors = FALSE)
  }
}
expanded <- do.call(rbind, expanded_rows)
cat(sprintf("Original connection rows: %d ; expanded rows: %d\n", nrow(connections), nrow(expanded)))

agg_edges <- aggregate(weight ~ From + To, data = expanded,
                       FUN = function(x) mean(x, na.rm = TRUE), na.action = na.pass)
agg_edges <- agg_edges[!is.na(agg_edges$weight), ]
cat(sprintf("Unique edges after averaging: %d\n", nrow(agg_edges)))

concept_set  <- concepts
missing_from <- setdiff(unique(agg_edges$From), concept_set)
missing_to   <- setdiff(unique(agg_edges$To),   concept_set)
if (length(missing_from) == 0 && length(missing_to) == 0) cat("All edge concepts found in Elements list.\n")

matrix_df <- matrix(0.0, nrow = length(concepts), ncol = length(concepts),
                    dimnames = list(concepts, concepts))
for (i in seq_len(nrow(agg_edges))) {
  f <- agg_edges$From[i]; t <- agg_edges$To[i]; w <- agg_edges$weight[i]
  if (f %in% concept_set && t %in% concept_set) matrix_df[f, t] <- w
}
nonzero <- sum(matrix_df != 0)
cat(sprintf("Matrix: %d x %d, non-zero entries: %d\n", nrow(matrix_df), ncol(matrix_df), nonzero))

out <- matrix_df
out_chr <- matrix(as.character(out), nrow = nrow(out), dimnames = dimnames(out))
out_chr[out == 0.0] <- ""
out_df <- data.frame(Concept = rownames(out_chr), out_chr, check.names = FALSE, stringsAsFactors = FALSE)
write.csv(out_df, OUTPUT_FILE, row.names = FALSE, na = "")
cat(sprintf("Saved: %s\n", OUTPUT_FILE))


# #############################################################################
#   SECTION 2 — SHARK DEPREDATION SCENARIO
# #############################################################################

# -----------------------------------------------------------------------------
#   2a. Depredation scenario — data & baseline/scenario/change tables       [03]
#       Activates depredation in each sector model; writes per-sector Excel and
#       the original red/green comparison figures.
# -----------------------------------------------------------------------------
cat("\n\n=====================  SECTION 2a  Depredation data (03)  =====================\n")
local({
  library(openxlsx); library(ggplot2)
  DATA_DIR    <- file.path("..", "data")
  RESULTS_DIR <- file.path("..", "results")
  FIG_DIR     <- file.path(RESULTS_DIR, "figures")
  dir.create(FIG_DIR, showWarnings = FALSE, recursive = TRUE)
  
  SECTOR_MODELS <- list(
    Recreational = file.path(DATA_DIR, "Recreational_FinalModel_22April.csv"),
    Charter      = file.path(DATA_DIR, "Charter_FinalModel_22April.csv"),
    Commercial   = file.path(DATA_DIR, "Commercial_FinalModel_22April.csv"))
  DEPREDATION_NAME <- list(Recreational = "Shark Depredation",
                           Charter = "SHARK DEPREDATION", Commercial = "Shark Depredation")
  OUTPUT_FILE <- file.path(RESULTS_DIR, "step1_depredation_scenario.xlsx")
  
  loaded <- list()
  for (name in names(SECTOR_MODELS)) {
    adj <- load_adjacency_matrix(SECTOR_MODELS[[name]]); loaded[[name]] <- adj
    stopifnot(DEPREDATION_NAME[[name]] %in% adj$concepts)
    cat(sprintf("%s: %d concepts\n", name, length(adj$concepts)))
  }
  
  results <- list()
  for (name in names(SECTOR_MODELS)) {
    result <- run_scenario(loaded[[name]]$W, loaded[[name]]$concepts,
                           c(DEPREDATION_NAME[[name]]), clamp_value = 1.0)
    results[[name]] <- result
    cat(sprintf("%s: %s in %d iterations\n", name,
                if (result$converged) "Converged" else "NOT converged", result$iterations))
  }
  
  # Shared concepts across sectors (case-insensitive)
  sector_concepts <- lapply(names(SECTOR_MODELS), function(n) setNames(loaded[[n]]$concepts, tolower(loaded[[n]]$concepts)))
  names(sector_concepts) <- names(SECTOR_MODELS)
  all_lower <- sort(unique(unlist(lapply(sector_concepts, names))))
  shared <- list()
  for (cl in all_lower) {
    present_in <- names(sector_concepts)[vapply(sector_concepts, function(sc) cl %in% names(sc), logical(1))]
    if (length(present_in) >= 2) shared[[length(shared) + 1]] <- list(concept_lower = cl, sectors = present_in)
  }
  sector_order <- c("Recreational", "Charter", "Commercial")
  comparison_rows <- list()
  for (s in shared) {
    row <- list(Concept = s$concept_lower)
    for (sector_name in sector_order) {
      col <- paste0(sector_name, "_Change")
      row[[col]] <- if (sector_name %in% s$sectors)
        round(results[[sector_name]]$change[[ sector_concepts[[sector_name]][[s$concept_lower]] ]], 4) else NA_real_
    }
    comparison_rows[[length(comparison_rows) + 1]] <- as.data.frame(row, stringsAsFactors = FALSE)
  }
  comparison_df <- do.call(rbind, comparison_rows)
  rownames(comparison_df) <- comparison_df$Concept; comparison_df$Concept <- NULL
  
  # Bar chart: top 10 changes per sector (original red/green)
  bar_data <- do.call(rbind, lapply(sector_order, function(name) {
    df <- results_to_dataframe(results[[name]], loaded[[name]]$concepts)
    df <- head(df[df$Concept != DEPREDATION_NAME[[name]], ], 10)
    df$Sector <- name; df$order <- seq_len(nrow(df)); df
  }))
  bar_data$Sector <- factor(bar_data$Sector, levels = sector_order)
  bar_data$dir <- ifelse(bar_data$Change < 0, "neg", "pos")
  bar_data$ylab <- factor(paste(bar_data$Sector, bar_data$order),
                          levels = rev(paste(bar_data$Sector, bar_data$order)))
  p1 <- ggplot(bar_data, aes(x = Change, y = ylab, fill = dir)) +
    geom_col() + geom_vline(xintercept = 0, linewidth = 0.3) +
    facet_wrap(~ Sector, scales = "free", ncol = 3) +
    scale_fill_manual(values = c(neg = "#d32f2f", pos = "#388e3c"), guide = "none") +
    scale_y_discrete(labels = function(x) bar_data$Concept[match(x, bar_data$ylab)]) +
    labs(x = "Change from Baseline", y = NULL,
         title = "Impact of Shark Depredation Activation - Top 10 Affected Concepts") +
    theme_minimal(base_size = 9) + theme(plot.title = element_text(face = "bold", hjust = 0.5))
  ggsave(file.path(FIG_DIR, "step1_depredation_comparison.png"), p1, width = 18, height = 8, dpi = 300)
  
  # Excel
  wb <- createWorkbook()
  for (name in names(SECTOR_MODELS)) {
    addWorksheet(wb, name)
    writeData(wb, name, results_to_dataframe(results[[name]], loaded[[name]]$concepts))
  }
  addWorksheet(wb, "Cross_Sector_Comparison")
  writeData(wb, "Cross_Sector_Comparison", cbind(Concept = rownames(comparison_df), comparison_df))
  saveWorkbook(wb, OUTPUT_FILE, overwrite = TRUE)
  cat(sprintf("Saved: %s\n", OUTPUT_FILE))
})


# -----------------------------------------------------------------------------
#   2b. Depredation scenario — manuscript figures                          [09]
#       Contrast (main) vs baseline (SM) metric. Four styles: faceted lollipop,
#       dumbbell, crosswalk heatmap, and n=1 ego-networks. Colour-blind safe.
# -----------------------------------------------------------------------------
cat("\n\n=====================  SECTION 2b  Depredation figures (09)  =====================\n")
local({
  library(ggplot2); library(scales)
  DATA_DIR <- file.path("..", "data")
  OUT_DIR  <- file.path("..", "results", "figures_step1")
  dir.create(OUT_DIR, showWarnings = FALSE, recursive = TRUE)
  
  INC <- "#E08214"   # orange -> increase
  DEC <- "#2166AC"   # blue   -> decrease
  theme_pub <- function(b = 11) theme_minimal(base_size = b) +
    theme(plot.title = element_text(face = "bold", hjust = 0.5),
          plot.subtitle = element_text(colour = "grey30", hjust = 0.5),
          panel.grid.minor = element_blank(), panel.grid.major.y = element_blank(),
          strip.text = element_text(face = "bold"), legend.position = "bottom")
  sl <- function(x, n = 42) ifelse(nchar(x) > n, paste0(substr(x, 1, n - 1), "…"), x)
  DISPLAY_NAME <- c(
    "Participation in Tournaments - Shark mutilate catch Ex: BigRock in NC" = "Participation in Tournaments",
    "Fisher Satisfaction with Specific Species Fishing"                     = "Species Specific Fisher Satisfaction",
    "Changing Species Targeted by Recreational Fishing"                     = "Change in Target Recreational Species",
    "Lack of economic incentives to overcome the impacts of depredation"    = "Lack of Economic Incentives",
    "Shifting Distributions of Shark Populations"                           = "Shifting Shark Distributions")
  relabel <- function(x) { y <- unname(DISPLAY_NAME[x]); ifelse(is.na(y), x, y) }
  save_fig <- function(p, name, w, h, dpi = 350) {
    ggsave(file.path(OUT_DIR, paste0(name, ".png")), p, width = w, height = h, dpi = dpi, limitsize = FALSE)
    ok <- tryCatch({ ggsave(file.path(OUT_DIR, paste0(name, ".pdf")), p, width = w, height = h, device = grDevices::cairo_pdf, limitsize = FALSE); TRUE }, error = function(e) FALSE)
    if (!ok) tryCatch(ggsave(file.path(OUT_DIR, paste0(name, ".pdf")), p, width = w, height = h, limitsize = FALSE), error = function(e) NULL)
    cat(sprintf("Saved: %s.{png,pdf}\n", name))
  }
  
  MODEL_PATHS <- list(
    Recreational  = file.path(DATA_DIR, "Recreational_FinalModel_22April.csv"),
    Charter       = file.path(DATA_DIR, "Charter_FinalModel_22April.csv"),
    Commercial    = file.path(DATA_DIR, "Commercial_FinalModel_22April.csv"),
    Aggregated_30 = file.path(DATA_DIR, "MentalModeler_Galveston_Reduced30.csv"))
  MLAB  <- c(Recreational = "Recreational", Charter = "Charter", Commercial = "Commercial", Aggregated_30 = "Reduced Aggregated Model")
  MORDER <- unname(MLAB)
  DEP <- list(Recreational = "Shark Depredation", Charter = "SHARK DEPREDATION",
              Commercial = "Shark Depredation", Aggregated_30 = "Shark Depredation")
  loaded <- lapply(MODEL_PATHS, load_adjacency_matrix)
  fix_names <- function(x) { x[x == "Research & Funding"] <- "Research and Funding"; x[x == "SHARK DEPREDATION"] <- "Shark Depredation"; x }
  for (m in names(loaded)) { loaded[[m]]$concepts <- fix_names(loaded[[m]]$concepts)
  dimnames(loaded[[m]]$W) <- list(loaded[[m]]$concepts, loaded[[m]]$concepts) }
  DEP <- lapply(DEP, fix_names)
  
  state_at <- function(model, val) {
    s <- run_fcm(loaded[[model]]$W, loaded[[model]]$concepts, clamped = setNames(list(val), DEP[[model]]))$steady_state
    vapply(loaded[[model]]$concepts, function(c) s[[c]], numeric(1))
  }
  effect_table <- function(model, metric = c("contrast", "baseline")) {
    metric <- match.arg(metric); cs <- loaded[[model]]$concepts
    high <- state_at(model, 1.0)
    if (metric == "contrast") low <- state_at(model, 0.0)
    else { b <- run_fcm(loaded[[model]]$W, cs, clamped = NULL)$steady_state; low <- vapply(cs, function(c) b[[c]], numeric(1)) }
    d <- data.frame(Concept = cs, Low = low, High = high, Effect = high - low, stringsAsFactors = FALSE)
    d[d$Concept != DEP[[model]], ]
  }
  
  CROSSWALK <- list(
    "Fishermen's Income"                = c(Recreational = "Fishermen's Income", Charter = NA, Commercial = "Fishermen's Income", Aggregated_30 = "Fishermen's Income"),
    "Fishing Profitability"             = c(Recreational = NA, Charter = NA, Commercial = "Fishing Profitability", Aggregated_30 = "Fishing Profitability"),
    "Catch / Fishery Output"            = c(Recreational = NA, Charter = NA, Commercial = "Catch In Commercial Fishing", Aggregated_30 = "Catch In Commercial Fishing"),
    "Fishery Efficiency / Operations"   = c(Recreational = NA, Charter = NA, Commercial = "Efficiency of Commercial Fishery", Aggregated_30 = "Shrimping Operations"),
    "Crew / Public Safety"              = c(Recreational = "Crew Safety", Charter = NA, Commercial = "Crew Safety", Aggregated_30 = "Safety & Social Concerns"),
    "Gear & Vessel Damage"              = c(Recreational = "Damage to Gear", Charter = "Damage to Gear", Commercial = "Damage to Gear", Aggregated_30 = "Gear & Vessel Damage"),
    "Gear Costs"                        = c(Recreational = "Gear Costs", Charter = NA, Commercial = "Gear Costs", Aggregated_30 = "Gear Costs"),
    "Economic Loss / Costs"             = c(Recreational = NA, Charter = "Economic Loss", Commercial = "Cost of Fines from Gear Damage (shrimpers)", Aggregated_30 = "Economic Losses from Depredation"),
    "Fisher & Client Satisfaction"      = c(Recreational = "Client Satisfaction", Charter = NA, Commercial = NA, Aggregated_30 = "Fisher & Client Satisfaction"),
    "Fishing Behavioural Adaptation"    = c(Recreational = "Changing Species Targeted by Recreational Fishing", Charter = "Change Fishing Locations", Commercial = NA, Aggregated_30 = "Fishing Behavioral Adaptations"),
    "HMS Management Response"           = c(Recreational = "HMS Slow Management Response", Charter = NA, Commercial = NA, Aggregated_30 = "HMS Management"),
    "Insurance / Charter Risks"         = c(Recreational = "Insurance Risks for Charters", Charter = NA, Commercial = NA, Aggregated_30 = "Insurance & Charter Risks"),
    "Shark Populations / Distribution"  = c(Recreational = NA, Charter = "Shifting Distributions of Shark Populations", Commercial = NA, Aggregated_30 = "Shark Populations"),
    "Recreational Participation"        = c(Recreational = "Participation in Tournaments - Shark mutilate catch Ex: BigRock in NC", Charter = NA, Commercial = NA, Aggregated_30 = "Recreational Participation"))
  
  GROUP_COLOURS <- c(
    "Central Concept" = "#B8860B", "Ecological & Biological Factors" = "#9AD354",
    "Human Dimensions" = "#F8895E", "Fisheries Operations & Practices" = "#E382BA",
    "Fisheries Research & Management" = "#8695C2", "Policy & Economics" = "#5EB99B", "Other" = "grey60")
  CATEGORY <- c(
    "Ability to Identify Shark Species"="Human Dimensions","BP oil spill reduced prey populations (shark's food)"="Ecological & Biological Factors",
    "Bycatch Reduction Devices (Shrimp fishery)"="Fisheries Operations & Practices","Catch In Commercial Fishing"="Fisheries Operations & Practices",
    "Catch and Release Trips"="Fisheries Operations & Practices","Change Fishing Locations"="Fisheries Operations & Practices",
    "Changing Species Targeted by Recreational Fishing"="Fisheries Operations & Practices","Client Satisfaction"="Human Dimensions",
    "Climate Change"="Ecological & Biological Factors","Commercial Shark Fisheries"="Fisheries Research & Management",
    "Cost of Fines from Gear Damage (shrimpers)"="Policy & Economics","Crew Safety"="Human Dimensions",
    "Damage to Gear"="Fisheries Operations & Practices","Damage to Vessels by the Sharks Frenzy"="Fisheries Operations & Practices",
    "Demand for Sharks"="Human Dimensions","Dolphin Populations"="Ecological & Biological Factors","Economic Loss"="Policy & Economics",
    "Economic Losses from Depredation"="Policy & Economics","Efficiency of Commercial Fishery"="Fisheries Research & Management",
    "Effort in Cleaning and Processing Sharks"="Fisheries Operations & Practices","Environmental & Climate Drivers"="Ecological & Biological Factors",
    "Fisher & Client Satisfaction"="Human Dimensions","Fisher Satisfaction with Specific Species Fishing"="Human Dimensions",
    "Fisheries Management Effectiveness"="Fisheries Research & Management","Fisheries Management Effectiveness (Reef Fish)"="Fisheries Research & Management",
    "Fishermen's Income"="Human Dimensions","Fishing Behavioral Adaptations"="Fisheries Operations & Practices","Fishing Profitability"="Policy & Economics",
    "Fishing Time"="Fisheries Operations & Practices","Gear & Vessel Damage"="Policy & Economics","Gear Costs"="Policy & Economics",
    "Government Enforcement and Federal Regulations"="Policy & Economics","HMS Management"="Fisheries Research & Management",
    "HMS Management actions"="Fisheries Research & Management","HMS Slow Management Response"="Fisheries Research & Management",
    "Habitat & Ecosystem Quality"="Ecological & Biological Factors","Habitat Loss (Removal of rigs and platforms)"="Ecological & Biological Factors",
    "Health Conscious"="Human Dimensions","Hooked Fish Behavior"="Ecological & Biological Factors","Increase Shark Quotas"="Policy & Economics",
    "Increased Fishing Interest in Social Media"="Human Dimensions","Increased Number of Artificial Reefs"="Ecological & Biological Factors",
    "Insurance & Charter Risks"="Fisheries Operations & Practices","Insurance Risks for Charters"="Fisheries Operations & Practices",
    "Interest by Charter Clients"="Human Dimensions","Knowledge & Awareness"="Human Dimensions","Knowledge of Depredation"="Human Dimensions",
    "Labor Time to Repair Gear"="Policy & Economics","Lack of economic incentives to overcome the impacts of depredation"="Policy & Economics",
    "Legislation"="Policy & Economics","More Frequent Fishing Spots (Hotspots)"="Fisheries Operations & Practices","More Recreational Fishermen"="Human Dimensions",
    "Number Of Shrimp Boats"="Fisheries Operations & Practices","Participation in Tournaments - Shark mutilate catch Ex: BigRock in NC"="Human Dimensions",
    "People's Desire to Save Sharks"="Human Dimensions","Public Fear - Dangling your feet"="Human Dimensions","Public Perception & Media"="Human Dimensions",
    "Public Positive Perception"="Human Dimensions","Recreational Fishing Hotspots"="Fisheries Operations & Practices","Recreational Participation"="Human Dimensions",
    "Recreational fishing boats following commercial fishing vessels. This attracts sharks to shrimper boats"="Fisheries Operations & Practices",
    "Regulatory Framework"="Policy & Economics","Research and Funding"="Fisheries Research & Management","Safety & Social Concerns"="Human Dimensions",
    "Shark Attraction To Fishing Activity"="Ecological & Biological Factors","Shark Attraction to Vessels"="Ecological & Biological Factors",
    "Shark Behavioral Responses"="Ecological & Biological Factors","Shark Conservation"="Policy & Economics","Shark Conservation & Quota Policy"="Policy & Economics",
    "Shark Depredation"="Central Concept","Shark Fin Market"="Policy & Economics","Shark Fin Market and Demand"="Policy & Economics",
    "Shark Finning Law - Illegal to separate fins from body"="Policy & Economics","Shark Fishery"="Fisheries Research & Management",
    "Shark Learning Behavior"="Ecological & Biological Factors","Shark Populations"="Ecological & Biological Factors",
    "Shark vessel–food association (attraction and following fishing boats due to learned food rewards)"="Ecological & Biological Factors",
    "Shifting Distributions of Shark Populations"="Ecological & Biological Factors","Shrimping Operations"="Fisheries Operations & Practices",
    "Switch from Bait to Lures/Jigs"="Fisheries Operations & Practices","Water Temperature"="Ecological & Biological Factors")
  CANON_CATEGORY <- c(
    "Fishermen's Income"="Human Dimensions","Fishing Profitability"="Policy & Economics","Catch / Fishery Output"="Fisheries Operations & Practices",
    "Fishery Efficiency / Operations"="Fisheries Operations & Practices","Crew / Public Safety"="Human Dimensions",
    "Gear & Vessel Damage"="Policy & Economics","Gear Costs"="Policy & Economics","Economic Loss / Costs"="Policy & Economics",
    "Fisher & Client Satisfaction"="Human Dimensions","Fishing Behavioural Adaptation"="Fisheries Operations & Practices",
    "HMS Management Response"="Fisheries Research & Management","Insurance / Charter Risks"="Fisheries Operations & Practices",
    "Shark Populations / Distribution"="Ecological & Biological Factors","Recreational Participation"="Human Dimensions")
  cat_of <- function(concept) { v <- unname(CATEGORY[concept]); v[is.na(v)] <- "Other"; v }
  
  fig_lollipop <- function(metric, fname, subtitle, top_n = 12, thresh = 0.02) {
    d <- do.call(rbind, lapply(names(MODEL_PATHS), function(m) {
      e <- effect_table(m, metric); e <- e[order(-abs(e$Effect)), ]
      e <- e[abs(e$Effect) >= thresh, ]; if (nrow(e) > top_n) e <- e[seq_len(top_n), ]
      if (nrow(e) == 0) return(NULL); e$Model <- MLAB[[m]]; e$row <- seq_len(nrow(e)); e }))
    d$Model <- factor(d$Model, levels = MORDER)
    d$Direction <- ifelse(d$Effect >= 0, "increase", "decrease")
    d$key <- factor(paste(d$Model, d$row), levels = rev(paste(d$Model, d$row)))
    d$lab <- sl(relabel(d$Concept))
    p <- ggplot(d, aes(x = Effect, y = key, colour = Direction)) +
      geom_segment(aes(x = 0, xend = Effect, yend = key), linewidth = 0.6) +
      geom_point(size = 2.4) + geom_vline(xintercept = 0, colour = "grey45", linewidth = 0.3) +
      facet_wrap(~ Model, scales = "free", ncol = 2) +
      scale_colour_manual(values = c(increase = INC, decrease = DEC), name = "Response to more depredation") +
      scale_y_discrete(labels = function(x) d$lab[match(x, d$key)]) +
      labs(title = "How increasing shark depredation propagates through each model",
           subtitle = subtitle, x = "Effect on concept (steady-state change)", y = NULL) + theme_pub()
    save_fig(p, fname, 12, 8)
  }
  fig_dumbbell <- function(metric, fname, subtitle, lolab, hilab, top_n = 12, thresh = 0.02) {
    d <- do.call(rbind, lapply(names(MODEL_PATHS), function(m) {
      e <- effect_table(m, metric); e <- e[order(-abs(e$Effect)), ]
      e <- e[abs(e$Effect) >= thresh, ]; if (nrow(e) > top_n) e <- e[seq_len(top_n), ]
      if (nrow(e) == 0) return(NULL); e$Model <- MLAB[[m]]; e$row <- seq_len(nrow(e)); e }))
    d$Model <- factor(d$Model, levels = MORDER)
    d$Category <- factor(cat_of(d$Concept), levels = names(GROUP_COLOURS))
    d$key <- factor(paste(d$Model, d$row), levels = rev(paste(d$Model, d$row)))
    d$lab <- sl(relabel(d$Concept))
    long <- rbind(
      data.frame(key = d$key, Model = d$Model, val = d$Low,  State = lolab, Category = d$Category),
      data.frame(key = d$key, Model = d$Model, val = d$High, State = hilab, Category = d$Category))
    long$State <- factor(long$State, levels = c(lolab, hilab))
    p <- ggplot() +
      geom_segment(data = d, aes(y = key, yend = key, x = Low, xend = High, colour = Category), linewidth = 0.9) +
      geom_point(data = long, aes(y = key, x = val, shape = State, colour = Category), size = 2.5, stroke = 0.9) +
      facet_wrap(~ Model, scales = "free_y", ncol = 2) +
      scale_colour_manual(values = GROUP_COLOURS, name = "Category", drop = TRUE) +
      scale_shape_manual(values = c(1, 19), name = NULL) +
      scale_y_discrete(labels = function(x) d$lab[match(x, d$key)]) +
      guides(shape  = guide_legend(order = 1, nrow = 1),
             colour = guide_legend(order = 2, nrow = 2, byrow = TRUE, override.aes = list(shape = 15, size = 4))) +
      labs(title = "", x = "", y = NULL) + theme_pub() +
      theme(legend.box = "vertical", legend.justification = "center", legend.box.just = "center",
            legend.text = element_text(size = 8), legend.margin = margin(2, 2, 2, 2),
            plot.margin = margin(t = 6, r = 8, b = 12, l = 8))
    save_fig(p, fname, 12, 9.2, dpi = 1200)
  }
  fig_heatmap <- function(metric, fname, subtitle, limit) {
    eff <- lapply(names(MODEL_PATHS), function(m) { e <- effect_table(m, metric); setNames(e$Effect, e$Concept) })
    names(eff) <- names(MODEL_PATHS)
    rows <- list()
    for (canon in names(CROSSWALK)) for (m in names(MODEL_PATHS)) {
      tgt <- CROSSWALK[[canon]][[m]]
      val <- if (is.na(tgt)) NA_real_ else unname(eff[[m]][tgt])
      rows[[length(rows) + 1]] <- data.frame(Canon = canon, Model = MLAB[[m]], Value = val, stringsAsFactors = FALSE)
    }
    h <- do.call(rbind, rows)
    keep <- names(which(tapply(abs(h$Value), h$Canon, function(x) any(x >= 0.05, na.rm = TRUE))))
    h <- h[h$Canon %in% keep, ]
    ord <- names(sort(tapply(abs(h$Value), h$Canon, function(x) sum(x, na.rm = TRUE)), decreasing = TRUE))
    h$Canon <- factor(h$Canon, levels = rev(ord)); h$Model <- factor(h$Model, levels = MORDER)
    h$Category <- factor(unname(CANON_CATEGORY[as.character(h$Canon)]), levels = names(GROUP_COLOURS))
    ycols <- unname(GROUP_COLOURS[CANON_CATEGORY[levels(h$Canon)]]); ycols[is.na(ycols)] <- "grey40"
    p <- ggplot(h, aes(x = Model, y = Canon, fill = Value)) +
      geom_tile(colour = "white", linewidth = 0.6) +
      geom_text(aes(label = ifelse(is.na(Value), "", sprintf("%.2f", Value))), size = 2.6, colour = "grey15") +
      geom_point(aes(colour = Category), alpha = 0, na.rm = TRUE) +
      scale_fill_gradient2(low = DEC, mid = "#F7F7F7", high = INC, midpoint = 0,
                           limits = c(-limit, limit), oob = scales::squish, na.value = "grey88",
                           name = "Effect of \nincreased depredation") +
      scale_colour_manual(values = GROUP_COLOURS, name = "Category", drop = TRUE) +
      guides(fill   = guide_colourbar(order = 1, position = "right"),
             colour = guide_legend(order = 2, position = "bottom", nrow = 2, override.aes = list(alpha = 1, shape = 15, size = 4))) +
      labs(title = "Downstream effect of increasing shark depredation, by model", x = NULL, y = NULL) +
      theme_pub() +
      theme(legend.key.width = unit(1.2, "cm"), axis.text.y = element_text(colour = ycols, face = "bold"),
            legend.title = element_text(hjust = 0.5), legend.margin = margin(4, 4, 4, 4),
            plot.margin = margin(t = 6, r = 10, b = 16, l = 10))
    save_fig(p, fname, 10.5, 9)
  }
  fig_ego <- function(fname) {
    png(file.path(OUT_DIR, paste0(fname, ".png")), width = 2600, height = 2600, res = 300)
    op <- par(mfrow = c(2, 2), mar = c(1, 1, 2.5, 1)); on.exit({ par(op); dev.off() })
    for (m in names(MODEL_PATHS)) {
      W <- loaded[[m]]$W; cs <- loaded[[m]]$concepts; i <- match(DEP[[m]], cs)
      js <- which(W[i, ] != 0); w <- W[i, js]; nm <- cs[js]
      o <- order(w); js <- js[o]; w <- w[o]; nm <- nm[o]; n <- length(js)
      plot(NA, xlim = c(-1.6, 1.6), ylim = c(-1.6, 1.6), asp = 1, axes = FALSE, xlab = "", ylab = "",
           main = MLAB[[m]], font.main = 2)
      if (n > 0) {
        ang <- seq(pi/2, pi/2 + 2*pi, length.out = n + 1)[seq_len(n)]; x <- cos(ang); y <- sin(ang)
        for (k in seq_len(n)) {
          col <- if (w[k] > 0) INC else DEC
          segments(0, 0, 0.86 * x[k], 0.86 * y[k], col = col, lwd = 1 + 3 * abs(w[k]))
          arrows(0.6 * x[k], 0.6 * y[k], 0.86 * x[k], 0.86 * y[k], col = col, length = 0.06, lwd = 1 + 3 * abs(w[k]))
        }
        points(x, y, pch = 21, bg = "white", col = "grey40", cex = 1.1)
        text(1.02 * x, 1.02 * y, labels = sl(nm, 30), cex = 0.52, adj = ifelse(x >= 0, 0, 1), xpd = NA)
      }
      points(0, 0, pch = 21, bg = "grey25", col = "black", cex = 3.4)
      text(0, 0, "Shark\nDepredation", col = "white", cex = 0.5, font = 2)
    }
    cat(sprintf("Saved: %s.png\n", fname))
  }
  
  cat("--- MAIN (contrast metric) ---\n")
  fig_lollipop("contrast", "fig1_lollipop_contrast", "Effect metric: depredation forced high (1.0) minus forced low (0.0)")
  fig_dumbbell("contrast", "fig2_dumbbell_contrast", "State of each concept when depredation is low (0.0) vs high (1.0)",
               lolab = "Depredation low (0.0)", hilab = "Depredation high (1.0)")
  fig_heatmap ("contrast", "fig3_heatmap_contrast", "", 0.40)
  fig_ego     ("fig4_ego_networks")
  cat("--- SUPPLEMENTARY (baseline metric) ---\n")
  fig_lollipop("baseline", "figS_lollipop_baseline", "Effect metric: depredation clamped to 1.0 minus free-running baseline")
  fig_dumbbell("baseline", "figS_dumbbell_baseline", "State of each concept at baseline vs depredation clamped to 1.0",
               lolab = "Baseline", hilab = "Depredation = 1.0")
  fig_heatmap ("baseline", "figS_heatmap_baseline", "", 0.10)
})


# -----------------------------------------------------------------------------
#   2c. Depredation scenario — activation heatmap (change-from-baseline)    [10]
#       Supplementary heatmap: depredation clamped to 1.0 minus baseline.
# -----------------------------------------------------------------------------
cat("\n\n=====================  SECTION 2c  Activation heatmap (10)  =====================\n")
local({
  library(ggplot2); library(scales)
  DATA_DIR <- file.path("..", "data")
  OUT_DIR  <- file.path("..", "results", "figures_step1")
  dir.create(OUT_DIR, showWarnings = FALSE, recursive = TRUE)
  INC <- "#E08214"; DEC <- "#2166AC"
  GROUP_COLOURS <- c(
    "Central Concept"="#B8860B","Ecological & Biological Factors"="#9AD354","Human Dimensions"="#F8895E",
    "Fisheries Operations & Practices"="#E382BA","Fisheries Research & Management"="#8695C2",
    "Policy & Economics"="#5EB99B","Other"="grey60")
  sl <- function(x, n = 42) ifelse(nchar(x) > n, paste0(substr(x, 1, n - 1), "…"), x)
  MODEL_PATHS <- list(
    Recreational  = file.path(DATA_DIR, "Recreational_FinalModel_22April.csv"),
    Charter       = file.path(DATA_DIR, "Charter_FinalModel_22April.csv"),
    Commercial    = file.path(DATA_DIR, "Commercial_FinalModel_22April.csv"),
    Aggregated_30 = file.path(DATA_DIR, "MentalModeler_Galveston_Reduced30.csv"))
  MLAB  <- c(Recreational = "Recreational", Charter = "Charter", Commercial = "Commercial", Aggregated_30 = "Reduced Aggregated Model")
  MORDER <- unname(MLAB)
  DEP <- list(Recreational = "Shark Depredation", Charter = "SHARK DEPREDATION",
              Commercial = "Shark Depredation", Aggregated_30 = "Shark Depredation")
  loaded <- lapply(MODEL_PATHS, load_adjacency_matrix)
  CROSSWALK <- list(
    "Fishermen's Income"                = c(Recreational = "Fishermen's Income", Charter = NA, Commercial = "Fishermen's Income", Aggregated_30 = "Fishermen's Income"),
    "Fishing Profitability"             = c(Recreational = NA, Charter = NA, Commercial = "Fishing Profitability", Aggregated_30 = "Fishing Profitability"),
    "Catch / Fishery Output"            = c(Recreational = NA, Charter = NA, Commercial = "Catch In Commercial Fishing", Aggregated_30 = "Catch In Commercial Fishing"),
    "Fishery Efficiency / Operations"   = c(Recreational = NA, Charter = NA, Commercial = "Efficiency of Commercial Fishery", Aggregated_30 = "Shrimping Operations"),
    "Crew / Public Safety"              = c(Recreational = "Crew Safety", Charter = NA, Commercial = "Crew Safety", Aggregated_30 = "Safety & Social Concerns"),
    "Gear & Vessel Damage"              = c(Recreational = "Damage to Gear", Charter = "Damage to Gear", Commercial = "Damage to Gear", Aggregated_30 = "Gear & Vessel Damage"),
    "Gear Costs"                        = c(Recreational = "Gear Costs", Charter = NA, Commercial = "Gear Costs", Aggregated_30 = "Gear Costs"),
    "Economic Loss / Costs"             = c(Recreational = NA, Charter = "Economic Loss", Commercial = "Cost of Fines from Gear Damage (shrimpers)", Aggregated_30 = "Economic Losses from Depredation"),
    "Fisher & Client Satisfaction"      = c(Recreational = "Client Satisfaction", Charter = NA, Commercial = NA, Aggregated_30 = "Fisher & Client Satisfaction"),
    "Fishing Behavioural Adaptation"    = c(Recreational = "Changing Species Targeted by Recreational Fishing", Charter = "Change Fishing Locations", Commercial = NA, Aggregated_30 = "Fishing Behavioral Adaptations"),
    "HMS Management Response"           = c(Recreational = "HMS Slow Management Response", Charter = NA, Commercial = NA, Aggregated_30 = "HMS Management"),
    "Insurance / Charter Risks"         = c(Recreational = "Insurance Risks for Charters", Charter = NA, Commercial = NA, Aggregated_30 = "Insurance & Charter Risks"),
    "Shark Populations / Distribution"  = c(Recreational = NA, Charter = "Shifting Distributions of Shark Populations", Commercial = NA, Aggregated_30 = "Shark Populations"),
    "Recreational Participation"        = c(Recreational = "Participation in Tournaments - Shark mutilate catch Ex: BigRock in NC", Charter = NA, Commercial = NA, Aggregated_30 = "Recreational Participation"))
  CANON_CATEGORY <- c(
    "Fishermen's Income"="Human Dimensions","Fishing Profitability"="Policy & Economics","Catch / Fishery Output"="Fisheries Operations & Practices",
    "Fishery Efficiency / Operations"="Fisheries Operations & Practices","Crew / Public Safety"="Human Dimensions",
    "Gear & Vessel Damage"="Policy & Economics","Gear Costs"="Policy & Economics","Economic Loss / Costs"="Policy & Economics",
    "Fisher & Client Satisfaction"="Human Dimensions","Fishing Behavioural Adaptation"="Fisheries Operations & Practices",
    "HMS Management Response"="Fisheries Research & Management","Insurance / Charter Risks"="Fisheries Operations & Practices",
    "Shark Populations / Distribution"="Ecological & Biological Factors","Recreational Participation"="Human Dimensions")
  effect_activation <- function(model) {
    cs <- loaded[[model]]$concepts; W <- loaded[[model]]$W
    base <- run_fcm(W, cs, clamped = NULL)$steady_state
    scen <- run_fcm(W, cs, clamped = setNames(list(1.0), DEP[[model]]))$steady_state
    setNames(vapply(cs, function(c) scen[[c]] - base[[c]], numeric(1)), cs)
  }
  eff <- lapply(names(MODEL_PATHS), effect_activation); names(eff) <- names(MODEL_PATHS)
  LIMIT <- 0.05; KEEP <- 0.01
  rows <- list()
  for (canon in names(CROSSWALK)) for (m in names(MODEL_PATHS)) {
    tgt <- CROSSWALK[[canon]][[m]]; val <- if (is.na(tgt)) NA_real_ else unname(eff[[m]][tgt])
    rows[[length(rows) + 1]] <- data.frame(Canon = canon, Model = MLAB[[m]], Value = val, stringsAsFactors = FALSE)
  }
  h <- do.call(rbind, rows)
  keep <- names(which(tapply(abs(h$Value), h$Canon, function(x) any(x >= KEEP, na.rm = TRUE))))
  h <- h[h$Canon %in% keep, ]
  ord <- names(sort(tapply(abs(h$Value), h$Canon, function(x) sum(x, na.rm = TRUE)), decreasing = TRUE))
  h$Canon <- factor(h$Canon, levels = rev(ord)); h$Model <- factor(h$Model, levels = MORDER)
  h$Category <- factor(unname(CANON_CATEGORY[as.character(h$Canon)]), levels = names(GROUP_COLOURS))
  ycols <- unname(GROUP_COLOURS[CANON_CATEGORY[levels(h$Canon)]]); ycols[is.na(ycols)] <- "grey40"
  p <- ggplot(h, aes(x = Model, y = Canon, fill = Value)) +
    geom_tile(colour = "white", linewidth = 0.6) +
    geom_text(aes(label = ifelse(is.na(Value), "", sprintf("%.3f", Value))), size = 2.6, colour = "grey15") +
    geom_point(aes(colour = Category), alpha = 0, na.rm = TRUE) +
    scale_fill_gradient2(low = DEC, mid = "#F7F7F7", high = INC, midpoint = 0,
                         limits = c(-LIMIT, LIMIT), oob = scales::squish, na.value = "grey88",
                         name = "Change if depredation\nactivated to 1.0") +
    scale_colour_manual(values = GROUP_COLOURS, name = "Concept category", drop = TRUE) +
    guides(fill = guide_colourbar(order = 1, position = "right"),
           colour = guide_legend(order = 2, position = "bottom", nrow = 2, override.aes = list(alpha = 1, shape = 15, size = 4))) +
    labs(title = "Change in each concept when shark depredation is activated (= 1.0)",
         subtitle = "Change from baseline. Grey = concept absent from model.", x = NULL, y = NULL) +
    theme_minimal(base_size = 11) +
    theme(plot.title = element_text(face = "bold", hjust = 0.5), plot.subtitle = element_text(colour = "grey30", hjust = 0.5),
          panel.grid = element_blank(), axis.text.y = element_text(colour = ycols, face = "bold"),
          legend.key.width = unit(1.2, "cm"), legend.title = element_text(hjust = 0.5),
          plot.margin = margin(t = 6, r = 10, b = 16, l = 10))
  ggsave(file.path(OUT_DIR, "heatmap_depredation_activated.png"), p, width = 10.5, height = 8, dpi = 350)
  tryCatch(ggsave(file.path(OUT_DIR, "heatmap_depredation_activated.pdf"), p, width = 10.5, height = 8, device = grDevices::cairo_pdf),
           error = function(e) ggsave(file.path(OUT_DIR, "heatmap_depredation_activated.pdf"), p, width = 10.5, height = 8))
  cat("Saved: heatmap_depredation_activated.{png,pdf}\n")
})


# #############################################################################
#   SECTION 3 — THEMATIC SCENARIOS (Ecological / Management / Fisheries-Behaviour)
# #############################################################################

# -----------------------------------------------------------------------------
#   3a. Thematic scenarios — individual + additive + comparison            [07]
#       Runs each themed driver individually (clamp 1.0) and as additive bundles
#       (all concepts of a theme at once), and compares the two. Writes Excel.
# -----------------------------------------------------------------------------
cat("\n\n=====================  SECTION 3a  Thematic scenarios (07)  =====================\n")
local({
  library(openxlsx); library(ggplot2)
  DATA_DIR    <- file.path("..", "data")
  RESULTS_DIR <- file.path("..", "results")
  FIG_DIR     <- file.path(RESULTS_DIR, "figures")
  dir.create(FIG_DIR, showWarnings = FALSE, recursive = TRUE)
  OUTPUT_FILE <- file.path(RESULTS_DIR, "step3_thematic_scenarios_revised.xlsx")
  DISPLAY_NAME <- c(
    "Participation in Tournaments - Shark mutilate catch Ex: BigRock in NC" = "Participation in Tournaments",
    "Fisher Satisfaction with Specific Species Fishing"                     = "Species Specific Fisher Satisfaction",
    "Changing Species Targeted by Recreational Fishing"                     = "Change in Target Recreational Species",
    "Lack of economic incentives to overcome the impacts of depredation"    = "Lack of Economic Incentives",
    "Shifting Distributions of Shark Populations"                           = "Shifting Shark Distributions")
  relabel <- function(x) { y <- unname(DISPLAY_NAME[x]); ifelse(is.na(y), x, y) }
  MODEL_PATHS <- list(
    Recreational  = file.path(DATA_DIR, "Recreational_FinalModel_22April.csv"),
    Charter       = file.path(DATA_DIR, "Charter_FinalModel_22April.csv"),
    Commercial    = file.path(DATA_DIR, "Commercial_FinalModel_22April.csv"),
    Aggregated_30 = file.path(DATA_DIR, "MentalModeler_Galveston_Reduced30.csv"))
  SCENARIOS <- list(
    Recreational = list(
      Ecological = c("Habitat Loss (Removal of rigs and platforms)"),
      Management = c("Shark Finning Law - Illegal to separate fins from body", "Commercial Shark Fisheries", "HMS Slow Management Response"),
      Fisheries_Behavior = c("Demand for Sharks", "Public Positive Perception", "More Frequent Fishing Spots (Hotspots)")),
    Charter = list(
      Ecological = c("Shark Populations", "Habitat Loss (Removal of rigs and platforms)", "Shark vessel–food association (attraction and following fishing boats due to learned food rewards)"),
      Management = c("HMS Management actions", "Shark Conservation", "Increase Shark Quotas"),
      Fisheries_Behavior = c("Knowledge of Depredation", "Damage to Gear", "Change Fishing Locations")),
    Commercial = list(
      Ecological = c("Recreational fishing boats following commercial fishing vessels. This attracts sharks to shrimper boats", "Shark Populations"),
      Management = c("Commercial Shark Fisheries"),
      Fisheries_Behavior = c("Catch In Commercial Fishing", "People's Desire to Save Sharks", "Damage to Gear")),
    Aggregated_30 = list(
      Ecological = c("Shark Populations", "Habitat & Ecosystem Quality", "Shark Attraction To Fishing Activity"),
      Management = c("Commercial Shark Fisheries", "Economic Losses from Depredation", "Regulatory Framework"),
      Fisheries_Behavior = c("Shark Fin Market and Demand", "Public Perception & Media", "Recreational Fishing Hotspots")))
  total <- sum(vapply(SCENARIOS, function(m) sum(vapply(m, length, integer(1))), integer(1)))
  cat(sprintf("Total individual scenarios to run: %d\n", total))
  loaded <- lapply(MODEL_PATHS, load_adjacency_matrix)
  
  # Run individual scenarios
  all_results <- list()
  for (model_name in names(SCENARIOS)) {
    W <- loaded[[model_name]]$W; concepts <- loaded[[model_name]]$concepts
    all_results[[model_name]] <- list()
    for (theme in names(SCENARIOS[[model_name]])) {
      all_results[[model_name]][[theme]] <- list()
      for (activated_concept in SCENARIOS[[model_name]][[theme]]) {
        all_results[[model_name]][[theme]][[activated_concept]] <-
          run_scenario(W, concepts, c(activated_concept), clamp_value = 1.0)
      }
    }
  }
  # Per-model-theme bar figures (original red/green)
  for (model_name in names(all_results)) {
    concepts <- loaded[[model_name]]$concepts
    for (theme in names(all_results[[model_name]])) {
      scenario_results <- all_results[[model_name]][[theme]]
      bar_data <- do.call(rbind, lapply(names(scenario_results), function(activated) {
        df <- results_to_dataframe(scenario_results[[activated]], concepts)
        df <- head(df[df$Concept != activated, ], 10)
        short_name <- if (nchar(activated) > 45) paste0(substr(activated, 1, 45), "...") else activated
        df$Activated <- short_name; df$order <- seq_len(nrow(df)); df }))
      bar_data$Activated <- factor(bar_data$Activated, levels = unique(bar_data$Activated))
      bar_data$dir <- ifelse(bar_data$Change < 0, "neg", "pos")
      bar_data$ylab <- factor(paste(bar_data$Activated, bar_data$order), levels = rev(unique(paste(bar_data$Activated, bar_data$order))))
      n_scenarios <- length(scenario_results)
      p <- ggplot(bar_data, aes(x = Change, y = ylab, fill = dir)) +
        geom_col() + geom_vline(xintercept = 0, linewidth = 0.3) +
        facet_wrap(~ Activated, scales = "free", nrow = 1) +
        scale_fill_manual(values = c(neg = "#d32f2f", pos = "#388e3c"), guide = "none") +
        scale_y_discrete(labels = function(x) relabel(bar_data$Concept[match(x, bar_data$ylab)])) +
        labs(x = "Change from Baseline", y = NULL, title = sprintf("%s - %s Scenarios (Top 10 Changes)", model_name, theme)) +
        theme_minimal(base_size = 8) + theme(plot.title = element_text(face = "bold", hjust = 0.5))
      ggsave(file.path(FIG_DIR, sprintf("step3rev_%s_%s.png", model_name, theme)), p, width = 6 * n_scenarios, height = 7, dpi = 300, limitsize = FALSE)
    }
  }
  # Summary + Excel
  model_abbr <- c(Recreational = "Rec", Charter = "Cha", Commercial = "Com", Aggregated_30 = "A30")
  theme_abbr <- c(Ecological = "Eco", Management = "Mgt", Fisheries_Behavior = "Beh")
  summary_rows <- list()
  for (model_name in names(all_results)) {
    concepts <- loaded[[model_name]]$concepts
    for (theme in names(all_results[[model_name]])) for (activated in names(all_results[[model_name]][[theme]])) {
      df <- results_to_dataframe(all_results[[model_name]][[theme]][[activated]], concepts)
      df_excl <- df[df$Concept != activated, ]
      summary_rows[[length(summary_rows) + 1]] <- data.frame(
        Model = model_name, Theme = theme, Activated_Concept = activated,
        Concepts_Increased = sum(df_excl$Change > 0.001), Concepts_Decreased = sum(df_excl$Change < -0.001),
        Max_Positive_Change = round(max(df_excl$Change), 4), Max_Negative_Change = round(min(df_excl$Change), 4),
        Mean_Abs_Change = round(mean(abs(df_excl$Change)), 4), stringsAsFactors = FALSE)
    }
  }
  summary_df <- do.call(rbind, summary_rows)
  wb <- createWorkbook(); addWorksheet(wb, "Summary"); writeData(wb, "Summary", summary_df)
  for (model_name in names(all_results)) {
    concepts <- loaded[[model_name]]$concepts
    for (theme in names(all_results[[model_name]])) for (activated in names(all_results[[model_name]][[theme]])) {
      df <- results_to_dataframe(all_results[[model_name]][[theme]][[activated]], concepts)
      concept_short <- gsub("/", "-", substr(activated, 1, 20))
      sheet_name <- substr(sprintf("%s_%s_%s", model_abbr[[model_name]], theme_abbr[[theme]], concept_short), 1, 31)
      base_name <- sheet_name; k <- 1
      while (sheet_name %in% names(wb)) { k <- k + 1; sheet_name <- substr(paste0(base_name, k), 1, 31) }
      addWorksheet(wb, sheet_name); writeData(wb, sheet_name, df)
    }
  }
  saveWorkbook(wb, OUTPUT_FILE, overwrite = TRUE); cat(sprintf("Saved: %s\n", OUTPUT_FILE))
  
  # Additive (co-activation) scenarios + comparison
  ADD_OUTPUT_FILE <- file.path(RESULTS_DIR, "step3_additive_scenarios.xlsx")
  CMP_OUTPUT_FILE <- file.path(RESULTS_DIR, "step3_individual_vs_additive.xlsx")
  DEP_NAME <- list(Recreational = "Shark Depredation", Charter = "SHARK DEPREDATION",
                   Commercial = "Shark Depredation", Aggregated_30 = "Shark Depredation")
  additive_results <- list()
  for (model_name in names(SCENARIOS)) {
    W <- loaded[[model_name]]$W; concepts <- loaded[[model_name]]$concepts
    additive_results[[model_name]] <- list()
    for (theme in names(SCENARIOS[[model_name]]))
      additive_results[[model_name]][[theme]] <- run_scenario(W, concepts, SCENARIOS[[model_name]][[theme]], clamp_value = 1.0)
  }
  add_summary_rows <- list(); wb_add <- createWorkbook(); addWorksheet(wb_add, "Summary")
  for (model_name in names(additive_results)) {
    concepts <- loaded[[model_name]]$concepts
    for (theme in names(additive_results[[model_name]])) {
      result <- additive_results[[model_name]][[theme]]; bundle <- SCENARIOS[[model_name]][[theme]]
      df <- results_to_dataframe(result, concepts); df_excl <- df[!(df$Concept %in% bundle), ]
      add_summary_rows[[length(add_summary_rows) + 1]] <- data.frame(
        Model = model_name, Theme = theme, N_Concepts_Activated = length(bundle),
        Activated_Concepts = paste(bundle, collapse = "; "),
        Concepts_Increased = sum(df_excl$Change > 0.001), Concepts_Decreased = sum(df_excl$Change < -0.001),
        Max_Positive_Change = round(max(df_excl$Change), 4), Max_Negative_Change = round(min(df_excl$Change), 4),
        Mean_Abs_Change = round(mean(abs(df_excl$Change)), 4), stringsAsFactors = FALSE)
      sh <- substr(sprintf("%s_%s_ALL", model_abbr[[model_name]], theme_abbr[[theme]]), 1, 31)
      addWorksheet(wb_add, sh); writeData(wb_add, sh, df)
    }
  }
  writeData(wb_add, "Summary", do.call(rbind, add_summary_rows))
  saveWorkbook(wb_add, ADD_OUTPUT_FILE, overwrite = TRUE); cat(sprintf("Saved: %s\n", ADD_OUTPUT_FILE))
  
  cmp_rows <- list()
  for (model_name in names(SCENARIOS)) {
    concepts <- loaded[[model_name]]$concepts; dep <- DEP_NAME[[model_name]]
    for (theme in names(SCENARIOS[[model_name]])) {
      bundle <- SCENARIOS[[model_name]][[theme]]; add <- additive_results[[model_name]][[theme]]$change
      sum_change <- setNames(rep(0, length(concepts)), concepts)
      for (c in bundle) { ind <- all_results[[model_name]][[theme]][[c]]$change; for (x in concepts) sum_change[[x]] <- sum_change[[x]] + ind[[x]] }
      mask <- !(concepts %in% bundle)
      a <- vapply(concepts[mask], function(x) add[[x]], numeric(1)); s <- vapply(concepts[mask], function(x) sum_change[[x]], numeric(1))
      corr  <- if (sd(a) > 1e-9 && sd(s) > 1e-9) cor(a, s) else NA_real_
      ratio <- if (sum(abs(s)) > 0) sum(abs(a)) / sum(abs(s)) else NA_real_
      sum_dep <- sum(vapply(bundle, function(c) all_results[[model_name]][[theme]][[c]]$change[[dep]], numeric(1)))
      cmp_rows[[length(cmp_rows) + 1]] <- data.frame(
        Model = model_name, Theme = theme, N_Activated = length(bundle),
        Additive_Depredation = round(add[[dep]], 4), Sum_Individual_Depredation = round(sum_dep, 4),
        Vector_Correlation = round(corr, 3), Additive_vs_Sum_Ratio = round(ratio, 3), stringsAsFactors = FALSE)
    }
  }
  cmp_df <- do.call(rbind, cmp_rows); cat("\nIndividual vs additive comparison:\n"); print(cmp_df, row.names = FALSE)
  wb_cmp <- createWorkbook(); addWorksheet(wb_cmp, "Comparison"); writeData(wb_cmp, "Comparison", cmp_df)
  saveWorkbook(wb_cmp, CMP_OUTPUT_FILE, overwrite = TRUE); cat(sprintf("Saved: %s\n", CMP_OUTPUT_FILE))
})


# -----------------------------------------------------------------------------
#   3b. Thematic scenarios — cross-model synthesis figures                 [08]
#       Effect of each lever on Shark Depredation + heatmap of levers x shared
#       outcome concepts + management spotlight (absolute & normalised).
# -----------------------------------------------------------------------------
cat("\n\n=====================  SECTION 3b  Cross-model figures (08)  =====================\n")
local({
  library(ggplot2); library(scales)
  DATA_DIR <- file.path("..", "data")
  FIG_DIR  <- file.path("..", "results", "figures_manuscript")
  dir.create(FIG_DIR, showWarnings = FALSE, recursive = TRUE)
  SECTOR_COLS <- c("Recreational" = "#0072B2", "Charter" = "#E69F00", "Commercial" = "#CC79A7", "Reduced (30)" = "#333333")
  DIV_LOW <- "#B35806"; DIV_MID <- "#F7F7F7"; DIV_HIGH <- "#542788"
  theme_pub <- function(base = 11) theme_minimal(base_size = base) +
    theme(plot.title = element_text(face = "bold", hjust = 0), plot.subtitle = element_text(colour = "grey30"),
          panel.grid.minor = element_blank(), panel.grid.major.y = element_blank(),
          strip.text = element_text(face = "bold"), legend.position = "bottom")
  short_lab <- function(x, n = 40) ifelse(nchar(x) > n, paste0(substr(x, 1, n - 1), "…"), x)
  save_fig <- function(p, name, w, h) {
    ggsave(file.path(FIG_DIR, paste0(name, ".png")), p, width = w, height = h, dpi = 350, limitsize = FALSE)
    ok <- tryCatch({ ggsave(file.path(FIG_DIR, paste0(name, ".pdf")), p, width = w, height = h, device = grDevices::cairo_pdf, limitsize = FALSE); TRUE }, error = function(e) FALSE)
    if (!ok) tryCatch(ggsave(file.path(FIG_DIR, paste0(name, ".pdf")), p, width = w, height = h, limitsize = FALSE), error = function(e) NULL)
    cat(sprintf("Saved: %s.{png,pdf}\n", name))
  }
  MODEL_PATHS <- list(
    Recreational  = file.path(DATA_DIR, "Recreational_FinalModel_22April.csv"),
    Charter       = file.path(DATA_DIR, "Charter_FinalModel_22April.csv"),
    Commercial    = file.path(DATA_DIR, "Commercial_FinalModel_22April.csv"),
    Aggregated_30 = file.path(DATA_DIR, "MentalModeler_Galveston_Reduced30.csv"))
  loaded <- lapply(MODEL_PATHS, load_adjacency_matrix)
  MODEL_LABEL <- c(Recreational = "Recreational", Charter = "Charter", Commercial = "Commercial", Aggregated_30 = "Reduced (30)")
  MODEL_ORDER <- c("Recreational", "Charter", "Commercial", "Reduced (30)")
  SCENARIOS <- list(
    Recreational = list(
      Ecological = c("Habitat Loss (Removal of rigs and platforms)"),
      Management = c("Shark Finning Law - Illegal to separate fins from body", "Commercial Shark Fisheries", "HMS Slow Management Response"),
      Fisheries_Behavior = c("Demand for Sharks", "Public Positive Perception", "More Frequent Fishing Spots (Hotspots)")),
    Charter = list(
      Ecological = c("Shark Populations", "Habitat Loss (Removal of rigs and platforms)", "Shark vessel–food association (attraction and following fishing boats due to learned food rewards)"),
      Management = c("HMS Management actions", "Shark Conservation", "Increase Shark Quotas"),
      Fisheries_Behavior = c("Knowledge of Depredation", "Damage to Gear", "Change Fishing Locations")),
    Commercial = list(
      Ecological = c("Recreational fishing boats following commercial fishing vessels. This attracts sharks to shrimper boats", "Shark Populations"),
      Management = c("Commercial Shark Fisheries"),
      Fisheries_Behavior = c("Catch In Commercial Fishing", "People's Desire to Save Sharks", "Damage to Gear")),
    Aggregated_30 = list(
      Ecological = c("Shark Populations", "Habitat & Ecosystem Quality", "Shark Attraction To Fishing Activity"),
      Management = c("Commercial Shark Fisheries", "Economic Losses from Depredation", "Regulatory Framework"),
      Fisheries_Behavior = c("Shark Fin Market and Demand", "Public Perception & Media", "Recreational Fishing Hotspots")))
  THEME_LABEL <- c(Ecological = "Ecological", Management = "Management", Fisheries_Behavior = "Fisheries / Behaviour")
  OUTCOMES <- list(
    "Shark Depredation"  = c(Recreational = "Shark Depredation", Charter = "SHARK DEPREDATION", Commercial = "Shark Depredation", Aggregated_30 = "Shark Depredation"),
    "Gear Damage"        = c(Recreational = "Damage to Gear", Charter = "Damage to Gear", Commercial = "Damage to Gear", Aggregated_30 = "Gear & Vessel Damage"),
    "Public Perception"  = c(Recreational = "Public Positive Perception", Charter = "Public Positive Perception", Commercial = "Public Positive Perception", Aggregated_30 = "Public Perception & Media"),
    "Fishermen's Income" = c(Recreational = "Fishermen's Income", Charter = NA, Commercial = "Fishermen's Income", Aggregated_30 = "Fishermen's Income"))
  OUTCOME_ORDER <- names(OUTCOMES)
  rows <- list()
  for (model in names(SCENARIOS)) {
    W <- loaded[[model]]$W; concepts <- loaded[[model]]$concepts
    for (theme in names(SCENARIOS[[model]])) for (lever in SCENARIOS[[model]][[theme]]) {
      res <- run_scenario(W, concepts, c(lever), clamp_value = 1.0)
      for (outcome in OUTCOME_ORDER) {
        target <- OUTCOMES[[outcome]][[model]]
        val <- if (is.na(target) || identical(target, lever)) NA_real_ else as.numeric(res$change[[target]])
        rows[[length(rows) + 1]] <- data.frame(Model = model, ModelLabel = MODEL_LABEL[[model]], Theme = theme,
                                               ThemeLabel = THEME_LABEL[[theme]], Lever = lever, Outcome = outcome, Change = val, stringsAsFactors = FALSE)
      }
    }
  }
  tidy <- do.call(rbind, rows)
  tidy$ModelLabel <- factor(tidy$ModelLabel, levels = MODEL_ORDER)
  tidy$ThemeLabel <- factor(tidy$ThemeLabel, levels = THEME_LABEL[c("Ecological", "Management", "Fisheries_Behavior")])
  tidy$Outcome    <- factor(tidy$Outcome, levels = OUTCOME_ORDER)
  tidy$Norm <- tidy$Change / ave(abs(tidy$Change), tidy$Model, FUN = function(x) { m <- max(x, na.rm = TRUE); if (is.finite(m) && m > 0) m else 1 })
  write.csv(tidy, file.path(FIG_DIR, "scenario_shared_outcome_values.csv"), row.names = FALSE)
  lever_key <- function(d) paste(d$Model, d$Theme, d$Lever)
  build_fig1 <- function(value_col, xlab, fname, subtitle) {
    d <- tidy[tidy$Outcome == "Shark Depredation", ]; d$val <- d[[value_col]]
    d <- d[order(d$ThemeLabel, d$ModelLabel, -abs(d$val)), ]
    d$key <- factor(lever_key(d), levels = rev(unique(lever_key(d)))); d$LeverShort <- short_lab(d$Lever, 46)
    p <- ggplot(d, aes(x = val, y = key, fill = ModelLabel)) +
      geom_col(width = 0.7) + geom_vline(xintercept = 0, colour = "grey40", linewidth = 0.3) +
      facet_grid(ThemeLabel ~ ., scales = "free_y", space = "free_y") +
      scale_fill_manual(values = SECTOR_COLS, name = "Sector / model", drop = FALSE) +
      scale_y_discrete(labels = function(x) d$LeverShort[match(x, d$key)]) +
      labs(title = "Effect of each scenario on Shark Depredation", subtitle = subtitle, x = xlab, y = NULL) + theme_pub()
    save_fig(p, fname, 10, 9)
  }
  build_fig1("Change", "Change in Shark Depredation (from baseline)", "fig1_depredation_effect_abs", "Each bar = one activated driver; grouped by theme, coloured by sector/model")
  build_fig1("Norm", "Change in Shark Depredation (normalised per model, −1 to 1)", "fig1_depredation_effect_norm", "Per-model normalised (each model scaled to its own maximum |change|)")
  build_fig2 <- function(value_col, fillname, fname, subtitle, limit) {
    d <- tidy; d$val <- d[[value_col]]
    d <- d[order(d$ModelLabel, d$Theme, d$Lever), ]
    d$key <- factor(lever_key(d), levels = rev(unique(lever_key(d)))); d$LeverShort <- short_lab(d$Lever, 44)
    p <- ggplot(d, aes(x = Outcome, y = key, fill = val)) +
      geom_tile(colour = "white", linewidth = 0.5) +
      geom_text(aes(label = ifelse(is.na(val), "", sprintf("%.2f", val))), size = 2.4, colour = "grey15") +
      facet_grid(ModelLabel ~ ., scales = "free_y", space = "free_y") +
      scale_fill_gradient2(low = DIV_LOW, mid = DIV_MID, high = DIV_HIGH, midpoint = 0,
                           limits = c(-limit, limit), oob = scales::squish, na.value = "grey88", name = fillname) +
      scale_y_discrete(labels = function(x) d$LeverShort[match(x, d$key)]) +
      labs(title = "Scenario effects on shared outcome concepts, by sector", subtitle = subtitle, x = NULL, y = NULL) +
      theme_pub(base = 10) + theme(axis.text.x = element_text(angle = 20, hjust = 1), legend.key.width = unit(1.3, "cm"))
    save_fig(p, fname, 8.5, 11)
  }
  build_fig2("Change", "Change from baseline\n(orange = decrease, purple = increase)", "fig2_shared_outcomes_heatmap_abs", "Grey = concept absent from that model. Values are absolute change from baseline.", 0.20)
  build_fig2("Norm", "Normalised change\n(orange = decrease, purple = increase)", "fig2_shared_outcomes_heatmap_norm", "Per-model normalised to each model's maximum |change|. Grey = concept absent.", 1.0)
  build_fig3 <- function(value_col, xlab, fname, subtitle) {
    d <- tidy[tidy$Theme == "Management" & tidy$Outcome == "Shark Depredation", ]; d$val <- d[[value_col]]
    d <- d[order(d$ModelLabel, -abs(d$val)), ]
    d$key <- factor(lever_key(d), levels = rev(unique(lever_key(d)))); d$LeverShort <- short_lab(d$Lever, 50)
    p <- ggplot(d, aes(x = val, y = key, fill = ModelLabel)) +
      geom_col(width = 0.7) + geom_vline(xintercept = 0, colour = "grey40", linewidth = 0.3) +
      scale_fill_manual(values = SECTOR_COLS, name = "Sector / model", drop = FALSE) +
      scale_y_discrete(labels = function(x) d$LeverShort[match(x, d$key)]) +
      labs(title = "Management measures: effect on Shark Depredation", subtitle = subtitle, x = xlab, y = NULL) + theme_pub()
    save_fig(p, fname, 9, 5)
  }
  build_fig3("Change", "Change in Shark Depredation (from baseline)", "fig3_management_spotlight_abs", "Management-theme levers across all four models")
  build_fig3("Norm", "Change in Shark Depredation (normalised per model)", "fig3_management_spotlight_norm", "Per-model normalised to each model's maximum |change|")
})


# -----------------------------------------------------------------------------
#   3c. Thematic scenarios — heatmaps                                       [11]
#       Additive headline (4 model panels x 3 theme bundles) + one detail
#       heatmap per model (individual activations). Category-coloured labels,
#       single shared legend, blue<->orange diverging fill.
# -----------------------------------------------------------------------------
cat("\n\n=====================  SECTION 3c  Thematic heatmaps (11)  =====================\n")
local({
  library(ggplot2); library(scales); library(patchwork)
  DATA_DIR <- file.path("..", "data")
  OUT_DIR  <- file.path("..", "results", "figures_themed_Scenarios")
  dir.create(OUT_DIR, showWarnings = FALSE, recursive = TRUE)
  INC <- "#E08214"; DEC <- "#2166AC"
  GROUP_COLOURS <- c(
    "Central Concept"="#C9A227","Ecological & Biological Factors"="#9AD354","Human Dimensions"="#F8895E",
    "Fisheries Operations & Practices"="#E382BA","Fisheries Research & Management"="#8695C2",
    "Policy & Economics"="#5EB99B","Other"="grey60")
  sl <- function(x, n = 40) ifelse(nchar(x) > n, paste0(substr(x, 1, n - 1), "…"), x)
  theme_pub <- function(b = 9) theme_minimal(base_size = b) +
    theme(plot.title = element_text(face = "bold", hjust = 0.5), panel.grid = element_blank(),
          strip.text = element_text(face = "bold"), legend.position = "right")
  wrap2 <- function(x, width = 16) vapply(x, function(s) paste(strwrap(as.character(s), width = width), collapse = "\n"), character(1), USE.NAMES = FALSE)
  DISPLAY_NAME <- c(
    "Participation in Tournaments - Shark mutilate catch Ex: BigRock in NC" = "Participation in Tournaments",
    "Fisher Satisfaction with Specific Species Fishing"                     = "Species Specific Fisher Satisfaction",
    "Changing Species Targeted by Recreational Fishing"                     = "Change in Target Recreational Species",
    "Lack of economic incentives to overcome the impacts of depredation"    = "Lack of Economic Incentives",
    "Shifting Distributions of Shark Populations"                           = "Shifting Shark Distributions",
    "Habitat Loss (Removal of rigs and platforms)"                          = "Habitat Loss",
    "Shark Finning Law - Illegal to separate fins from body"                = "Shark Finning Law",
    "Shark vessel–food association (attraction and following fishing boats due to learned food rewards)" = "Shark vessel-food association",
    "Recreational fishing boats following commercial fishing vessels. This attracts sharks to shrimper boats" = "Recreational fishing boats following commercial boats")
  relabel <- function(x) { y <- unname(DISPLAY_NAME[x]); ifelse(is.na(y), x, y) }
  CATEGORY <- c(
    "Ability to Identify Shark Species"="Human Dimensions","BP oil spill reduced prey populations (shark's food)"="Ecological & Biological Factors",
    "Bycatch Reduction Devices (Shrimp fishery)"="Fisheries Operations & Practices","Catch In Commercial Fishing"="Fisheries Operations & Practices",
    "Catch and Release Trips"="Fisheries Operations & Practices","Change Fishing Locations"="Fisheries Operations & Practices",
    "Changing Species Targeted by Recreational Fishing"="Fisheries Operations & Practices","Client Satisfaction"="Human Dimensions",
    "Climate Change"="Ecological & Biological Factors","Commercial Shark Fisheries"="Fisheries Research & Management",
    "Cost of Fines from Gear Damage (shrimpers)"="Policy & Economics","Crew Safety"="Human Dimensions",
    "Damage to Gear"="Fisheries Operations & Practices","Damage to Vessels by the Sharks Frenzy"="Fisheries Operations & Practices",
    "Demand for Sharks"="Human Dimensions","Dolphin Populations"="Ecological & Biological Factors","Economic Loss"="Policy & Economics",
    "Economic Losses from Depredation"="Policy & Economics","Efficiency of Commercial Fishery"="Fisheries Research & Management",
    "Effort in Cleaning and Processing Sharks"="Fisheries Operations & Practices","Environmental & Climate Drivers"="Ecological & Biological Factors",
    "Fisher & Client Satisfaction"="Human Dimensions","Fisher Satisfaction with Specific Species Fishing"="Human Dimensions",
    "Fisheries Management Effectiveness"="Fisheries Research & Management","Fisheries Management Effectiveness (Reef Fish)"="Fisheries Research & Management",
    "Fishermen's Income"="Human Dimensions","Fishing Behavioral Adaptations"="Fisheries Operations & Practices","Fishing Profitability"="Policy & Economics",
    "Fishing Time"="Fisheries Operations & Practices","Gear & Vessel Damage"="Policy & Economics","Gear Costs"="Policy & Economics",
    "Government Enforcement and Federal Regulations"="Policy & Economics","HMS Management"="Fisheries Research & Management",
    "HMS Management actions"="Fisheries Research & Management","HMS Slow Management Response"="Fisheries Research & Management",
    "Habitat & Ecosystem Quality"="Ecological & Biological Factors","Habitat Loss (Removal of rigs and platforms)"="Ecological & Biological Factors",
    "Health Conscious"="Human Dimensions","Hooked Fish Behavior"="Ecological & Biological Factors","Increase Shark Quotas"="Policy & Economics",
    "Increased Fishing Interest in Social Media"="Human Dimensions","Increased Number of Artificial Reefs"="Ecological & Biological Factors",
    "Insurance & Charter Risks"="Fisheries Operations & Practices","Insurance Risks for Charters"="Fisheries Operations & Practices",
    "Interest by Charter Clients"="Human Dimensions","Knowledge & Awareness"="Human Dimensions","Knowledge of Depredation"="Human Dimensions",
    "Labor Time to Repair Gear"="Policy & Economics","Lack of economic incentives to overcome the impacts of depredation"="Policy & Economics",
    "Legislation"="Policy & Economics","More Frequent Fishing Spots (Hotspots)"="Fisheries Operations & Practices","More Recreational Fishermen"="Human Dimensions",
    "Number Of Shrimp Boats"="Fisheries Operations & Practices","Participation in Tournaments - Shark mutilate catch Ex: BigRock in NC"="Human Dimensions",
    "People's Desire to Save Sharks"="Human Dimensions","Public Fear - Dangling your feet"="Human Dimensions","Public Perception & Media"="Human Dimensions",
    "Public Positive Perception"="Human Dimensions","Recreational Fishing Hotspots"="Fisheries Operations & Practices","Recreational Participation"="Human Dimensions",
    "Recreational fishing boats following commercial fishing vessels. This attracts sharks to shrimper boats"="Fisheries Operations & Practices",
    "Regulatory Framework"="Policy & Economics","Research and Funding"="Fisheries Research & Management","Safety & Social Concerns"="Human Dimensions",
    "Shark Attraction To Fishing Activity"="Ecological & Biological Factors","Shark Attraction to Vessels"="Ecological & Biological Factors",
    "Shark Behavioral Responses"="Ecological & Biological Factors","Shark Conservation"="Policy & Economics","Shark Conservation & Quota Policy"="Policy & Economics",
    "Shark Depredation"="Central Concept","Shark Fin Market"="Policy & Economics","Shark Fin Market and Demand"="Policy & Economics",
    "Shark Finning Law - Illegal to separate fins from body"="Policy & Economics","Shark Fishery"="Fisheries Research & Management",
    "Shark Learning Behavior"="Ecological & Biological Factors","Shark Populations"="Ecological & Biological Factors",
    "Shark vessel–food association (attraction and following fishing boats due to learned food rewards)"="Ecological & Biological Factors",
    "Shifting Distributions of Shark Populations"="Ecological & Biological Factors","Shrimping Operations"="Fisheries Operations & Practices",
    "Switch from Bait to Lures/Jigs"="Fisheries Operations & Practices","Water Temperature"="Ecological & Biological Factors")
  cat_of <- function(x) { v <- unname(CATEGORY[x]); v[is.na(v)] <- "Other"; v }
  MODEL_PATHS <- list(
    Recreational  = file.path(DATA_DIR, "Recreational_FinalModel_22April.csv"),
    Charter       = file.path(DATA_DIR, "Charter_FinalModel_22April.csv"),
    Commercial    = file.path(DATA_DIR, "Commercial_FinalModel_22April.csv"),
    Aggregated_30 = file.path(DATA_DIR, "MentalModeler_Galveston_Reduced30.csv"))
  MLAB <- c(Recreational = "Recreational", Charter = "Charter", Commercial = "Commercial", Aggregated_30 = "Reduced Aggregated Model")
  loaded <- lapply(MODEL_PATHS, load_adjacency_matrix)
  fix_names <- function(x) { x[x == "Research & Funding"] <- "Research and Funding"; x[x == "SHARK DEPREDATION"] <- "Shark Depredation"; x }
  for (m in names(loaded)) { loaded[[m]]$concepts <- fix_names(loaded[[m]]$concepts); dimnames(loaded[[m]]$W) <- list(loaded[[m]]$concepts, loaded[[m]]$concepts) }
  SCENARIOS <- list(
    Recreational = list(
      Ecological = c("Habitat Loss (Removal of rigs and platforms)"),
      Management = c("Shark Finning Law - Illegal to separate fins from body", "Commercial Shark Fisheries", "HMS Slow Management Response"),
      Fisheries_Behavior = c("Demand for Sharks", "Public Positive Perception", "More Frequent Fishing Spots (Hotspots)")),
    Charter = list(
      Ecological = c("Shark Populations", "Habitat Loss (Removal of rigs and platforms)", "Shark vessel–food association (attraction and following fishing boats due to learned food rewards)"),
      Management = c("HMS Management actions", "Shark Conservation", "Increase Shark Quotas"),
      Fisheries_Behavior = c("Knowledge of Depredation", "Damage to Gear", "Change Fishing Locations")),
    Commercial = list(
      Ecological = c("Recreational fishing boats following commercial fishing vessels. This attracts sharks to shrimper boats", "Shark Populations"),
      Management = c("Commercial Shark Fisheries"),
      Fisheries_Behavior = c("Catch In Commercial Fishing", "People's Desire to Save Sharks", "Damage to Gear")),
    Aggregated_30 = list(
      Ecological = c("Shark Populations", "Habitat & Ecosystem Quality", "Shark Attraction To Fishing Activity"),
      Management = c("Commercial Shark Fisheries", "Economic Losses from Depredation", "Regulatory Framework"),
      Fisheries_Behavior = c("Shark Fin Market and Demand", "Public Perception & Media", "Recreational Fishing Hotspots")))
  THEME_LABEL <- c(Ecological = "Ecological", Management = "Management", Fisheries_Behavior = "Fisheries / Behaviour")
  change_of <- function(model, activate) run_scenario(loaded[[model]]$W, loaded[[model]]$concepts, activate, clamp_value = 1.0)$change
  additive <- list(); individual <- list()
  for (m in names(SCENARIOS)) {
    additive[[m]] <- list(); individual[[m]] <- list()
    for (th in names(SCENARIOS[[m]])) {
      additive[[m]][[th]] <- change_of(m, SCENARIOS[[m]][[th]]); individual[[m]][[th]] <- list()
      for (c in SCENARIOS[[m]][[th]]) individual[[m]][[th]][[c]] <- change_of(m, c)
    }
  }
  LIMIT <- 0.20
  make_heatmap <- function(d, xlab, title, top_rows = 13, facet = FALSE, wrap_x = FALSE) {
    score <- tapply(abs(d$Value), d$Concept, function(x) max(x, na.rm = TRUE)); score[!is.finite(score)] <- 0
    keep <- names(sort(score, decreasing = TRUE)); keep <- head(keep[score[keep] > 0.005], top_rows)
    dep <- "Shark Depredation"; if (dep %in% d$Concept) keep <- union(keep, dep)
    d <- d[d$Concept %in% keep, ]
    ord <- names(sort(tapply(abs(d$Value), d$Concept, function(x) sum(x, na.rm = TRUE)), decreasing = TRUE))
    if (dep %in% ord) ord <- c(dep, setdiff(ord, dep))
    d$Concept  <- factor(d$Concept, levels = rev(ord))
    d$Category <- factor(cat_of(as.character(d$Concept)), levels = names(GROUP_COLOURS))
    ycols <- unname(GROUP_COLOURS[cat_of(levels(d$Concept))]); ycols[is.na(ycols)] <- "grey40"
    p <- ggplot(d, aes(x = Xkey, y = Concept, fill = Value)) +
      geom_tile(colour = "white", linewidth = 0.4) +
      geom_text(aes(label = ifelse(is.na(Value), "", ifelse(abs(Value) == 0, "–", sprintf("%.4f", Value)))), size = 2.1, colour = "grey20") +
      geom_point(aes(colour = Category), alpha = 0, na.rm = TRUE) +
      scale_fill_gradient2(low = DEC, mid = "#F7F7F7", high = INC, midpoint = 0,
                           limits = c(-LIMIT, LIMIT), oob = scales::squish, na.value = "grey88", name = "Change") +
      scale_colour_manual(values = GROUP_COLOURS, limits = names(GROUP_COLOURS), drop = FALSE, name = "Category") +
      scale_y_discrete(labels = function(x) sl(relabel(x), 38)) +
      scale_x_discrete(labels = function(x) if (wrap_x) wrap2(x, 16) else x) +
      guides(fill = guide_colourbar(order = 1),
             colour = guide_legend(order = 2, nrow = 2, byrow = TRUE, override.aes = list(alpha = 1, shape = 15, size = 4))) +
      labs(title = title, x = xlab, y = NULL) + theme_pub(9) +
      theme(axis.text.y = element_text(colour = ycols, face = "bold"),
            axis.text.x = element_text(angle = 0, hjust = 0.5, size = 7, lineheight = 0.85, margin = margin(t = 4)),
            axis.title.x = element_text(margin = margin(t = 14)), panel.spacing.x = unit(0.8, "lines"))
    if (facet) p <- p + facet_grid(~ Theme, scales = "free_x", space = "free_x")
    p
  }
  get_legend2 <- function(p) {
    g <- ggplotGrob(p); idx <- which(grepl("guide-box", g$layout$name))
    for (i in idx) { gb <- g$grobs[[i]]; if (!inherits(gb, "zeroGrob")) return(gb) }
    if (length(idx)) g$grobs[[idx[1]]] else grid::nullGrob()
  }
  save_paneled <- function(plots, title, fname, w, h, ncol = 2, leg_frac = 0.14) {
    dd <- data.frame(Category = factor(names(GROUP_COLOURS), levels = names(GROUP_COLOURS)),
                     Value = seq(-LIMIT, LIMIT, length.out = length(GROUP_COLOURS)))
    donor <- ggplot(dd, aes(Category, Category)) +
      geom_tile(aes(fill = Value)) + geom_point(aes(colour = Category)) +
      scale_fill_gradient2(low = DEC, mid = "#FFFFFF", high = INC, midpoint = 0, limits = c(-LIMIT, LIMIT), name = "Change") +
      scale_colour_manual(values = GROUP_COLOURS, name = "Category") +
      guides(fill = guide_colourbar(order = 1, barwidth = unit(6, "cm"), barheight = unit(0.5, "cm")),
             colour = guide_legend(order = 2, nrow = 2, byrow = TRUE, override.aes = list(shape = 15, size = 4))) +
      theme(legend.position = "bottom", legend.box = "vertical", legend.justification = "center",
            legend.key = element_rect(fill = "white", colour = NA))
    legend_grob <- get_legend2(donor)
    stripped <- lapply(plots, function(p) p + theme(legend.position = "none"))
    grid <- wrap_plots(stripped, ncol = ncol)
    combined <- wrap_plots(grid, wrap_elements(full = legend_grob), ncol = 1, heights = c(1, leg_frac)) +
      plot_annotation(title = title, theme = theme(plot.title = element_text(face = "bold", hjust = 0.5)))
    ggsave(file.path(OUT_DIR, paste0(fname, ".png")), combined, width = w, height = h, dpi = 350, limitsize = FALSE)
    tryCatch(ggsave(file.path(OUT_DIR, paste0(fname, ".pdf")), combined, width = w, height = h, device = grDevices::cairo_pdf, limitsize = FALSE),
             error = function(e) ggsave(file.path(OUT_DIR, paste0(fname, ".pdf")), combined, width = w, height = h, limitsize = FALSE))
    cat(sprintf("Saved: %s.{png,pdf}\n", fname))
  }
  # Headline: additive bundles (4 model panels, columns = 3 themes)
  panels <- lapply(names(SCENARIOS), function(m) {
    concepts <- loaded[[m]]$concepts; rows <- list()
    for (th in names(SCENARIOS[[m]])) {
      bundle <- SCENARIOS[[m]][[th]]; ch <- additive[[m]][[th]]
      for (cpt in concepts) rows[[length(rows) + 1]] <- data.frame(Concept = cpt, Xkey = THEME_LABEL[[th]],
                                                                   Value = if (cpt %in% bundle) NA_real_ else ch[[cpt]], stringsAsFactors = FALSE)
    }
    d <- do.call(rbind, rows); d$Xkey <- factor(d$Xkey, levels = unname(THEME_LABEL[names(SCENARIOS[[m]])]))
    make_heatmap(d, xlab = NULL, title = MLAB[[m]], top_rows = 12)
  })
  save_paneled(panels, "Additive thematic scenarios: effect of activating each theme bundle",
               "fig_thematic_additive_headline", w = 14, h = 12)
  # Detail: one figure per model (individual activations)
  for (m in names(SCENARIOS)) {
    concepts <- loaded[[m]]$concepts; rows <- list()
    for (th in names(SCENARIOS[[m]])) for (act in SCENARIOS[[m]][[th]]) {
      ch <- individual[[m]][[th]][[act]]
      for (cpt in concepts) rows[[length(rows) + 1]] <- data.frame(Concept = cpt, Xkey = sl(relabel(act), 40),
                                                                   Theme = THEME_LABEL[[th]], Value = if (cpt == act) NA_real_ else ch[[cpt]], stringsAsFactors = FALSE)
    }
    d <- do.call(rbind, rows)
    xlev <- sl(relabel(unlist(lapply(names(SCENARIOS[[m]]), function(t) SCENARIOS[[m]][[t]]), use.names = FALSE)), 40)
    d$Xkey  <- factor(d$Xkey, levels = unique(xlev)); d$Theme <- factor(d$Theme, levels = unname(THEME_LABEL[names(SCENARIOS[[m]])]))
    panel <- make_heatmap(d, xlab = "Activated concept", title = sprintf("%s - individual thematic scenarios", MLAB[[m]]),
                          top_rows = 15, facet = TRUE, wrap_x = TRUE)
    save_paneled(list(panel), title = NULL, fname = sprintf("fig_thematic_individual_%s", m), w = 15, h = 11, ncol = 1, leg_frac = 0.24)
  }
  cat("Thematic heatmaps written to: ", normalizePath(OUT_DIR), "\n", sep = "")
})


# #############################################################################
#   DONE
# #############################################################################
cat("\n\n", strrep("=", 79), "\n", sep = "")
cat("PIPELINE COMPLETE. Outputs in ../results/ :\n")
cat("  step1_depredation_scenario.xlsx        (depredation tables)          [03]\n")
cat("  step3_thematic_scenarios_revised.xlsx  (individual thematic)         [07]\n")
cat("  step3_additive_scenarios.xlsx          (additive bundles)            [07]\n")
cat("  step3_individual_vs_additive.xlsx      (comparison)                  [07]\n")
cat("  figures/                 depredation bars (03) + thematic bars (07)\n")
cat("  figures_step1/           depredation manuscript figs (09) + activation heatmap (10)\n")
cat("  figures_manuscript/      cross-model thematic figs (08)\n")
cat("  figures_themed_Scenarios/ thematic heatmaps (11)\n")
cat(strrep("=", 79), "\n", sep = "")
