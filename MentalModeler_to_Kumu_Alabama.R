# =============================================================================
# MentalModeler_to_Kumu_Alabama.R
# Converts Mental Modeler adjacency-matrix CSV files (one per US state fisher
# group from the Alabama workshop) into Kumu-formatted Excel files.
#
# Input : 5 Mental Modeler CSV files:
#           Alabama.csv, Florida.csv, Louisiana.csv,
#           Mississippi.csv, Texas.csv
# Output: 5 individual Kumu Excel files  (one per state)
#         1 Master Kumu Excel file       (all states combined; Alabama workshop)
#
# Author: generated with Claude (Anthropic) — April 2026
# Usage:  Set input_dir and output_dir below, then source() or Rscript
# =============================================================================

# =============================================================================
# ASSUMPTIONS & DECISIONS DOCUMENTED
# =============================================================================
#
# [A1] DATA FORMAT — INPUT
#      Input files are Mental Modeler adjacency-matrix CSV exports.
#      Row label (first column) = FROM concept (influencer).
#      Column header = TO concept (influenced).
#      Non-empty cell = directed connection with that signed weight.
#      Empty cell  = no connection.
#      The matrix is square (same set of concepts as rows and columns).
#
# [A2] CONCEPT NAME NORMALISATION
#      All five CSV files use consistent Title Case concept labels.
#      - Individual Kumu files: labels kept AS-IS from each CSV.
#      - Master Kumu: ALL concept labels converted to Title Case
#        (stringr::str_to_title) before cross-state matching, ensuring
#        minor capitalisation differences do not create spurious duplicates.
#      NOTE: In an earlier version of these data, Louisiana used all-lowercase
#      labels and contained non-standard weights (see A3 below). Both issues
#      have been corrected in the current CSV files and the notes below
#      reflect the corrected data.
#
# [A3] NON-STANDARD STRENGTH VALUES
#      All five state CSVs use the standard Mental Modeler values ±0.5 / ±1.0.
#      No non-standard weights are present in the current input files.
#      The 0.75-threshold binning logic is retained in the mm_to_kumu()
#      function as a safety net and will report a warning to the console
#      if any non-standard value is detected in future data.
#
# [A4] STRENGTH BACK-CONVERSION TO KUMU SCALE
#      Mental Modeler uses ±0.5 / ±1.0. The Australia workflow uses ±1 / ±2
#      in Kumu. Mapping (consistent with Australia's reverse conversion):
#        MM  +0.5  →  Kumu +1  (Moderate Positive)
#        MM  +1.0  →  Kumu +2  (Strong Positive)
#        MM  −0.5  →  Kumu −1  (Moderate Negative)
#        MM  −1.0  →  Kumu −2  (Strong Negative)
#      Sign is preserved: positive weight → increases target;
#                         negative weight → decreases target.
#
# [A5] CATEGORY ASSIGNMENT
#      Mental Modeler CSVs do not carry thematic category data.
#      Categories are assigned programmatically using a hardcoded lookup
#      table (category_lookup) following the same five-category scheme
#      used for the Australia workshop, plus a "Central Concept" class:
#        "Central Concept"
#        "Ecological & Biological Factors"
#        "Fisheries Operations & Practices"
#        "Human Dimensions"
#        "Policy & Economics"
#        "Fisheries Research & Management"
#      Category lookup keys are Title Case (str_to_title-normalised).
#      Any concept not found in the lookup table will receive NA and
#      a warning will be printed to the console.
#      See the category_lookup tribble below for all assignments.
#
# [A6] "DEPREDATION" AS CENTRAL NODE
#      "Depredation" is present as both a row and column in every state's
#      matrix (normalised to Title Case in the Master file). It is included
#      as a standard element in each state's Elements sheet and in the Master
#      Elements sheet, tagged accordingly. It is NOT given special treatment
#      beyond the Title Case normalisation applied to all labels in [A2].
#
# [A7] "SHARK DEPREDATION" vs "DEPREDATION" — FLORIDA
#      Florida's map contains BOTH "Depredation" and "Shark Depredation" as
#      distinct concepts (connected: Shark Depredation → Depredation = +1).
#      These are kept as SEPARATE concepts and are NOT merged.
#      Their co-existence likely reflects the Florida participant group
#      differentiating the general depredation phenomenon from its shark-
#      specific form.
#
# [A8] SELF-LOOPS
#      Connections where From == To (after normalisation) are removed.
#      None were present in the source data.
#
# [A9] DEDUPLICATION IN MASTER — CONNECTIONS
#      Unique connection key: (From_norm, To_norm, Kumu_Strength).
#        Same key across multiple states → ONE row; Tags = pipe-joined states.
#        Same (From_norm, To_norm) with DIFFERENT Kumu_Strength → SEPARATE rows
#        (preserves diversity of stakeholder perceptions).
#      Both cases are reported in the console output.
#
# [A10] DEDUPLICATION IN MASTER — ELEMENTS
#       One row per unique concept label (Title Case). Tags = pipe-separated
#       list of all states where that concept appeared.
#
# [A11] TAG FORMAT
#       Tags use | (pipe, NO surrounding spaces), e.g. "Alabama|Florida".
#       Consistent with Australia Kumu workflow; prevents leading-space
#       parsing errors in Kumu.
#       - Individual files: Tags column = single state name (e.g. "Alabama")
#         for both Elements and Connections.
#       - Master file: Tags column = pipe-joined list of all states that share
#         that element or connection, sorted alphabetically.
#
# [A12] WORKSHOP CONTEXT
#       All five CSVs originate from a single workshop held in Alabama,
#       where participants were grouped by their home US state.
#       The five state tags therefore represent within-workshop subgroups,
#       not separate workshops.
#
# =============================================================================

library(readr)
library(dplyr)
library(tidyr)
library(stringr)
library(openxlsx)

# =============================================================================
# PARAMETERS — edit these as needed
# =============================================================================


input_dir  <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Alabama/MM/"   # Directory containing the 5 CSV files
output_dir <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Alabama/Kumu/" # Directory where output Excel files will be written

states <- c("Alabama", "Florida", "Louisiana", "Mississippi", "Texas")

# Each CSV is named <State>.csv
input_files <- setNames(
  file.path(input_dir, paste0(states, ".csv")),
  states
)

# =============================================================================
# STEP 1 — Helper functions
# =============================================================================

# ---- 1a. Read one Mental Modeler CSV → list(concepts, connections) ----------
read_mm_csv <- function(path, state_name) {
  raw <- read_csv(path, col_names = TRUE, show_col_types = FALSE)
  
  # Column 1 contains row (FROM) concept names; columns 2..n are TO concepts
  from_labels <- as.character(raw[[1]])
  to_labels   <- names(raw)[-1]         # header row, skip first ""
  
  # Build long-format connection table
  conns <- raw %>%
    rename(From_Label = 1) %>%
    mutate(From_Label = as.character(From_Label)) %>%
    pivot_longer(
      cols      = -From_Label,
      names_to  = "To_Label",
      values_to = "MM_Weight"
    ) %>%
    mutate(
      MM_Weight = suppressWarnings(as.numeric(MM_Weight)),
      State     = state_name
    ) %>%
    filter(!is.na(MM_Weight), MM_Weight != 0)   # keep only actual connections
  
  # Remove self-loops  [A8]
  n_self <- sum(conns$From_Label == conns$To_Label)
  if (n_self > 0) {
    message(sprintf("  [A8] %s: removed %d self-loop(s).", state_name, n_self))
    conns <- conns %>% filter(From_Label != To_Label)
  }
  
  all_concepts <- unique(c(from_labels, to_labels))
  
  list(
    concepts    = all_concepts,
    connections = conns
  )
}

# ---- 1b. Convert MM weight → Kumu signed strength  [A3][A4] -----------------
mm_to_kumu <- function(w) {
  # Bin by 0.75 threshold, preserve sign
  abs_w      <- abs(w)
  snapped_mm <- if_else(abs_w < 0.75, 0.5, 1.0)   # snap to nearest MM standard
  kumu_mag   <- if_else(snapped_mm == 0.5, 1L, 2L) # MM 0.5 → Kumu 1; MM 1 → Kumu 2
  as.integer(sign(w)) * kumu_mag
}

# ---- 1c. Strength label  -----------------------------------------------------
kumu_label <- function(k) {
  case_when(
    k ==  2L ~ "Strong Positive",
    k ==  1L ~ "Moderate Positive",
    k == -1L ~ "Moderate Negative",
    k == -2L ~ "Strong Negative",
    TRUE     ~ as.character(k)
  )
}

# ---- 1d. Excel styling  -------------------------------------------------------
h_style  <- createStyle(fontName = "Arial", fontSize = 11, fontColour = "#FFFFFF",
                        fgFill = "#2F4F4F", halign = "CENTER",
                        textDecoration = "Bold", wrapText = TRUE)
d_style  <- createStyle(fontName = "Arial", fontSize = 10)
pos_sty  <- createStyle(fontName = "Arial", fontSize = 10, fgFill = "#C8E6C9")
neg_sty  <- createStyle(fontName = "Arial", fontSize = 10, fgFill = "#FFCDD2")

# ---- 1e. Write a formatted Kumu workbook  ------------------------------------
write_kumu_wb <- function(elements_df, connections_df, notes_df,
                          output_path, sheet_label = "State") {
  wb <- createWorkbook()
  
  # ── Elements sheet ──────────────────────────────────────────────────────
  addWorksheet(wb, "Elements")
  writeData(wb, "Elements", elements_df)
  n_el <- nrow(elements_df)
  addStyle(wb, "Elements", h_style, rows = 1, cols = seq_len(ncol(elements_df)))
  addStyle(wb, "Elements", d_style, rows = 2:(n_el + 1),
           cols = seq_len(ncol(elements_df)), gridExpand = TRUE)
  setColWidths(wb, "Elements",
               cols   = seq_len(ncol(elements_df)),
               widths = c(44, 38, 56))  # Label, Category, Tags
  freezePane(wb, "Elements", firstRow = TRUE)
  
  # ── Connections sheet ────────────────────────────────────────────────────
  addWorksheet(wb, "Connections")
  writeData(wb, "Connections", connections_df)
  n_cn  <- nrow(connections_df)
  n_col <- ncol(connections_df)
  addStyle(wb, "Connections", h_style, rows = 1, cols = seq_len(n_col))
  pos_rows <- which(connections_df$Strength > 0) + 1
  neg_rows <- which(connections_df$Strength < 0) + 1
  if (length(pos_rows) > 0)
    addStyle(wb, "Connections", pos_sty,
             rows = pos_rows, cols = seq_len(n_col), gridExpand = TRUE)
  if (length(neg_rows) > 0)
    addStyle(wb, "Connections", neg_sty,
             rows = neg_rows, cols = seq_len(n_col), gridExpand = TRUE)
  setColWidths(wb, "Connections",
               cols   = seq_len(n_col),
               widths = c(44, 44, 10, 18, 22, 16, 56))
  freezePane(wb, "Connections", firstRow = TRUE)
  
  # ── Notes sheet ─────────────────────────────────────────────────────────
  addWorksheet(wb, "Notes & Assumptions")
  writeData(wb, "Notes & Assumptions", notes_df)
  bold_s  <- createStyle(fontName = "Arial", fontSize = 10, textDecoration = "Bold")
  plain_s <- createStyle(fontName = "Arial", fontSize = 10)
  addStyle(wb, "Notes & Assumptions", bold_s,
           rows = 2:(nrow(notes_df) + 1), cols = 1, gridExpand = TRUE)
  addStyle(wb, "Notes & Assumptions", plain_s,
           rows = 2:(nrow(notes_df) + 1), cols = 2, gridExpand = TRUE)
  setColWidths(wb, "Notes & Assumptions", cols = 1:2, widths = c(28, 96))
  
  saveWorkbook(wb, output_path, overwrite = TRUE)
  cat(sprintf("  Saved: %s\n", basename(output_path)))
  cat(sprintf("    Elements   : %d rows\n", n_el))
  cat(sprintf("    Connections: %d rows\n", n_cn))
}

# =============================================================================
# CATEGORY LOOKUP TABLE  [A5]
# Keys: Title Case (str_to_title-normalised) concept labels.
# Values: one of six categories following the Australia workshop scheme.
#
# CATEGORY ASSIGNMENTS — ALABAMA WORKSHOP (all states)
# ─────────────────────────────────────────────────────────────────────────────
# CENTRAL CONCEPT
#   Depredation                              [all 5 states]
#   Shark Depredation                        [Florida only — kept separate per A7]
#
# ECOLOGICAL & BIOLOGICAL FACTORS
#   Climate                                  [Texas]
#   Dolphin Depredation                      [Florida]
#   Dolphin Learning Behavior                [Alabama, Florida, Mississippi]
#   Ecosystem Health                         [Florida]
#   Erratic Fish Behavior On Hook            [Alabama]
#   Forage Fish Populations                  [Florida]
#   Localized Red Snapper Depletion          [Mississippi]
#   Many Small Sharks                        [Texas]
#   Muddy Water                              [Alabama]
#   Natural Bottom / Shelf                   [Louisiana]
#   Number Of Sharks And Dolphins In The Area [Mississippi]
#   Prey Populations (Menhaden)              [Louisiana]
#   Red Snapper                              [Alabama]
#   Reef Fish Populations                    [Alabama, Florida, Louisiana, Texas]
#   Reef Size                                [Alabama]
#   Shark Aggregation                        [Texas]
#   Shark Learning Behavior                  [Florida, Texas]
#   Shark Pheromones/ Liver In Water         [Alabama]
#   Shark Populations                        [Alabama, Florida, Louisiana, Mississippi, Texas]
#   Shark Populations On Reefs               [Alabama]
#   Shifting Stocks                          [Texas]
#   Single Large Shark                       [Texas]
#   Target Species Prevalence At Fishing Location [Alabama]
#   Water Temperature                        [Florida]
#
# FISHERIES OPERATIONS & PRACTICES
#   Boat Signature                           [Alabama]
#   Concentrated Artificial Reefs            [Mississippi]
#   Discarding Small Fish                    [Mississippi]
#   Dolphin Deterrents                       [Florida]
#   Fight Time                               [Mississippi]
#   Fishing Effort Throughout Season         [Mississippi]
#   Higher Fishing Pressure On Reefs         [Mississippi]
#   Hooked Fish Size                         [Alabama, Florida]
#   Increased Reeling Time                   [Alabama]
#   Keeping Depredated Fish                  [Alabama]
#   Lights On Shrimping Boats                [Alabama]
#   Number Of Fishermen                      [Alabama, Florida]
#   Number Of Shrimp Boats                   [Mississippi, Texas]
#   Shark Fishery                            [Florida]
#   Time Of Day                              [Alabama]
#   Time On Water                            [Alabama]
#   Widely Distributed Artificial Reefs      [Mississippi]
#
# HUMAN DIMENSIONS
#   Angler Competency                        [Florida, Mississippi]
#   Angler Satisfaction                      [Florida, Louisiana, Mississippi]
#   Angler Stewardship                       [Florida]
#   Captain Satisfaction                     [Alabama]
#   Captain Stress                           [Alabama]
#   Charter Captain Satisfaction             [Florida]
#   Coastal Heritage                         [Florida]
#   Commercial Captain Satisfaction          [Florida]
#   Commercial Satisfaction                  [Texas]
#   Difficulty To Id Species                 [Mississippi]
#   Economic Loss                            [Alabama, Florida]
#   Gear Economic Loss                       [Florida]
#   Human Injury                             [Florida]
#   Knowledge Of Depredation                 [Louisiana]
#   Miscommunication Between Fishermen + Scientists [Florida]
#   Private Rec Angler Satisfaction          [Alabama, Texas]
#   Public Opinion Of Sharks + Dolphins      [Florida]
#   Safety                                   [Texas]
#   Shark Bites                              [Florida]
#   Tourist Angler Satisfaction              [Alabama]
#
# POLICY & ECONOMICS
#   Ability To Commercially Harvest Sharks   [Mississippi]
#   Complex Species-Specific Regulations     [Mississippi]
#   Diverse Markets For Shark Products       [Florida]
#   Dolphin Regulatory Protection            [Florida]
#   Economic Loss (Commercial)               [Texas]
#   Economic Loss (Recreational)             [Texas]
#   Federal Commercial Shark Industry        [Texas]
#   Federal Regulation                       [Texas]
#   Federal Water Shark Fishery (Rec)        [Texas]
#   Mirror State Limits For Shark Harvest In Fed Waters [Mississippi]
#   Red Snapper Size And Bag Limits          [Mississippi]
#   Regulatory Consistency (Sharks)          [Florida]
#   Season Length                            [Louisiana]
#   Separation Between Gulf Council And Hms/Stock Assessment Process [Florida]
#   Shark Messaging From Feds                [Florida]
#   Shark Regulatory Protection              [Florida]
#   Shark Revenue (From Selling Catch)       [Florida]
#   State Water Shark Fishery (Rec)          [Texas]
#   Texas Fin Ban                            [Mississippi]
#
# FISHERIES RESEARCH & MANAGEMENT
#   Data Gaps (Sandbar Stock Assessment Re-Evaluation) [Florida]
#   Fisheries Management Effectiveness       [Alabama, Florida, Louisiana, Mississippi]
#   Fisheries Management Effectiveness (Reef Fish) [Texas]
#   Fisheries Management Effectiveness (Sharks)    [Texas]
#   Fishing Sustainability                   [Florida]
#   Hook Sale Long Time Series Data          [Florida]
#   Rec Harvest Of Sandbars                  [Florida]
#   Shark Harvest                            [Alabama, Louisiana]
#   Shark Landings (Legal)                   [Florida]
#   Single Species Shark Stock Assessments   [Louisiana]
# =============================================================================

category_lookup <- tribble(
  ~Label_TC,                                                          ~Category,
  # ── Central Concept ─────────────────────────────────────────────────────────
  "Depredation",                                                      "Central Concept",
  "Shark Depredation",                                                "Central Concept",
  # ── Ecological & Biological Factors ─────────────────────────────────────────
  "Climate",                                                          "Ecological & Biological Factors",
  "Dolphin Depredation",                                              "Ecological & Biological Factors",
  "Dolphin Learning Behavior",                                        "Ecological & Biological Factors",
  "Ecosystem Health",                                                 "Ecological & Biological Factors",
  "Hooked Fish Behavior",                                    "Ecological & Biological Factors",
  "Forage Fish Populations",                                          "Ecological & Biological Factors",
  "Localized Red Snapper Depletion",                                  "Ecological & Biological Factors",
  "Many Small Sharks",                                                "Ecological & Biological Factors",
  "Muddy Water",                                                      "Ecological & Biological Factors",
  "Natural Bottom / Shelf",                                           "Ecological & Biological Factors",
  "Number Of Sharks And Dolphins In The Area",                        "Ecological & Biological Factors",
  "Prey Populations (Menhaden)",                                      "Ecological & Biological Factors",
  "Red Snapper",                                                      "Ecological & Biological Factors",
  "Reef Fish Populations",                                            "Ecological & Biological Factors",
  "Reef Size",                                                        "Ecological & Biological Factors",
  "Shark Aggregation",                                                "Ecological & Biological Factors",
  "Shark Learning Behavior",                                          "Ecological & Biological Factors",
  "Shark Pheromones/ Liver In Water",                                 "Ecological & Biological Factors",
  "Shark Populations",                                                "Ecological & Biological Factors",
  "Shark Populations On Reefs",                                       "Ecological & Biological Factors",
  "Shifting Stocks",                                                  "Ecological & Biological Factors",
  "Single Large Shark",                                               "Ecological & Biological Factors",
  "Target Species Prevalence At Fishing Location",                    "Ecological & Biological Factors",
  "Water Temperature",                                                "Ecological & Biological Factors",
  # ── Fisheries Operations & Practices ────────────────────────────────────────
  "Boat Signature",                                                   "Fisheries Operations & Practices",
  "Concentrated Artificial Reefs",                                    "Fisheries Operations & Practices",
  "Discarding Small Fish",                                            "Fisheries Operations & Practices",
  "Dolphin Deterrents",                                               "Fisheries Operations & Practices",
  "Fight Time",                                                       "Fisheries Operations & Practices",
  "Fishing Effort",                                 "Fisheries Operations & Practices",
  "Higher Fishing Pressure On Reefs",                                 "Fisheries Operations & Practices",
  "Hooked Fish Size",                                                 "Fisheries Operations & Practices",
  "Increased Reeling Time",                                           "Fisheries Operations & Practices",
  "Keeping Depredated Fish",                                          "Fisheries Operations & Practices",
  "Lights On Shrimping Boats",                                        "Fisheries Operations & Practices",
  "Number Of Fishers",                                              "Fisheries Operations & Practices",
  "Number Of Shrimp Boats",                                           "Fisheries Operations & Practices",
  "Time Of Day",                                                      "Fisheries Operations & Practices",
  "Fishing Time",                                                    "Fisheries Operations & Practices",
  "Widely Distributed Artificial Reefs",                              "Fisheries Operations & Practices",
  # ── Human Dimensions ────────────────────────────────────────────────────────
  "Fisher Competency",                                                "Human Dimensions",
  "Fisher Satisfaction",                                              "Human Dimensions",
  "Angler Stewardship",                                               "Human Dimensions",
  "Captain Satisfaction",                                             "Human Dimensions",
  "Captain Stress",                                                   "Human Dimensions",
  "Charter Captain Satisfaction",                                     "Human Dimensions",
  "Coastal Heritage",                                                 "Human Dimensions",
  "Commercial Captain Satisfaction",                                  "Human Dimensions",
  "Commercial Satisfaction",                                          "Human Dimensions",
  "Ability To Identify Shark Species",                                         "Human Dimensions",
  "Human Injury",                                                     "Human Dimensions",
  "Knowledge Of Depredation",                                         "Human Dimensions",
  "Miscommunication Between Fishermen + Scientists",                  "Human Dimensions",
  "Private Rec Angler Satisfaction",                                  "Human Dimensions",
  "Public Opinion Of Sharks + Dolphins",                              "Human Dimensions",
  "Safety",                                                           "Human Dimensions",
  "Shark Bites",                                                      "Human Dimensions",
  "Tourist Angler Satisfaction",                                      "Human Dimensions",
  # ── Policy & Economics ──────────────────────────────────────────────────────
  "Economic Loss",                                                    "Policy & Economics",
  "Gear Costs",                                               "Policy & Economics",
  "Ability To Commercially Harvest Sharks",                           "Policy & Economics",
  "Complex Species-Specific Regulations",                             "Policy & Economics",
  "Diverse Markets For Shark Products",                               "Policy & Economics",
  "Dolphin Regulatory Protection",                                    "Policy & Economics",
  "Economic Loss (Commercial)",                                       "Policy & Economics",
  "Economic Loss (Recreational)",                                     "Policy & Economics",
  "Federal Commercial Shark Industry",                                "Policy & Economics",
  "Federal Regulation",                                               "Policy & Economics",
  "Federal Water Shark Fishery (Rec)",                                "Policy & Economics",
  "Mirror State Limits For Shark Harvest In Fed Waters",              "Policy & Economics",
  "Red Snapper Size And Bag Limits",                                  "Policy & Economics",
  "Regulatory Consistency (Sharks)",                                  "Policy & Economics",
  "Season Length",                                                    "Policy & Economics",
  "Separation Between Gulf Council And Hms/Stock Assessment Process", "Policy & Economics",
  "Shark Messaging From Feds",                                        "Policy & Economics",
  "Shark Conservation",                                      "Policy & Economics",
  "Shark Revenue (From Selling Catch)",                               "Policy & Economics",
  "State Water Shark Fishery (Rec)",                                  "Policy & Economics",
  "Texas Fin Ban",                                                    "Policy & Economics",
  # ── Fisheries Research & Management ─────────────────────────────────────────
  "Data Gaps (Sandbar Stock Assessment Re-Evaluation)",               "Fisheries Research & Management",
  "Fisheries Management Effectiveness",                               "Fisheries Research & Management",
  "Fisheries Management Effectiveness (Reef Fish)",                   "Fisheries Research & Management",
  "Fisheries Management Effectiveness (Sharks)",                      "Fisheries Research & Management",
  "Fishing Sustainability",                                           "Fisheries Research & Management",
  "Hook Sale Long Time Series Data",                                  "Fisheries Research & Management",
  "Rec Harvest Of Sandbars",                                          "Fisheries Research & Management",
  "Shark Harvest",                                                    "Fisheries Research & Management",
  "Shark Landings (Legal)",                                           "Fisheries Research & Management",
  "Shark Fishery",                                                    "Fisheries Research & Management",
  "Single Species Shark Stock Assessments",                           "Fisheries Research & Management"
)

# =============================================================================
# STEP 2 — Load all CSV data
# =============================================================================
cat("========================================================\n")
cat("MMP → Kumu Converter  |  Alabama Workshop\n")
cat("========================================================\n\n")

all_data <- list()
for (s in states) {
  cat(sprintf("Loading: %s.csv\n", s))
  d <- read_mm_csv(input_files[[s]], s)
  cat(sprintf("  Concepts: %d  |  Connections: %d\n",
              length(d$concepts), nrow(d$connections)))
  
  # Report non-standard values  [A3]
  nonstandard <- d$connections %>%
    filter(!abs(MM_Weight) %in% c(0.5, 1.0))
  if (nrow(nonstandard) > 0) {
    cat(sprintf("  [A3] WARNING: %d non-standard MM weight(s) found in %s:\n",
                nrow(nonstandard), s))
    nonstandard %>%
      select(From_Label, To_Label, MM_Weight) %>%
      print()
    cat("       These will be binned to nearest ±0.5 / ±1.0 (threshold 0.75).\n")
  }
  
  # Add Kumu strength
  d$connections <- d$connections %>%
    mutate(
      Kumu_Strength    = mm_to_kumu(MM_Weight),
      `Influence Type` = if_else(Kumu_Strength > 0, "Positive", "Negative"),
      `Strength Label` = kumu_label(Kumu_Strength)
    )
  
  all_data[[s]] <- d
  cat("\n")
}

# =============================================================================
# STEP 3 — Build & write individual Kumu files
# =============================================================================
cat("--------------------------------------------------------\n")
cat("STEP 3: Writing individual Kumu files\n")
cat("--------------------------------------------------------\n\n")

for (s in states) {
  cat(sprintf("Processing: %s\n", s))
  d <- all_data[[s]]
  
  # ── Elements: one row per concept, Category from lookup [A5], Tag = state ──
  elements_df <- tibble(Label = d$concepts) %>%
    mutate(
      Label_TC = str_to_title(Label),
      Tags     = s                         # single-state tag [A11]
    ) %>%
    left_join(category_lookup, by = "Label_TC") %>%
    select(Label, Category, Tags) %>%
    arrange(Label)
  
  # Warn if any concept is missing from the lookup table
  n_missing_cat <- sum(is.na(elements_df$Category))
  if (n_missing_cat > 0) {
    cat(sprintf("  [A5] WARNING: %d concept(s) not found in category_lookup:\n",
                n_missing_cat))
    elements_df %>%
      filter(is.na(Category)) %>%
      pull(Label) %>%
      cat(sep = "\n         ")
    cat("\n")
  }
  
  # ── Connections: From, To, Strength, Influence Type, Strength Label,
  #                Original_MM_Weight, Tags  [A11] ───────────────────────────
  connections_df <- d$connections %>%
    rename(
      From               = From_Label,
      To                 = To_Label,
      Strength           = Kumu_Strength,
      Original_MM_Weight = MM_Weight
    ) %>%
    select(From, To, Strength, `Influence Type`, `Strength Label`,
           Original_MM_Weight, Tags = State) %>%   # Tags = single state [A11]
    arrange(From, To, Strength)
  
  # ── Notes ──────────────────────────────────────────────────────────────────
  cat_summary <- elements_df %>%
    count(Category, name = "n") %>%
    arrange(Category) %>%
    mutate(line = sprintf("%s (%d)", Category, n)) %>%
    pull(line) %>%
    paste(collapse = "; ")
  
  notes_df <- tibble(
    Item   = c("SOURCE", "STATE", "CONCEPTS", "CONNECTIONS",
               "STRENGTH SCALE", "", "", "",
               "NON-STANDARD VALUES", "CATEGORIES", "TAG FORMAT", "KUMU IMPORT"),
    Detail = c(
      "Mental Modeler adjacency-matrix CSV export",
      s,
      sprintf("%d unique concept labels", length(d$concepts)),
      sprintf("%d directed connections", nrow(connections_df)),
      "+2 = Strong Positive (MM +1.0)  |  +1 = Moderate Positive (MM +0.5)",
      "-1 = Moderate Negative (MM -0.5)  |  -2 = Strong Negative (MM -1.0)",
      "Original MM weights stored in 'Original_MM_Weight' column",
      "Conversion: |MM_weight| < 0.75 → moderate (±1); ≥ 0.75 → strong (±2)",
      "None — all weights are standard ±0.5 or ±1.0",
      cat_summary,
      paste0("Tags column = '", s, "' for all elements and connections [A11]"),
      "kumu.io → New Project → Import → Spreadsheet → upload this file"
    )
  )
  
  out_path <- file.path(output_dir, sprintf("Kumu_%s_Alabama_Workshop.xlsx", s))
  write_kumu_wb(elements_df, connections_df, notes_df, out_path, sheet_label = s)
}

# =============================================================================
# STEP 4 — Build Master Kumu (all states combined)
# =============================================================================
cat("\n--------------------------------------------------------\n")
cat("STEP 4: Building Master Kumu (all states)\n")
cat("--------------------------------------------------------\n\n")

# Combine all connections; normalise concept labels to Title Case  [A2]
all_conns <- bind_rows(lapply(states, function(s) all_data[[s]]$connections)) %>%
  mutate(
    From_norm = str_to_title(From_Label),
    To_norm   = str_to_title(To_Label)
  )

# Remove self-loops after normalisation (catches any cross-case matches)  [A8]
n_self_norm <- sum(all_conns$From_norm == all_conns$To_norm)
if (n_self_norm > 0) {
  cat(sprintf("[A8] Removed %d self-loop(s) after Title Case normalisation.\n\n",
              n_self_norm))
  all_conns <- all_conns %>% filter(From_norm != To_norm)
}

# ---- 4a. Master Elements  [A10] --------------------------------------------
# Build concept → states mapping (from connections, which cover all concepts
# since every MM concept appears as at least a From or To node)
concept_states <- bind_rows(
  all_conns %>% select(Concept = From_norm, State),
  all_conns %>% select(Concept = To_norm,   State)
) %>%
  distinct() %>%
  group_by(Concept) %>%
  summarise(Tags = paste(sort(unique(State)), collapse = "|"), .groups = "drop")

# Join categories from lookup [A5]; Label_TC = Concept (already Title Case)
master_elements <- concept_states %>%
  left_join(category_lookup, by = c("Concept" = "Label_TC")) %>%
  rename(Label = Concept) %>%
  select(Label, Category, Tags) %>%
  arrange(Label)

# Warn if any concept is missing from the lookup table
n_missing_master <- sum(is.na(master_elements$Category))
if (n_missing_master > 0) {
  cat(sprintf("[A5] WARNING: %d concept(s) in Master have no category assignment:\n",
              n_missing_master))
  master_elements %>%
    filter(is.na(Category)) %>%
    pull(Label) %>%
    cat(sep = "\n       ")
  cat("\n\n")
}

cat(sprintf("Master Elements: %d unique concepts\n", nrow(master_elements)))

# Report shared concepts (appear in >1 state)
shared <- master_elements %>% filter(str_detect(Tags, "\\|"))
cat(sprintf("  Shared across ≥2 states: %d\n", nrow(shared)))
cat(sprintf("  State-unique:            %d\n\n",
            nrow(master_elements) - nrow(shared)))

if (nrow(shared) > 0) {
  cat("  Shared concepts:\n")
  shared %>% select(Label, Tags) %>% print(n = Inf)
  cat("\n")
}

# ---- 4b. Master Connections  [A9] ------------------------------------------
master_conns_raw <- all_conns %>%
  select(From_norm, To_norm, Kumu_Strength, MM_Weight, State)

# Deduplicate: same (From, To, Strength) → merge; different Strength → keep separate
master_conns_dedup <- master_conns_raw %>%
  group_by(From_norm, To_norm, Kumu_Strength) %>%
  summarise(
    Tags                = paste(sort(unique(State)), collapse = "|"),  # [A11]
    Original_MM_Weights = paste(sort(unique(round(MM_Weight, 4))), collapse = "; "),
    n_sources           = n(),
    .groups             = "drop"
  )

# Report merged duplicates (same From, To, Strength across states)
dup_merged <- master_conns_dedup %>% filter(n_sources > 1)
if (nrow(dup_merged) > 0) {
  cat(sprintf("[A9] %d connection(s) appear in multiple states with SAME strength → merged:\n",
              nrow(dup_merged)))
  dup_merged %>%
    select(From_norm, To_norm, Kumu_Strength, Tags) %>%
    print(n = Inf)
  cat("\n")
} else {
  cat("[A9] No connections found with same (From, To, Strength) across states.\n\n")
}

# Report same (From, To) with DIFFERENT strengths → kept as separate rows
diff_strength <- master_conns_dedup %>%
  group_by(From_norm, To_norm) %>%
  filter(n() > 1) %>%
  ungroup()

if (nrow(diff_strength) > 0) {
  n_pairs <- nrow(diff_strength %>% distinct(From_norm, To_norm))
  cat(sprintf("[A9] %d concept pair(s) have DIFFERENT strengths across states → kept as separate rows:\n",
              n_pairs))
  diff_strength %>%
    select(From_norm, To_norm, Kumu_Strength, Tags) %>%
    print(n = Inf)
  cat("\n")
} else {
  cat("[A9] No concept pairs found with conflicting strengths across states.\n\n")
}

cat(sprintf("Master Connections total: %d rows (%d raw → %d after deduplication)\n\n",
            nrow(master_conns_dedup), nrow(master_conns_raw),
            nrow(master_conns_dedup)))

# Build final connections data frame
master_connections <- master_conns_dedup %>%
  mutate(
    `Influence Type` = if_else(Kumu_Strength > 0, "Positive", "Negative"),
    `Strength Label` = kumu_label(Kumu_Strength)
  ) %>%
  rename(
    From               = From_norm,
    To                 = To_norm,
    Strength           = Kumu_Strength,
    Original_MM_Weight = Original_MM_Weights
  ) %>%
  select(From, To, Strength, `Influence Type`, `Strength Label`,
         Original_MM_Weight, Tags) %>%
  arrange(From, To, Strength)

# ---- 4c. Master Notes -------------------------------------------------------
n_pos <- sum(master_connections$Strength > 0)
n_neg <- sum(master_connections$Strength < 0)

master_cat_summary <- master_elements %>%
  count(Category, name = "n") %>%
  arrange(Category) %>%
  mutate(line = sprintf("%s (%d)", Category, n)) %>%
  pull(line) %>%
  paste(collapse = "; ")

notes_master <- tibble(
  Item   = c(
    "WORKSHOP", "STATES", "", "",
    "MASTER ELEMENTS", "MASTER CONNECTIONS", "  of which positive", "  of which negative",
    "SHARED CONCEPTS", "STATE-UNIQUE CONCEPTS",
    "CONCEPT NORMALISATION",
    "NON-STANDARD VALUES",
    "STRENGTH SCALE", "", "", "",
    "DEDUPLICATION (CONNECTIONS)", "", "",
    "CATEGORIES",
    "TAG FORMAT", "KUMU IMPORT"
  ),
  Detail = c(
    "Alabama Workshop (single workshop; participants grouped by home state)",
    paste(states, collapse = ", "),
    "Each state group represents within-workshop subgroup, NOT a separate workshop",
    "Region tags follow same convention as Australia workflow",
    sprintf("%d unique concept labels (Title Case normalised)", nrow(master_elements)),
    sprintf("%d unique connection rows", nrow(master_connections)),
    sprintf("%d", n_pos),
    sprintf("%d", n_neg),
    sprintf("%d concepts appear in ≥2 states", nrow(shared)),
    sprintf("%d concepts appear in only 1 state",
            nrow(master_elements) - nrow(shared)),
    "All concept labels use consistent Title Case; str_to_title applied as safety net [A2]",
    "All weights are standard ±0.5 or ±1.0; no non-standard values detected [A3]",
    "+2 = Strong Positive (MM +1.0)",
    "+1 = Moderate Positive (MM +0.5)",
    "-1 = Moderate Negative (MM -0.5)",
    "-2 = Strong Negative (MM -1.0)",
    "Unique key = (From Label, To Label, Kumu Strength) [A9]",
    "Same key across states → ONE row; Tags = pipe-joined states",
    "Same (From, To) with DIFFERENT strengths → SEPARATE rows (preserves diversity)",
    master_cat_summary,
    "Tags use | (pipe, no spaces): e.g. 'Alabama|Florida' [A11]",
    "kumu.io → New Project → Import → Spreadsheet → upload this file"
  )
)

# ---- 4d. Write Master Kumu workbook ----------------------------------------
master_path <- file.path(output_dir, "Kumu_MASTER_Alabama_Workshop.xlsx")
write_kumu_wb(master_elements, master_connections, notes_master,
              master_path, sheet_label = "All States")

# =============================================================================
# STEP 5 — Summary
# =============================================================================
cat("\n========================================================\n")
cat("DONE\n")
cat("========================================================\n\n")
cat("Individual Kumu files written:\n")
for (s in states) {
  cat(sprintf("  Kumu_%s_Alabama_Workshop.xlsx\n", s))
}
cat(sprintf("\nMaster Kumu file:\n  Kumu_MASTER_Alabama_Workshop.xlsx\n\n"))

cat("Category breakdown (Master):\n")
master_elements %>%
  count(Category, name = "n") %>%
  arrange(Category) %>%
  print(n = Inf)

cat("\nShared concepts (appear in ≥2 states):\n")
master_elements %>%
  filter(str_detect(Tags, "\\|")) %>%
  select(Label, Category, Tags) %>%
  arrange(Category, Label) %>%
  print(n = Inf)

cat("\nREMINDERS:\n")
cat("[A2] All labels use consistent Title Case; str_to_title applied as safety net in Master.\n")
cat("[A3] All weights are standard ±0.5 or ±1.0. No non-standard values detected.\n")
cat("[A5] Categories assigned via hardcoded lookup table. Any NA values printed above\n")
cat("     indicate concepts not yet in the lookup — add them manually.\n")
cat("[A7] 'Shark Depredation' and 'Depredation' are kept separate (Florida).\n")
cat("[A9] Check diff_strength report above for any cross-state strength conflicts.\n")
cat("[A11] Tags in individual files = single state name.\n")
cat("      Tags in Master = pipe-joined state list (e.g. 'Alabama|Florida').\n\n")

