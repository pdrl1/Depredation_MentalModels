# =============================================================
# MentalModeler_to_Kumu_Galveston.R
# Galveston Workshop — Three fisher groups:
#   Recreational  (Recreational_FinalModel.csv)
#   Charter       (TX_Charter_FinalModel.csv)
#   Commercial    (TX_Commercial_FinalModel.csv)
#
# Output: Kumu_MASTER_Galveston_Workshop.xlsx
# =============================================================
# ASSUMPTIONS
#  A1. Strength scale conversion:
#        MM ±0.5  → Kumu ±1 (Weak)
#        MM ±1.0  → Kumu ±2 (Strong)
#        Non-standard weights: |w| > 0.75 → ±2, else → ±1
#  A2. Rows/columns whose label begins with "Note" are researcher
#      annotations and are excluded from all processing.
#  A3. Connections with MM weight = 0 are excluded (no relationship).
#  A4. All labels are Title-Cased for consistent cross-group matching.
#  A5. Deduplication key: (From, To, Kumu Strength).
#      If the same (From, To) pair has different Kumu strengths across
#      groups, both records are retained and flagged in the console.
#  A6. Connection tags reflect the group(s) in which each connection
#      appears; multiple groups are pipe-separated (alphabetical order).
#  A7. Element tags reflect the group(s) in which each concept appears;
#      multiple groups are pipe-separated (alphabetical order).
#  A8. "Public Perceptions" (Recreational) and "Public Perception"
#      (Charter) differ by one letter after Title Case — kept separate.
#      Verify with participants whether they represent the same concept.
#  A9. The en-dash in "Shark Vessel–Food Association..." is Unicode
#      U+2013 and is preserved as-is from the Charter CSV.
# =============================================================

library(tidyverse)
library(openxlsx)


# ── 0. CONFIG ────────────────────────────────────────────────
input_files <- list(
  Recreational = "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/MM/Recreational_FinalModel.csv",
  Charter      = "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/MM/Charter_FinalModel.csv",
  Commercial   = "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/MM/Commercial_FinalModel.csv"
)
output_file <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Kumu_MASTER_Galveston_Workshop.xlsx"

# ── 1. CATEGORY LOOKUP ───────────────────────────────────────
# 72 unique Title-Cased concept labels → 6 thematic categories
category_lookup <- tribble(
  ~Label_TC,                                                                                                                ~Category,

  # ── Central Concept ──────────────────────────────────────────────────────────────────────────────────────────────────
  "Shark Depredation",                                                                                                      "Central Concept",

  # ── Ecological & Biological Factors ──────────────────────────────────────────────────────────────────────────────────
  "Shark Populations",                                                                                                      "Ecological & Biological Factors",
  "Dolphin Populations",                                                                                                    "Ecological & Biological Factors",
  "Shifting Distributions Of Shark Populations",                                                                            "Ecological & Biological Factors",
  "Shark Learning Behavior",                                                                                                "Ecological & Biological Factors",
  "Hooked Fish Behavior",                                                                                                   "Ecological & Biological Factors",
  "Shark Vessel\u2013Food Association (Attraction And Following Fishing Boats Due To Learned Food Rewards)",                "Ecological & Biological Factors",
  "Shark Attraction To Vessels",                                                                                            "Ecological & Biological Factors",
  "Climate Change",                                                                                                         "Ecological & Biological Factors",
  "Water Temperature",                                                                                                      "Ecological & Biological Factors",
  "Bp Oil Spill Reduced Prey Populations (Shark's Food)",                                                                   "Ecological & Biological Factors",
  "Habitat Loss (Removal of rigs and platforms)",                                                                           "Ecological & Biological Factors",
  "Increased Number Of Artificial Reefs",                                                                                   "Ecological & Biological Factors",

  # ── Fisheries Operations & Practices ─────────────────────────────────────────────────────────────────────────────────
  "Damage To Gear",                                                                                                         "Fisheries Operations & Practices",
  "Damage To Vessels By The Sharks Frenzy",                                                                                 "Fisheries Operations & Practices",
  "Bycatch Reduction Devices (Shrimp Fishery)",                                                                             "Fisheries Operations & Practices",
  "Switch From Bait To Lures/Jigs",                                                                                         "Fisheries Operations & Practices",
  "Catch And Release Trips",                                                                                                "Fisheries Operations & Practices",
  "Change Fishing Locations",                                                                                               "Fisheries Operations & Practices",
  "Effort In Cleaning And Processing Sharks",                                                                               "Fisheries Operations & Practices",
  "Insurance Risks For Charters",                                                                                           "Fisheries Operations & Practices",
  "More Frequent Fishing Spots (Hotspots)",                                                                                 "Fisheries Operations & Practices",
  "Catch In Commercial Fishing",                                                                                            "Fisheries Operations & Practices",
  "Number Of Shrimp Boats",                                                                                                 "Fisheries Operations & Practices",
  "Fishing Time",                                                                                                           "Fisheries Operations & Practices",
  "Recreational Fishing Boats Following Commercial Fishing Vessels. This Attracts Sharks To Shrimper Boats",                "Fisheries Operations & Practices",
  "Changing Species Targeted By Recreational Fishing",                                                                      "Fisheries Operations & Practices",

  # ── Human Dimensions ─────────────────────────────────────────────────────────────────────────────────────────────────
  "Fishermen's Income",                                                                                                     "Human Dimensions",
  "Public Fear - Dangling Your Feet",                                                                                       "Human Dimensions",
  "Public Positive Perception",                                                                                             "Human Dimensions",
  "Health Conscious",                                                                                                       "Human Dimensions",
  "Increased Fishing Interest In Social Media",                                                                             "Human Dimensions",
  "More Recreational Fishermen",                                                                                            "Human Dimensions",
  "Interest By Charter Clients",                                                                                            "Human Dimensions",
  "Captain Satisfaction",                                                                                                   "Human Dimensions",
  "Client Satisfaction",                                                                                                    "Human Dimensions",
  "Fisher Satisfaction With Specific Species Fishing",                                                                      "Human Dimensions",
  "Participation In Tournaments - Shark Mutilate Catch Ex: Bigrock In Nc",                                                  "Human Dimensions",
  "Knowledge Of Depredation",                                                                                               "Human Dimensions",
  "People's Desire To Save Sharks",                                                                                         "Human Dimensions",
  "Demand For Sharks",                                                                                                      "Human Dimensions",
  "Ability To Identify Shark Species",                                                                                      "Human Dimensions",
  "Crew Safety",                                                                                                            "Human Dimensions",


  # ── Policy & Economics ───────────────────────────────────────────────────────────────────────────────────────────────
  "Legislation",                                                                                                            "Policy & Economics",
  "Government Enforcement And Federal Regulations",                                                                         "Policy & Economics",
  "Removal Of Rigs And Platforms",                                                                                          "Policy & Economics",
  "Shark Finning Law - Illegal To Separate Fins From Body",                                                                 "Policy & Economics",
  "Shark Conservation",                                                                                                     "Policy & Economics",
  "Increase Shark Quotas",                                                                                                  "Policy & Economics",
  "Lack Of Economic Incentives To Overcome The Impacts Of Depredation",                                                     "Policy & Economics",
  "Shark Fin Market",                                                                                                       "Policy & Economics",
  "Gear Costs",                                                                                                             "Policy & Economics",
  "Cost Of Fines From Gear Damage (Shrimpers)",                                                                             "Policy & Economics",
  "Labor Time To Repair Gear",                                                                                              "Policy & Economics",
  "Economic Loss",                                                                                                          "Policy & Economics",
  "Fishing Profitability",                                                                                                  "Policy & Economics",

  
  # ── Fisheries Research & Management ──────────────────────────────────────────────────────────────────────────────────
  "Research And Funding",                                                                                                   "Fisheries Research & Management",
  "Fisheries Management Effectiveness (Reef Fish)",                                                                           "Fisheries Research & Management",
  "Hms Management Actions",                                                                                                 "Fisheries Research & Management",
  "Hms Slow Management Response",                                                                                           "Fisheries Research & Management",
  "Shark Fishery",                                                                                                         "Fisheries Research & Management",
  "Commercial Shark Fisheries",                                                                                             "Fisheries Research & Management",
  "Efficiency Of Commercial Fishery",                                                                                       "Fisheries Research & Management"
  
)

# ── 2. HELPER: parse a Mental Modeler CSV ────────────────────
parse_mm_csv <- function(filepath, group_tag) {
  raw <- read.csv(filepath, header = FALSE, stringsAsFactors = FALSE,
                  check.names = FALSE)

  # Row 1 = column headers (skip first cell); Col 1 = row labels (skip header)
  col_labels_raw <- as.character(raw[1, -1])
  row_labels_raw <- as.character(raw[-1,  1])
  mat            <- raw[-1, -1]

  # Title-case and trim whitespace
  tc <- function(x) str_to_title(trimws(x))
  col_tc <- tc(col_labels_raw)
  row_tc <- tc(row_labels_raw)

  # ── Exclude researcher annotation rows/columns (start with "Note") ──
  keep_col <- !grepl("^Note", col_labels_raw, ignore.case = FALSE)
  keep_row <- !grepl("^Note", row_labels_raw, ignore.case = FALSE)
  col_tc   <- col_tc[keep_col]
  row_tc   <- row_tc[keep_row]
  mat      <- mat[keep_row, keep_col, drop = FALSE]

  # ── Extract non-zero connections ────────────────────────────
  conns <- tibble()
  for (i in seq_len(nrow(mat))) {
    for (j in seq_len(ncol(mat))) {
      val_str <- trimws(as.character(mat[i, j]))
      if (val_str == "" || is.na(val_str)) next
      val <- suppressWarnings(as.numeric(val_str))
      if (is.na(val) || val == 0) next          # skip blank and zero-weight
      conns <- bind_rows(conns, tibble(
        From      = row_tc[i],
        To        = col_tc[j],
        MM_Weight = val,
        Group     = group_tag
      ))
    }
  }

  list(labels = unique(c(col_tc, row_tc)), connections = conns)
}

# ── 3. HELPER: MM weight → Kumu strength ─────────────────────
# Standard: ±0.5 → ±1 | ±1.0 → ±2
# Non-standard: |w| > 0.75 → ±2, else → ±1
mm_to_kumu <- function(w) {
  as.integer(ifelse(abs(w) > 0.75, sign(w) * 2L, sign(w) * 1L))
}

# ── 4. PARSE ALL THREE CSVs ──────────────────────────────────
parsed <- imap(input_files, parse_mm_csv)

# ── 5. COMBINE ALL LABELS ────────────────────────────────────
all_labels <- unique(unlist(map(parsed, "labels")))

# ── 6. COMBINE ALL CONNECTIONS ───────────────────────────────
all_conns_raw <- bind_rows(map(parsed, "connections")) %>%
  mutate(Kumu_Strength = mm_to_kumu(MM_Weight))

# ── 7. DETECT STRENGTH CONFLICTS ACROSS GROUPS ───────────────
conflict_pairs <- all_conns_raw %>%
  group_by(From, To) %>%
  summarise(n_strengths = n_distinct(Kumu_Strength), .groups = "drop") %>%
  filter(n_strengths > 1)

if (nrow(conflict_pairs) > 0) {
  message("\n⚠  CONFLICTING STRENGTHS detected (verify with participants):")
  all_conns_raw %>%
    semi_join(conflict_pairs, by = c("From", "To")) %>%
    arrange(From, To, Group) %>%
    select(From, To, Group, MM_Weight, Kumu_Strength) %>%
    print(n = Inf)
} else {
  message("✅ No conflicting strengths detected across groups.")
}

# ── 8. BUILD ELEMENTS SHEET ──────────────────────────────────
# For each element, record which group(s) contain it
element_groups <- map_dfr(names(parsed), function(g) {
  tibble(Label = parsed[[g]]$labels, Group = g)
}) %>%
  group_by(Label) %>%
  summarise(Tags = paste(sort(unique(Group)), collapse = "|"), .groups = "drop")

elements_df <- tibble(Label = all_labels) %>%
  left_join(category_lookup, by = c("Label" = "Label_TC")) %>%
  left_join(element_groups,  by = "Label") %>%
  arrange(Label)

# Warn about any concepts not in the category lookup
missing_cats <- elements_df %>% filter(is.na(Category))
if (nrow(missing_cats) > 0) {
  warning(
    "⚠  Missing category for ", nrow(missing_cats), " element(s):\n  ",
    paste(missing_cats$Label, collapse = "\n  "),
    "\n  Add these to category_lookup and re-run."
  )
}

# ── 9. BUILD CONNECTIONS SHEET ───────────────────────────────
# Deduplicate by (From, To, Kumu_Strength); merge group tags
connections_df <- all_conns_raw %>%
  group_by(From, To, Kumu_Strength) %>%
  summarise(
    Original_MM_Weight = paste(sort(unique(as.character(MM_Weight))), collapse = "|"),
    Tags               = paste(sort(unique(Group)), collapse = "|"),
    .groups            = "drop"
  ) %>%
  mutate(
    `Influence Type` = ifelse(Kumu_Strength > 0, "Positive", "Negative"),
    `Strength Label` = case_when(
      Kumu_Strength ==  2 ~ "Strong Positive",
      Kumu_Strength ==  1 ~ "Moderate Positive",
      Kumu_Strength == -1 ~ "Moderate Negative",
      Kumu_Strength == -2 ~ "Strong Negative",
      TRUE                ~ "Unknown"
    )
  ) %>%
  select(
    From, To,
    Strength         = Kumu_Strength,
    `Influence Type`,
    `Strength Label`,
    Original_MM_Weight,
    Tags
  ) %>%
  arrange(From, To)

# ── 10. BUILD NOTES & ASSUMPTIONS SHEET ──────────────────────
n_elements    <- nrow(elements_df)
n_raw_conns   <- nrow(all_conns_raw)
n_unique_conns <- nrow(connections_df)
n_conflicts   <- nrow(conflict_pairs)

notes_df <- tibble(
  Item = c(
    "Workshop",
    "Groups (tags)",
    "Input files",
    "Script",
    "Date generated",
    "─── COUNTS ───",
    "Unique elements",
    "Raw connections (all groups combined)",
    "Unique connections (deduplicated)",
    "Conflicting strength pairs",
    "─── ASSUMPTIONS ───",
    "A1 — Strength conversion",
    "A2 — Note rows excluded",
    "A3 — Zero-weight connections",
    "A4 — Title Case normalization",
    "A5 — Deduplication key",
    "A6 — Connection tags",
    "A7 — Element tags",
    "A8 — Public Perceptions vs Public Perception",
    "A9 — En-dash in Shark Vessel-Food Association",
    "─── REVIEW FLAGS ───",
    "Conflicting strengths",
    "Missing categories",
    "Concept label similarity"
  ),
  Detail = c(
    "Galveston Workshop",
    "Recreational | Charter | Commercial",
    paste(unname(unlist(input_files)), collapse = " | "),
    "MentalModeler_to_Kumu_Galveston.R",
    as.character(Sys.Date()),
    "",
    as.character(n_elements),
    as.character(n_raw_conns),
    as.character(n_unique_conns),
    if (n_conflicts > 0) paste(n_conflicts, "pair(s) — see Connections sheet for detail") else "None detected",
    "",
    "MM \u00b10.5 \u2192 Kumu \u00b11 | MM \u00b11.0 \u2192 Kumu \u00b12 | non-standard |w|>0.75 \u2192 \u00b12, else \u00b11",
    "Rows/columns beginning with 'Note' excluded (researcher annotations)",
    "Connections with MM weight = 0 excluded",
    "All labels converted to Title Case for cross-group matching",
    "(From, To, Kumu Strength) — same (From, To) with different strengths retained separately",
    "Each connection tagged with group(s) it appears in; pipe-separated alphabetically",
    "Each element tagged with group(s) it appears in; pipe-separated alphabetically",
    "'Public Perceptions' (Recreational) vs 'Public Perception' (Charter) — verify if same concept",
    "Unicode U+2013 en-dash preserved from Charter CSV label",
    "",
    if (n_conflicts > 0) paste(n_conflicts, "pair(s) — verify strength direction with participants") else "None",
    if (nrow(missing_cats) > 0) paste(missing_cats$Label, collapse = "; ") else "None",
    "'Public Perceptions' (Recreational) vs 'Public Perception' (Charter) may be the same concept"
  )
)

# ── 11. WRITE EXCEL ──────────────────────────────────────────
wb <- createWorkbook()

# ── Style helpers ──
make_header_style <- function(fill = "#2C4770") {
  createStyle(fontColour = "#FFFFFF", fgFill = fill,
              halign = "LEFT", textDecoration = "Bold",
              border = "Bottom", borderColour = "#FFFFFF",
              fontSize = 11)
}
sep_style  <- createStyle(fgFill = "#D9D9D9", textDecoration = "Bold")
flag_style <- createStyle(fgFill = "#FFE699")   # yellow highlight for warnings

# ── Elements sheet ──
addWorksheet(wb, "Elements")
writeData(wb, "Elements", elements_df)
addStyle(wb, "Elements", make_header_style("#2C4770"),
         rows = 1, cols = 1:ncol(elements_df), gridExpand = TRUE)
setColWidths(wb, "Elements", cols = 1:ncol(elements_df), widths = "auto")

# Highlight elements with missing category
if (nrow(missing_cats) > 0) {
  miss_rows <- which(elements_df$Label %in% missing_cats$Label) + 1  # +1 for header
  addStyle(wb, "Elements", flag_style,
           rows = miss_rows, cols = 1:ncol(elements_df), gridExpand = TRUE)
}

# ── Connections sheet ──
addWorksheet(wb, "Connections")
writeData(wb, "Connections", connections_df)
addStyle(wb, "Connections", make_header_style("#2C4770"),
         rows = 1, cols = 1:ncol(connections_df), gridExpand = TRUE)
setColWidths(wb, "Connections", cols = 1:ncol(connections_df), widths = "auto")

# Highlight conflicting strength rows (rows where same From-To has multiple records)
if (nrow(conflict_pairs) > 0) {
  conflict_rows <- which(
    paste(connections_df$From, connections_df$To) %in%
    paste(conflict_pairs$From, conflict_pairs$To)
  ) + 1
  addStyle(wb, "Connections", flag_style,
           rows = conflict_rows, cols = 1:ncol(connections_df), gridExpand = TRUE)
}

# ── Notes & Assumptions sheet ──
addWorksheet(wb, "Notes & Assumptions")
writeData(wb, "Notes & Assumptions", notes_df)
addStyle(wb, "Notes & Assumptions", make_header_style("#4A7C59"),
         rows = 1, cols = 1:2, gridExpand = TRUE)
setColWidths(wb, "Notes & Assumptions", cols = 1:2, widths = c(40, 80))

# Highlight separator rows (items starting with "─")
sep_rows <- which(grepl("^\u2500", notes_df$Item)) + 1
if (length(sep_rows) > 0) {
  addStyle(wb, "Notes & Assumptions", sep_style,
           rows = sep_rows, cols = 1:2, gridExpand = TRUE)
}

saveWorkbook(wb, output_file, overwrite = TRUE)

# ── 12. CONSOLE SUMMARY ──────────────────────────────────────
cat("\n")
cat("============================================================\n")
cat("  Kumu Master — Galveston Workshop\n")
cat("============================================================\n")
cat("  Elements:              ", n_elements,     "\n")
cat("  Raw connections:       ", n_raw_conns,    "\n")
cat("  Unique connections:    ", n_unique_conns, "\n")
cat("  Strength conflicts:    ", n_conflicts,    "\n")
cat("  Missing categories:    ", nrow(missing_cats), "\n")
cat("  Output file:           ", output_file,   "\n")
cat("============================================================\n")
if (nrow(missing_cats) > 0) {
  cat("⚠  Add these to category_lookup and re-run:\n")
  cat("  ", paste(missing_cats$Label, collapse = "\n   "), "\n")
}


#Note: Manually added into the xlsx because Mental Modeler does not understand: 
# -->  CLient Satifaction to Depredation +1 and -1