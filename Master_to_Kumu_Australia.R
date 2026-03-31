# =============================================================================
# Master_to_Kumu.R
# Converts Master.xlsx (Australian shark depredation mental models)
# into a Kumu-ready Excel file (Elements + Connections sheets)
#
# Author: generated with Claude (Anthropic) - March 2026
# Usage:  Set input_file and output_file paths below, then source() or Rscript
# =============================================================================

# Required packages
library(readxl)
library(dplyr)
library(tidyr)
library(stringr)
library(openxlsx)


# =============================================================================
# ASSUMPTIONS
# =============================================================================
#
# A1. SIGN CONVENTION
#     The Master "Elements" sheet has a "Type" column encoding relationship
#     direction:
#       Type = "Yes"  ->  POSITIVE (+): concept INCREASES the target
#       Type = "No"   ->  NEGATIVE (-): concept DECREASES the target
#     Verified against known examples:
#       Deterrents -> Depredation        : always "No"  (negative) OK
#       Shark Habituation -> Depredation : always "Yes" (positive) OK
#
# A2. STRENGTH SCALE
#     Original values: 1 = moderate, 2 = strong.
#     After applying sign (A1): -2, -1, +1, +2.
#     These raw signed values are kept in the Kumu output.
#     Mental Modeler / MMP files use a different normalised scale (divide by 2).
#
# A3. REGION ASSIGNMENT
#     Each connection is assigned to the region of its FROM element.
#     If FROM = "Depredation" (central node), the TO element's region is used.
#     All inter-element connections were verified to be within the same workshop
#     -- no cross-workshop connections exist in the dataset.
#     Region groupings (per project team specification):
#       Western Australia    : Perth, Geraldton, Exmouth, Broome
#       North Australia      : Darwin
#       Queensland           : Townsville, Brisbane
#       New South Wales : Coffs Harbour
#     NOTE: The project team originally wrote "South West Australia" for
#     Coffs Harbour, but that location is in New South Wales (east coast).
#     This script uses "New South Wales". Edit region_map below to revert.
#
# A4. ELEMENT DEDUPLICATION
#     The Master Elements sheet has 297 data rows but only 70 unique Label
#     names. The same label (e.g. "Shark Abundance") appears multiple times
#     because it was mentioned by different workshops/groups.
#     Kumu gets ONE row per unique label; Tags list every region that used it.
#     "Depredation" is added as a 71st element (central node -- it only appears
#     in the From/To columns of Connections, never as its own Label row).
#
# A5. CONNECTION DEDUPLICATION
#     A connection is uniquely identified by (From Label, To Label, Signed Strength).
#     Rules:
#       Same triple in multiple regions  -> ONE row; regions joined with |
#       Same (From, To) with DIFFERENT strengths -> SEPARATE rows kept
#         (preserves diversity of stakeholder perceptions)
#     Result: 455 raw rows -> 298 unique rows.
#
# A6. TAG FORMAT
#     Tags use | (pipe, NO surrounding spaces) as the multi-value separator.
#     Example: "North Australia|Western Australia"
#     Kumu splits on | correctly. Using "; " (semicolon + space) causes a
#     leading-space bug where "Western Australia" != " Western Australia".
#
# A7. CATEGORY DEDUPLICATION (the 95-row bug fix)
#     21 of the 70 unique Labels have different Category values assigned by
#     different workshops. Using group_by(Label, Category) would produce 95
#     rows because each (Label, Category) combination becomes its own row.
#     Fix: a category_corrections lookup assigns ONE canonical Category per
#     label, applied BEFORE grouping; then group_by(Label) alone is used.
#     Resolution strategy: majority vote (>=60%) or domain logic for ties.
#     All 21 decisions are documented in the category_corrections block below.
# =============================================================================


# ---- File paths -- edit these -----------------------------------------------

input_file  <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/Mental Models from Marcus/Australia/Master.xlsx"

output_file <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Australia/Kumu/Kumu_Australia_SharkDepredation_R.xlsx"


# ---- Load data --------------------------------------------------------------
elements_raw <- read_excel(input_file, sheet = "Elements")
conns_raw    <- read_excel(input_file, sheet = "Connections")

# Clean column names and Type column (critical for A1)
names(elements_raw) <- str_trim(names(elements_raw))
names(conns_raw)    <- str_trim(names(conns_raw))
elements_raw$Type   <- str_trim(as.character(elements_raw$Type))
conns_raw$Type      <- str_trim(as.character(conns_raw$Type))


# ---- Region mapping (A3) ----------------------------------------------------
region_map <- c(
  "Perth"         = "Western Australia",
  "Geraldton"     = "Western Australia",
  "Exmouth"       = "Western Australia",
  "Broome"        = "Western Australia",
  "Darwin"        = "North Australia",
  "Townsville"    = "Queensland",
  "Brisbane"      = "Queensland",
  "Coffs Harbour" = "New South Wales"
)

# Named lookup vectors: key = element ID as character string
# NOTE: Excel reads numeric IDs as doubles (e.g. 1.0).
#       as.integer() strips the decimal -> "1", "4", "6" etc.
id_to_label  <- setNames(elements_raw$Label,
                         as.character(as.integer(elements_raw$ID)))
id_to_region <- setNames(region_map[elements_raw$Workshop],
                         as.character(as.integer(elements_raw$ID)))


# ---- Helper: resolve a From/To value to its Label string --------------------
# Input: a single value -- either a numeric ID (possibly 1.0) or "Depredation"
resolve_label <- function(val) {
  s <- str_trim(as.character(val))
  s <- sub("\\.0+$", "", s)          # strip trailing .0 from numeric IDs
  if (s == "Depredation") return("Depredation")
  label <- id_to_label[s]
  if (!is.na(label)) return(unname(label))
  return(s)                           # fallback: return raw value
}

# ---- Helper: resolve the region for a connection (A3) -----------------------
resolve_region <- function(from_val, to_val) {
  sf <- str_trim(sub("\\.0+$", "", as.character(from_val)))
  if (sf != "Depredation") {
    r <- id_to_region[sf]
    if (!is.na(r)) return(unname(r))
  }
  st <- str_trim(sub("\\.0+$", "", as.character(to_val)))
  if (st != "Depredation") {
    r <- id_to_region[st]
    if (!is.na(r)) return(unname(r))
  }
  return(NA_character_)
}


# ---- Category corrections (A7) ----------------------------------------------
#
# A7. CATEGORY INCONSISTENCY
#     21 Labels have different Category values assigned across workshops.
#     Without correction, group_by(Label, Category) creates duplicate rows
#     (95 rows instead of 71).
#     Fix: apply a single canonical Category per Label before grouping,
#     then group_by(Label) only.
#
#     Resolution rules:
#       - Clear majority (>=60 %): use most-common value.
#       - Exact tie: use the category that best fits the concept domain
#         (documented per label below).
#
#     Labels corrected and their chosen category:
#
#       Best Handling/Release Practices (Fish) : "Fisheries Operations & Practices"
#         (3-way tie 1/1/1 -- fishing handling practices are operational)
#       Charter/Rec Fisheries (Fish)           : "Fisheries Operations & Practices" (2/3 majority)
#       Ecosystem Imbalance                    : "Ecological & Biological Factors"  (5/6 majority)
#       Environmental Organization Influence   : "Policy & Economics"               (5/7 majority)
#       Fish Discards                          : "Fisheries Operations & Practices" (6/7 majority)
#       Fisher Expertise                       : "Human Dimensions"
#         (2/4 tie -- expertise is a personal/human skill factor)
#       Fisheries Health                       : "Fisheries Research & Management"
#         (3-way tie 1/1/1 -- health of fish stocks = research/management domain)
#       Fishing Effort Concentration           : "Fisheries Operations & Practices" (7/8 majority)
#       Mismanagement                          : "Policy & Economics"
#         (4/4 tie -- failure of governance/policy rather than research)
#       Number of Fishers                      : "Fisheries Operations & Practices" (9/10 majority)
#       Political Correctness                  : "Human Dimensions"                 (4/6 majority)
#       Proximity to Structure                 : "Ecological & Biological Factors"
#         (2/2 tie -- spatial/habitat feature, ecological factor)
#       Public Education/Marketing             : "Human Dimensions"                 (5/8 majority)
#       Seasonal/Area Closures                 : "Fisheries Research & Management"  (5/6 majority)
#       Shark Abundance                        : "Ecological & Biological Factors"  (8/9 majority)
#       Shark Harvest                          : "Fisheries Research & Management"  (6/7 majority)
#       Shark Harvest (Commercial)             : "Fisheries Research & Management"  (7/8 majority)
#       Shark Market Demand                    : "Policy & Economics"               (4/5 majority)
#       Shark Protections                      : "Policy & Economics"
#         (1/1 tie -- shark protection regulations are policy instruments)
#       Target Species Depth                   : "Ecological & Biological Factors"  (3/5 majority)
#       Tourism                                : "Policy & Economics"               (4/5 majority)
#
category_corrections <- c(
  "Best Handling/Release Practices (Fish)" = "Fisheries Operations & Practices",
  "Charter/Rec Fisheries (Fish)"           = "Fisheries Operations & Practices",
  "Ecosystem Imbalance"                    = "Ecological & Biological Factors",
  "Environmental Organization Influence"   = "Policy & Economics",
  "Fish Discards"                          = "Fisheries Operations & Practices",
  "Fisher Expertise"                       = "Human Dimensions",
  "Fisheries Health"                       = "Fisheries Research & Management",
  "Fishing Effort Concentration"           = "Fisheries Operations & Practices",
  "Mismanagement"                          = "Policy & Economics",
  "Number of Fishers"                      = "Fisheries Operations & Practices",
  "Political Correctness"                  = "Human Dimensions",
  "Proximity to Structure"                 = "Ecological & Biological Factors",
  "Public Education/Marketing"             = "Human Dimensions",
  "Seasonal/Area Closures"                 = "Fisheries Research & Management",
  "Shark Abundance"                        = "Ecological & Biological Factors",
  "Shark Harvest"                          = "Fisheries Research & Management",
  "Shark Harvest (Commercial)"             = "Fisheries Research & Management",
  "Shark Market Demand"                    = "Policy & Economics",
  "Shark Protections"                      = "Policy & Economics",
  "Target Species Depth"                   = "Ecological & Biological Factors",
  "Tourism"                                = "Policy & Economics"
)


# ---- Build Elements (A4) ----------------------------------------------------
#
# Step 1: apply canonical category where a correction exists
# Step 2: group by Label only (not Category) -> exactly one row per unique label
#
elements_df <- elements_raw %>%
  mutate(
    Region   = region_map[Workshop],
    Category = if_else(Label %in% names(category_corrections),
                       category_corrections[Label],
                       Category)
  ) %>%
  group_by(Label) %>%
  summarise(
    Category = first(Category),                            # A7: one canonical value
    Tags     = paste(sort(unique(Region)), collapse = "|"), # A6: pipe, no spaces
    .groups  = "drop"
  ) %>%
  arrange(Label)

# Add Depredation as the 71st element (central node)
depredation_row <- tibble(
  Label    = "Depredation",
  Category = "Central Concept",
  Tags     = paste(sort(unique(region_map)), collapse = "|")
)
elements_df <- bind_rows(elements_df, depredation_row) %>% arrange(Label)

cat(sprintf("Elements: %d unique labels\n", nrow(elements_df)))


# ---- Build Connections (A1, A2, A3, A5) -------------------------------------

# Resolve all raw rows to (FromLabel, ToLabel, SignedStrength, Region)
# Using sapply/mapply -- vectorised, avoids slow rowwise()
conns_resolved <- conns_raw %>%
  mutate(
    FromLabel      = sapply(From,     resolve_label),
    ToLabel        = sapply(To,       resolve_label),
    SignedStrength = ifelse(           # A1: sign convention
      Type == "Yes",
      as.integer(Strength),
      -as.integer(Strength)),
    Region         = mapply(resolve_region, From, To)   # A3
  ) %>%
  filter(!is.na(Region))

cat(sprintf("Raw connection rows (resolved): %d\n", nrow(conns_resolved)))

# Deduplicate: (FromLabel, ToLabel, SignedStrength) -> one row per unique triple
# Same pair with DIFFERENT strengths -> kept as separate rows (A5)
connections_df <- conns_resolved %>%
  group_by(FromLabel, ToLabel, SignedStrength) %>%
  summarise(
    Tags = paste(sort(unique(Region)), collapse = "|"),  # A6
    .groups = "drop"
  ) %>%
  mutate(
    `Influence Type` = if_else(SignedStrength > 0, "Positive", "Negative"),
    `Strength Label` = case_when(
      SignedStrength ==  2 ~ "Strong Positive",
      SignedStrength ==  1 ~ "Moderate Positive",
      SignedStrength == -1 ~ "Moderate Negative",
      SignedStrength == -2 ~ "Strong Negative",
      TRUE                 ~ as.character(SignedStrength)
    )
  ) %>%
  rename(From = FromLabel, To = ToLabel, Strength = SignedStrength) %>%
  select(From, To, Strength, `Influence Type`, `Strength Label`, Tags) %>%
  arrange(From, To, Strength)

n_removed <- nrow(conns_resolved) - nrow(connections_df)
cat(sprintf("Connections after deduplication: %d rows (%d duplicates removed)\n",
            nrow(connections_df), n_removed))

multi <- connections_df %>%
  group_by(From, To) %>% filter(n() > 1) %>%
  summarise(Strengths = paste(sort(Strength), collapse = ", "), .groups = "drop")
cat(sprintf("Pairs kept with multiple strength rows: %d\n", nrow(multi)))


# ---- Write to Excel ---------------------------------------------------------
wb <- createWorkbook()

h_style  <- createStyle(fontName = "Arial", fontSize = 11, fontColour = "#FFFFFF",
                        fgFill = "#2F4F4F", halign = "CENTER", textDecoration = "Bold")
d_style  <- createStyle(fontName = "Arial", fontSize = 10)
pos_sty  <- createStyle(fontName = "Arial", fontSize = 10, fgFill = "#C8E6C9")
neg_sty  <- createStyle(fontName = "Arial", fontSize = 10, fgFill = "#FFCDD2")

# Elements sheet
addWorksheet(wb, "Elements")
writeData(wb, "Elements", elements_df)
addStyle(wb, "Elements", h_style, rows = 1, cols = 1:3)
addStyle(wb, "Elements", d_style,
         rows = 2:(nrow(elements_df)+1), cols = 1:3, gridExpand = TRUE)
setColWidths(wb, "Elements", cols = 1:3, widths = c(44, 36, 56))
freezePane(wb, "Elements", firstRow = TRUE)

# Connections sheet
addWorksheet(wb, "Connections")
writeData(wb, "Connections", connections_df)
addStyle(wb, "Connections", h_style, rows = 1, cols = 1:6)
pos_rows <- which(connections_df$Strength > 0) + 1
neg_rows <- which(connections_df$Strength < 0) + 1
if (length(pos_rows) > 0)
  addStyle(wb, "Connections", pos_sty, rows = pos_rows, cols = 1:6, gridExpand = TRUE)
if (length(neg_rows) > 0)
  addStyle(wb, "Connections", neg_sty, rows = neg_rows, cols = 1:6, gridExpand = TRUE)
setColWidths(wb, "Connections", cols = 1:6, widths = c(44, 44, 12, 18, 22, 56))
freezePane(wb, "Connections", firstRow = TRUE)

# Notes sheet
addWorksheet(wb, "Notes & Assumptions")
notes_df <- tibble(
  Item = c("TAG FORMAT","","SIGN CONVENTION","","DEDUPLICATION","","",
           "REGIONS","","","","ELEMENTS","CONNECTIONS","",
           "STRENGTH SCALE","","","","KUMU IMPORT",""),
  Detail = c(
    "Tags use | (pipe, no spaces). Example: North Australia|Western Australia",
    "Kumu splits on |. Avoid '; ' -- trailing space breaks tag matching (A6).",
    "Type='Yes' in Master -> Positive; Type='No' -> Negative (A1)",
    "Verified: Deterrents->Depredation=Negative; Shark Habituation->Depredation=Positive",
    "Unique key = (From Label, To Label, Signed Strength) (A5)",
    "Same key in multiple regions -> ONE row, all regions joined with |",
    "Same (From, To) with DIFFERENT strengths -> kept as SEPARATE rows",
    "Western Australia   : Perth, Geraldton, Exmouth, Broome",
    "North Australia     : Darwin",
    "Queensland          : Townsville, Brisbane",
    "South East Australia: Coffs Harbour  [project wrote 'South West'; corrected to SE geographically]",
    paste(nrow(elements_df), "unique labels = 70 from Master Label column + Depredation central node (A4)"),
    paste(nrow(connections_df), "unique rows from", nrow(conns_resolved), "raw rows (A5)"),
    paste(n_removed, "duplicates removed;", nrow(multi), "pairs kept as separate rows for different strengths"),
    "+2 = Strong Positive   (strongly increases target)",
    "+1 = Moderate Positive (moderately increases target)",
    "-1 = Moderate Negative (moderately decreases target)",
    "-2 = Strong Negative   (strongly decreases target)",
    "kumu.io -> New Project -> Import -> Spreadsheet -> upload this file",
    "For region filtering: in each View add connection filter Tags.contains('Region Name')"
  )
)
writeData(wb, "Notes & Assumptions", notes_df)
bold_style <- createStyle(fontName="Arial", fontSize=10, textDecoration="Bold")
plain_style <- createStyle(fontName="Arial", fontSize=10)
addStyle(wb, "Notes & Assumptions", bold_style,
         rows=2:(nrow(notes_df)+1), cols=1, gridExpand=TRUE)
addStyle(wb, "Notes & Assumptions", plain_style,
         rows=2:(nrow(notes_df)+1), cols=2, gridExpand=TRUE)
setColWidths(wb, "Notes & Assumptions", cols=1:2, widths=c(22, 92))

# Save
saveWorkbook(wb, output_file, overwrite = TRUE)
cat(sprintf("\nSaved: %s\n", output_file))
cat(sprintf("  Elements sheet   : %d rows\n", nrow(elements_df)))
cat(sprintf("  Connections sheet: %d rows\n", nrow(connections_df)))


