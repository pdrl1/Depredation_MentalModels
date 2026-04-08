# =============================================================================
# Kumu → Mental Modeler CSV Converter
# Input:  Master.xlsx (two sheets: Elements, Connections)
# Output: 5 CSV files in Mental Modeler adjacency-matrix format
#         - All_Australia_MentalModeler.csv
#         - North_Australia_MentalModeler.csv
#         - Queensland_MentalModeler.csv
#         - New_South_Wales_MentalModeler.csv
#         - Western_Australia_MentalModeler.csv
#
# =============================================================================
# ASSUMPTIONS & CHANGES DOCUMENTED
# =============================================================================
#
# [A1] DATA SOURCE — Connections are read from the "Connections" sheet (455
#      valid rows). The "Elements" sheet also embeds one connection per element
#      in columns From/To/Direction/Strength/Type (columns G–K). Comparing the
#      two sheets revealed that 296 of those 297 embedded connections are already
#      present in the Connections sheet. The one that is not (Culling → ID 164,
#      Geraldton) references element ID 164, which does not exist anywhere in
#      the Elements sheet and was therefore treated as a data-entry error in the
#      original spreadsheet. It is excluded. The Connections sheet is used as
#      the sole authoritative source.
#
# [A2] 'DEPREDATION' NODE — 'Depredation' appears as From/To in the Connections
#      sheet but has no corresponding row in the Elements sheet. It is treated
#      as a valid concept node present in all regions where it participates.
#
# [A3] CONNECTION SIGN (TYPE COLUMN) — The 'Type' column in the Connections
#      sheet encodes relationship polarity:
#        Type = 'Yes'  →  POSITIVE influence (concept increases the target)
#        Type = 'No'   →  NEGATIVE influence (concept decreases the target)
#      'Yes ' (trailing whitespace) is normalised to 'Yes'.
#      This dataset has 309 positive and 146 negative connections.
#
# [A4] STRENGTH → MENTAL MODELER SCALE — Kumu uses magnitude 1 (weak) or 2
#      (strong). Combined with Type (sign), these map to Mental Modeler's
#      −1..+1 scale as follows:
#        Strength 2, Type Yes  →  +1.0  (strong positive)
#        Strength 1, Type Yes  →  +0.5  (weak positive)
#        Strength 1, Type No   →  −0.5  (weak negative)
#        Strength 2, Type No   →  −1.0  (strong negative)
#
# [A5] WORKSHOP ASSIGNMENT FOR CONNECTIONS — A connection is assigned to a
#      workshop (and thus a region) based on the FROM element's workshop.
#      Exception: when From = 'Depredation', the connection is assigned to the
#      TO element's workshop (Depredation has no workshop of its own).
#      There are 0 cross-workshop element-to-element connections in this dataset
#      so this rule is unambiguous.
#
# [A6] REGION → WORKSHOP MAPPING (as specified):
#        North Australia   : Darwin
#        Queensland        : Townsville, Brisbane
#        New South Wales   : Coffs Harbour  ← NOTE: data spells it 'Coffs Harbour'
#                                               (not 'Coffs Harbor' as in the request)
#        Western Australia : Perth, Geraldton, Exmouth, Broome
#      Connections whose Conn_Workshop is NA (cannot be assigned) are only
#      present in the All Australia output.
#
# [A7] BIDIRECTIONAL CONNECTIONS — 'undirected' and 'mutual' connections are
#      each expanded into TWO directed edges (A→B and B→A) with the same
#      signed strength. 'directed ' (trailing space) is normalised to
#      'directed'. Counts in dataset: 401 directed, 30 undirected, 21 mutual,
#      3 'directed ' (space).
#
# [A8] AVERAGING DUPLICATE CONNECTIONS — Mental Modeler accepts only one value
#      per (From, To) concept-pair. Within each regional output, all connections
#      sharing the same (From_Label, To_Label) pair are AVERAGED (mean).
#      This occurs when different workshops within the same region independently
#      drew the same conceptual link with differing strengths or signs. 55
#      concept labels are shared across multiple workshops. Averaging preserves
#      group consensus and can yield intermediate values (e.g. −0.25, 0.75).
#
# [A9] CONCEPT INCLUSION — All concepts that appear as either From or To in at
#      least one connection for a given region are included in that region's
#      matrix. Every one of the 298 elements in the dataset participates in at
#      least one connection, so no element is silently dropped.
#
# [A10] SELF-LOOPS — No self-loop connections (A→A) exist in the source data.
#       Any that might arise from bidirectional expansion are removed.
#
# [A11] EMPTY ROWS — The Connections sheet contains 926 completely empty rows
#       (spreadsheet artefact). These are silently dropped.
#
# [A12] OUTPUT FORMAT — Each CSV is an adjacency matrix where:
#         Row    = FROM concept (cause / influencer)
#         Column = TO concept  (effect / influenced)
#         Cell   = averaged signed influence strength
#                  (0 means no connection; non-zero values ∈ {±0.5, ±1.0} or
#                   intermediate averages such as ±0.25, ±0.75, etc.)
#         First column header = "Concepts", contains row concept names.
#       This matches the Mental Modeler CSV import format.
#
# =============================================================================

# ── Packages ──────────────────────────────────────────────────────────────────
# Install if missing:

library(readxl)
library(dplyr)
library(tidyr)
library(stringr)

# =============================================================================
# PARAMETERS — edit these as needed
# =============================================================================

input_file  <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/Mental Models from Marcus/Australia/Master.xlsx"
output_dir <- "."             # Directory where output CSVs will be written

# Region → Workshop mapping  [A6]
# NOTE: use exact spellings found in the data file
region_map <- list(
  "North_Australia"   = c("Darwin"),
  "Queensland"        = c("Townsville", "Brisbane"),
  "New_South_Wales"   = c("Coffs Harbour"),   # data: 'Harbour', not 'Harbor'
  "Western_Australia" = c("Perth", "Geraldton", "Exmouth", "Broome")
)

# =============================================================================
# STEP 1 — Load data
# =============================================================================
cat("========================================================\n")
cat("Kumu -> Mental Modeler Converter\n")
cat("========================================================\n\n")
cat("Loading data from:", input_file, "\n")

elements_raw    <- read_excel(input_file, sheet = "Elements")
connections_raw <- read_excel(input_file, sheet = "Connections")

# Keep only the columns we need; drop entirely empty rows  [A11]
elements <- elements_raw %>%
  select(ID, Label, Category, State, Workshop) %>%
  filter(!is.na(ID)) %>%
  mutate(
    ID       = as.character(ID),
    Label    = as.character(Label),
    Workshop = as.character(Workshop)
  )

connections <- connections_raw %>%
  select(From, To, Direction, Strength, Type) %>%   # 'Type ' has trailing space in sheet
  rename(Type = Type) %>%
  filter(!is.na(From)) %>%                              # drop empty rows  [A11]
  mutate(
    From      = as.character(From),
    To        = as.character(To),
    Direction = trimws(as.character(Direction)),         # normalise trailing spaces  [A7]
    Strength  = as.numeric(Strength),
    Type      = trimws(as.character(Type))              # normalise 'Yes ' → 'Yes'  [A3]
  )

cat(sprintf("  Elements loaded   : %d rows\n", nrow(elements)))
cat(sprintf("  Connections loaded: %d rows (after dropping empty rows)\n\n",
            nrow(connections)))

# =============================================================================
# STEP 2 — Build lookup tables: element ID → Label and Workshop
# =============================================================================
id_to_label    <- setNames(elements$Label,    elements$ID)
id_to_workshop <- setNames(elements$Workshop, elements$ID)

# =============================================================================
# STEP 3 — Resolve IDs to labels, assign workshops  [A2][A5]
# =============================================================================
connections <- connections %>%
  mutate(
    # Resolve From
    From_Label = case_when(
      From == "Depredation"           ~ "Depredation",
      From %in% names(id_to_label)    ~ id_to_label[From],
      TRUE                            ~ NA_character_
    ),
    From_Workshop = case_when(
      From == "Depredation"           ~ NA_character_,  # Depredation has no workshop
      From %in% names(id_to_workshop) ~ id_to_workshop[From],
      TRUE                            ~ NA_character_
    ),
    
    # Resolve To
    To_Label = case_when(
      To == "Depredation"             ~ "Depredation",
      To %in% names(id_to_label)      ~ id_to_label[To],
      TRUE                            ~ NA_character_
    ),
    To_Workshop = case_when(
      To == "Depredation"             ~ NA_character_,
      To %in% names(id_to_workshop)   ~ id_to_workshop[To],
      TRUE                            ~ NA_character_
    )
  )

# Warn about unresolvable IDs (e.g. the Culling → 164 orphan from Elements)  [A1]
unresolved <- connections %>% filter(is.na(From_Label) | is.na(To_Label))
if (nrow(unresolved) > 0) {
  cat(sprintf("NOTE: %d connection(s) reference element IDs not found in the Elements\n",
              nrow(unresolved)))
  cat("      sheet and will be skipped (see assumption A1):\n")
  print(unresolved %>% select(From, To, Direction, Strength, Type))
  cat("\n")
}

connections <- connections %>% filter(!is.na(From_Label) & !is.na(To_Label))

# Assign each connection to a workshop  [A5]
connections <- connections %>%
  mutate(
    Conn_Workshop = if_else(From == "Depredation", To_Workshop, From_Workshop)
  )

cat(sprintf("Connections after label resolution: %d\n", nrow(connections)))
cat("Workshop distribution of connections:\n")
print(table(connections$Conn_Workshop, useNA = "ifany"))
cat("\n")

# =============================================================================
# STEP 4 — Expand bidirectional connections  [A7]
# =============================================================================
bidirectional_types <- c("undirected", "mutual")

directed_conns <- connections %>%
  filter(!Direction %in% bidirectional_types)

bidir_forward <- connections %>%
  filter(Direction %in% bidirectional_types)

# Reverse: swap From↔To using temp columns to avoid circular-rename issues
bidir_reverse <- bidir_forward %>%
  mutate(
    tmp_From     = To,          tmp_To     = From,
    tmp_FL       = To_Label,    tmp_TL     = From_Label,
    tmp_FW       = To_Workshop, tmp_TW     = From_Workshop
  ) %>%
  select(-From, -To, -From_Label, -To_Label, -From_Workshop, -To_Workshop) %>%
  rename(
    From          = tmp_From,  To          = tmp_To,
    From_Label    = tmp_FL,    To_Label    = tmp_TL,
    From_Workshop = tmp_FW,    To_Workshop = tmp_TW
  ) %>%
  mutate(
    Conn_Workshop = if_else(From == "Depredation", To_Workshop, From_Workshop)
  )

connections_expanded <- bind_rows(directed_conns, bidir_forward, bidir_reverse)

cat(sprintf("After bidirectional expansion: %d connections\n",
            nrow(connections_expanded)))
cat(sprintf("  (%d reverse edges added from %d undirected/mutual connections)\n\n",
            nrow(bidir_reverse), nrow(bidir_forward)))

#bidirectional_types <- c("undirected", "mutual") correctly catches both categories. 
#bidir_reverse swaps From↔To and is bound with bind_rows(directed_conns, bidir_forward, bidir_reverse), 
#producing exactly the 51 additional edges. The 401 directed connections flow through directed_conns untouched.

#Averaging within regions (including bidirectional overlaps): This is the most important part. 
#Because bind_rows keeps all copies of a From→To pair — whether they originate from two separate directed connections, 
#two workshops, or a bidirectional expansion overlapping a directed connection — 
#group_by(From_Label, To_Label) %>% summarise(MM_Strength = mean(...)) in make_mm_csv averages all of them together. 
#This is exactly what the paragraph describes.
# =============================================================================
# STEP 5 — Compute signed Mental Modeler strength  [A3][A4]
# =============================================================================
# Type = 'Yes' → positive (+), Type = 'No' → negative (−)
# Strength 1 → magnitude 0.5, Strength 2 → magnitude 1.0

connections_expanded <- connections_expanded %>%
  mutate(
    Sign = case_when(
      Type == "Yes" ~  1,
      Type == "No"  ~ -1,
      TRUE          ~  NA_real_
    ),
    Magnitude = case_when(
      Strength == 1 ~ 0.5,
      Strength == 2 ~ 1.0,
      TRUE          ~ NA_real_
    ),
    MM_Strength = Sign * Magnitude
  )

# Report and drop any connections with missing Type or Strength
bad <- sum(is.na(connections_expanded$MM_Strength))
if (bad > 0) {
  cat(sprintf("WARNING: %d connection(s) have an unrecognised Type or Strength", bad))
  cat(" and will be skipped:\n")
  print(connections_expanded %>%
          filter(is.na(MM_Strength)) %>%
          select(From_Label, To_Label, Direction, Strength, Type))
  cat("\n")
  connections_expanded <- connections_expanded %>% filter(!is.na(MM_Strength))
}

# Remove self-loops  [A10]
self_loops <- connections_expanded %>% filter(From_Label == To_Label)
if (nrow(self_loops) > 0) {
  cat(sprintf("NOTE: Removed %d self-loop(s).\n\n", nrow(self_loops)))
  connections_expanded <- connections_expanded %>% filter(From_Label != To_Label)
}

# Summary of signs
pos_n <- sum(connections_expanded$MM_Strength > 0)
neg_n <- sum(connections_expanded$MM_Strength < 0)
cat(sprintf("Final connection pool: %d total  (%d positive, %d negative)\n\n",
            nrow(connections_expanded), pos_n, neg_n))

# =============================================================================
# STEP 6 — Helper: build and write one Mental Modeler CSV  [A8][A9][A12]
# =============================================================================

make_mm_csv <- function(conn_subset, region_name, output_dir) {
  
  if (nrow(conn_subset) == 0) {
    cat(sprintf("  WARNING: No connections found for '%s'. Skipping.\n", region_name))
    return(invisible(NULL))
  }
  
  # Average duplicate (From_Label, To_Label) pairs within this region  [A8]
  conn_avg <- conn_subset %>%
    group_by(From_Label, To_Label) %>%
    summarise(
      MM_Strength = mean(MM_Strength, na.rm = TRUE),
      n_sources   = n(),        # how many raw connections contributed
      .groups     = "drop"
    )
  #group_by(From_Label, To_Label) %>% summarise(MM_Strength = mean(...)) in make_mm_csv averages all of them together. 
  
  n_avg_pairs <- sum(conn_avg$n_sources > 1)
  if (n_avg_pairs > 0) {
    cat(sprintf("  Averaged %d concept-pair(s) with multiple connections.\n",
                n_avg_pairs))
  }
  
  # All unique concepts in this region (rows AND columns)  [A9]
  all_concepts <- sort(unique(c(conn_avg$From_Label, conn_avg$To_Label)))
  n_concepts   <- length(all_concepts)
  
  # Build adjacency matrix initialised to 0
  mat <- matrix(0,
                nrow     = n_concepts,
                ncol     = n_concepts,
                dimnames = list(all_concepts, all_concepts))
  
  for (i in seq_len(nrow(conn_avg))) {
    f <- conn_avg$From_Label[i]
    t <- conn_avg$To_Label[i]
    mat[f, t] <- conn_avg$MM_Strength[i]
  }
  
  # Convert to data frame; prepend concept-name column  [A12]
  df <- as.data.frame(mat, check.names = FALSE)
  df <- cbind(Concepts = rownames(df), df)
  rownames(df) <- NULL
  
  # Write CSV
  filename <- file.path(output_dir,
                        paste0(region_name, "_MentalModeler.csv"))
  write.csv(df, filename, row.names = FALSE, quote = TRUE)
  
  n_pos <- sum(conn_avg$MM_Strength > 0)
  n_neg <- sum(conn_avg$MM_Strength < 0)
  cat(sprintf("  Written : %s\n", basename(filename)))
  cat(sprintf("            %d concepts | %d connections (%d positive, %d negative)",
              n_concepts, nrow(conn_avg), n_pos, n_neg))
  
  # Flag intermediate averaged values (not cleanly ±0.5 or ±1.0)
  non_std <- conn_avg %>%
    filter(MM_Strength != 0, !MM_Strength %in% c(0.5, 1.0, -0.5, -1.0))
  if (nrow(non_std) > 0) {
    cat(sprintf("\n            %d pair(s) have intermediate averaged values (e.g. ±0.25, ±0.75)",
                nrow(non_std)))
  }
  cat("\n")
  
  invisible(df)
}

# =============================================================================
# STEP 7 — All Australia
# =============================================================================
cat("--------------------------------------------------------\n")
cat("Generating: All Australia\n")
make_mm_csv(connections_expanded, "All_Australia", output_dir)

# =============================================================================
# STEP 8 — One CSV per region
# =============================================================================
for (region_name in names(region_map)) {
  cat("--------------------------------------------------------\n")
  workshops_in_region <- region_map[[region_name]]
  cat(sprintf("Generating: %s  (workshops: %s)\n",
              gsub("_", " ", region_name),
              paste(workshops_in_region, collapse = ", ")))
  
  region_conns <- connections_expanded %>%
    filter(Conn_Workshop %in% workshops_in_region)
  
  make_mm_csv(region_conns, region_name, output_dir)
}

# =============================================================================
# STEP 9 — Done
# =============================================================================
cat("========================================================\n")
cat("Done! Output files written to:", normalizePath(output_dir), "\n")
cat("========================================================\n\n")
cat("REMINDER [A6]: 'Coffs Harbour' (data spelling) = 'Coffs Harbor'\n")
cat("(user request spelling). The NSW file uses the data spelling.\n\n")
cat("REMINDER [A8]: Averaged values can be intermediate (e.g. -0.25,\n")
cat("0.75). These occur where two workshops in the same region drew the\n")
cat("same conceptual link with different strengths or opposite signs.\n")
cat("Review these pairs manually if needed.\n")


#Averaged pairs that result in exactly 0 are still written to conn_avg internally 
#and assigned a 0 in the matrix, rather than being explicitly filtered out beforehand. 
#This is inconsequential for Mental Modeler (0 = no connection), 
#but it means n_connections in the summary line will not count that pair (since only MM_Strength > 0 and MM_Strength < 0 are tallied). 
#So the printed connection counts will correctly exclude those zero-averaged pairs. Everything is consistent.


conn_subset <- connections_expanded   # or filter for a region

conn_avg <- conn_subset %>%
  group_by(From_Label, To_Label) %>%
  summarise(
    MM_Strength = mean(MM_Strength, na.rm = TRUE),
    n_sources   = n(),
    .groups     = "drop"
  )

View(conn_avg)




