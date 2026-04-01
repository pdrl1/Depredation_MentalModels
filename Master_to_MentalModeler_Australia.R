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
# [A1] DATA SOURCE — Connections are read exclusively from the "Connections"
#      sheet (455 valid rows). The "Elements" sheet also embeds one connection
#      per element (each element's primary link to 'Depredation'), but these
#      are a subset of the Connections sheet. Using the Connections sheet avoids
#      duplication and ensures all inter-element connections are captured.
#
# [A2] 'DEPREDATION' NODE — 'Depredation' appears as From/To in the Connections
#      sheet but is NOT listed as a row in the Elements sheet. It is treated as
#      a valid concept node and is included in every regional CSV in which it
#      participates (via its connecting elements).
#
# [A3] WORKSHOP ASSIGNMENT FOR CONNECTIONS — A connection is assigned to a
#      workshop (and thus a region) based on the FROM element's workshop.
#      Exception: when From = 'Depredation', the connection is assigned to the
#      TO element's workshop (since 'Depredation' itself has no workshop).
#      There are 0 cross-workshop element-to-element connections in this dataset,
#      so this rule never causes ambiguity for those connections.
#
# [A4] REGION → WORKSHOP MAPPING (as specified by user):
#      - North Australia  : Darwin
#      - Queensland       : Townsville, Brisbane
#      - New South Wales  : Coffs Harbour   ← NOTE: data spells it 'Coffs Harbour'
#                                               (not 'Coffs Harbor' as in the request)
#      - Western Australia: Perth, Geraldton, Exmouth, Broome
#      Connections not belonging to any of the above workshops are only included
#      in the All Australia output.
#
# [A5] BIDIRECTIONAL CONNECTIONS — Connections with Direction = 'undirected' or
#      'mutual' are expanded into TWO directed connections (A→B and B→A), both
#      with the same strength. 'directed ' (trailing whitespace) is normalised
#      to 'directed'. In this dataset: 401 directed, 30 undirected, 21 mutual,
#      3 'directed ' (with space).
#
# [A6] STRENGTH MAPPING — All strength values in the data are 1 (weak) or 2
#      (strong). There are NO negative relationship values in the dataset.
#      They are mapped to the Mental Modeler −1..+1 scale as follows:
#        Kumu 1  →  Mental Modeler +0.5  (weak positive influence)
#        Kumu 2  →  Mental Modeler +1.0  (strong positive influence)
#      IMPORTANT: All connections are treated as POSITIVE (increasing) influences
#      because no negative/decreasing valence is recorded in the source data.
#      Users should manually review and flip signs where the real-world
#      relationship is negative (e.g. "more enforcement → less depredation").
#
# [A7] AVERAGING DUPLICATE CONNECTIONS — Mental Modeler cannot handle multiple
#      arrows between the same concept pair. Within each regional output, all
#      connections sharing the same (From_Label, To_Label) pair are averaged.
#      This occurs when different workshops in the same region independently
#      drew the same conceptual link with different strengths (e.g. Exmouth and
#      Broome both drew "Fishing Effort → Depredation" with strengths 2 and 1
#      respectively; the WA output records 1.5 for that link).
#      55 concept labels appear in more than one workshop, so averaging is
#      meaningful for regional aggregation.
#
# [A8] CONCEPTS INCLUDED IN EACH REGIONAL CSV — Only concepts that appear in
#      at least one connection for that region are included in the matrix.
#      Concepts from other regions that happen to share the same label are still
#      distinct rows in the data but collapse when the label is used as the
#      matrix index (see A7).
#
# [A9] SELF-LOOPS — No self-loop connections (A→A) exist in the source data.
#      If any were introduced by bidirectional expansion (not expected), they
#      are removed.
#
# [A10] EMPTY CONNECTIONS — The Connections sheet contains 926 completely empty
#       rows (artefact of the spreadsheet layout). These are silently dropped.
#
# [A11] OUTPUT FORMAT — Each CSV is an adjacency matrix where:
#       - Row = FROM concept (cause)
#       - Column = TO concept (effect)
#       - Cell value = averaged, sign-adjusted influence strength (0, ±0.5, ±1.0)
#       - The first column is labelled "Concepts" and contains row concept names
#       This matches the Mental Modeler import format.
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

# Region → Workshop mapping  [A4]
# NOTE: use exact spellings from the data file
region_map <- list(
  "North_Australia"   = c("Darwin"),
  "Queensland"        = c("Townsville", "Brisbane"),
  "New_South_Wales"   = c("Coffs Harbour"),   # data uses 'Harbour', not 'Harbor'
  "Western_Australia" = c("Perth", "Geraldton", "Exmouth", "Broome")
)

# =============================================================================
# STEP 1 — Load data
# =============================================================================
cat("========================================================\n")
cat("Kumu → Mental Modeler Converter\n")
cat("========================================================\n\n")
cat("Loading data from:", input_file, "\n")

elements_raw    <- read_excel(input_file, sheet = "Elements")
connections_raw <- read_excel(input_file, sheet = "Connections")

# Keep only the columns we need; drop entirely empty rows  [A10]
elements <- elements_raw %>%
  select(ID, Label, Category, State, Workshop) %>%
  filter(!is.na(ID)) %>%
  mutate(
    ID       = as.character(ID),
    Label    = as.character(Label),
    Workshop = as.character(Workshop)
  )

connections <- connections_raw %>%
  select(From, To, Direction, Strength) %>%
  filter(!is.na(From)) %>%                      # drop empty rows  [A10]
  mutate(
    From      = as.character(From),
    To        = as.character(To),
    Direction = trimws(as.character(Direction)), # normalise trailing spaces  [A5]
    Strength  = as.numeric(Strength)
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
# STEP 3 — Resolve IDs to labels and assign workshops  [A2][A3]
# =============================================================================
connections <- connections %>%
  mutate(
    # Resolve From
    From_Label = case_when(
      From == "Depredation"            ~ "Depredation",
      From %in% names(id_to_label)     ~ id_to_label[From],
      TRUE                             ~ NA_character_
    ),
    From_Workshop = case_when(
      From == "Depredation"            ~ NA_character_,   # Depredation has no workshop
      From %in% names(id_to_workshop)  ~ id_to_workshop[From],
      TRUE                             ~ NA_character_
    ),

    # Resolve To
    To_Label = case_when(
      To == "Depredation"              ~ "Depredation",
      To %in% names(id_to_label)       ~ id_to_label[To],
      TRUE                             ~ NA_character_
    ),
    To_Workshop = case_when(
      To == "Depredation"              ~ NA_character_,
      To %in% names(id_to_workshop)    ~ id_to_workshop[To],
      TRUE                             ~ NA_character_
    )
  )

# Warn about any unresolvable IDs
unresolved <- connections %>% filter(is.na(From_Label) | is.na(To_Label))
if (nrow(unresolved) > 0) {
  cat(sprintf("WARNING: %d connection(s) could not be resolved to labels",
              nrow(unresolved)))
  cat(" and will be skipped:\n")
  print(unresolved %>% select(From, To, Direction, Strength))
  cat("\n")
}

connections <- connections %>% filter(!is.na(From_Label) & !is.na(To_Label))

# Assign connection to a workshop  [A3]
#   - If From is a real element → use From element's workshop
#   - If From is 'Depredation'  → use To element's workshop
connections <- connections %>%
  mutate(
    Conn_Workshop = if_else(From == "Depredation", To_Workshop, From_Workshop)
  )

cat(sprintf("Connections after label resolution: %d\n", nrow(connections)))
cat("Workshop distribution of connections:\n")
print(table(connections$Conn_Workshop, useNA = "ifany"))
cat("\n")

# =============================================================================
# STEP 4 — Expand bidirectional connections  [A5]
# =============================================================================
bidirectional_types <- c("undirected", "mutual")

directed_conns <- connections %>%
  filter(!Direction %in% bidirectional_types)

bidir_forward <- connections %>%
  filter(Direction %in% bidirectional_types)

# Reverse direction: swap From↔To labels and workshops
# We use mutate() rather than rename() for the swap to avoid circular-rename issues
bidir_reverse <- bidir_forward %>%
  mutate(
    tmp_From          = To,           tmp_To          = From,
    tmp_From_Label    = To_Label,     tmp_To_Label    = From_Label,
    tmp_From_Workshop = To_Workshop,  tmp_To_Workshop = From_Workshop
  ) %>%
  select(-From, -To, -From_Label, -To_Label, -From_Workshop, -To_Workshop) %>%
  rename(
    From          = tmp_From,          To          = tmp_To,
    From_Label    = tmp_From_Label,    To_Label    = tmp_To_Label,
    From_Workshop = tmp_From_Workshop, To_Workshop = tmp_To_Workshop
  ) %>%
  # Re-assign Conn_Workshop after the swap
  mutate(
    Conn_Workshop = if_else(From == "Depredation", To_Workshop, From_Workshop)
  )

connections_expanded <- bind_rows(directed_conns, bidir_forward, bidir_reverse)

cat(sprintf("Connections after bidirectional expansion: %d\n",
            nrow(connections_expanded)))
cat(sprintf("  (added %d reverse edges from %d undirected/mutual connections)\n\n",
            nrow(bidir_reverse), nrow(bidir_forward)))

# =============================================================================
# STEP 5 — Map Kumu strength to Mental Modeler scale  [A6]
# =============================================================================
connections_expanded <- connections_expanded %>%
  mutate(
    MM_Strength = case_when(
      Strength == 1 ~ 0.5,
      Strength == 2 ~ 1.0,
      TRUE          ~ NA_real_    # should not occur in this dataset
    )
  )

na_str <- sum(is.na(connections_expanded$MM_Strength))
if (na_str > 0) {
  cat(sprintf("WARNING: %d connection(s) have an unrecognised strength value",
              na_str))
  cat(" and will be skipped.\n\n")
  connections_expanded <- connections_expanded %>% filter(!is.na(MM_Strength))
}

# Remove self-loops (just in case)  [A9]
self_loops <- connections_expanded %>% filter(From_Label == To_Label)
if (nrow(self_loops) > 0) {
  cat(sprintf("NOTE: Removed %d self-loop(s).\n\n", nrow(self_loops)))
  connections_expanded <- connections_expanded %>% filter(From_Label != To_Label)
}

# =============================================================================
# STEP 6 — Helper function: build and write a Mental Modeler CSV  [A7][A8][A11]
# =============================================================================

make_mm_csv <- function(conn_subset, region_name, output_dir) {

  if (nrow(conn_subset) == 0) {
    cat(sprintf("  WARNING: No connections found for '%s'. Skipping.\n", region_name))
    return(invisible(NULL))
  }

  # Average duplicate (From_Label, To_Label) pairs  [A7]
  conn_avg <- conn_subset %>%
    group_by(From_Label, To_Label) %>%
    summarise(
      MM_Strength = mean(MM_Strength, na.rm = TRUE),
      n_averaged  = n(),
      .groups     = "drop"
    )

  n_averaged_pairs <- sum(conn_avg$n_averaged > 1)
  if (n_averaged_pairs > 0) {
    cat(sprintf("  Averaged %d concept-pair(s) that had multiple connections.\n",
                n_averaged_pairs))
  }

  # All unique concepts (both as source and target)  [A8]
  all_concepts <- sort(unique(c(conn_avg$From_Label, conn_avg$To_Label)))
  n_concepts   <- length(all_concepts)

  # Build adjacency matrix (initialised to 0)
  mat <- matrix(0,
                nrow     = n_concepts,
                ncol     = n_concepts,
                dimnames = list(all_concepts, all_concepts))

  for (i in seq_len(nrow(conn_avg))) {
    f <- conn_avg$From_Label[i]
    t <- conn_avg$To_Label[i]
    mat[f, t] <- conn_avg$MM_Strength[i]
  }

  # Convert to data frame; first column = concept names  [A11]
  df <- as.data.frame(mat, check.names = FALSE)
  df <- cbind(Concepts = rownames(df), df)
  rownames(df) <- NULL

  # Write CSV
  filename <- file.path(output_dir,
                        paste0(region_name, "_MentalModeler.csv"))
  write.csv(df, filename, row.names = FALSE, quote = TRUE)

  cat(sprintf("  Written: %s\n", basename(filename)))
  cat(sprintf("           %d concepts, %d unique connections",
              n_concepts, nrow(conn_avg)))

  # Report any values that are not standard (0, 0.5, 1.0)
  # after averaging, values like 0.75 are possible
  non_standard <- conn_avg %>%
    filter(MM_Strength != 0, MM_Strength != 0.5, MM_Strength != 1.0)
  if (nrow(non_standard) > 0) {
    cat(sprintf(", %d averaged value(s) are non-standard (e.g. 0.75)",
                nrow(non_standard)))
  }
  cat("\n")

  invisible(df)
}

# =============================================================================
# STEP 7 — Generate All Australia CSV
# =============================================================================
cat("--------------------------------------------------------\n")
cat("Generating: All Australia\n")
make_mm_csv(connections_expanded, "All_Australia", output_dir)

# =============================================================================
# STEP 8 — Generate one CSV per region
# =============================================================================
for (region_name in names(region_map)) {
  cat("--------------------------------------------------------\n")
  workshops_in_region <- region_map[[region_name]]
  cat(sprintf("Generating: %s (workshops: %s)\n",
              gsub("_", " ", region_name),
              paste(workshops_in_region, collapse = ", ")))

  region_conns <- connections_expanded %>%
    filter(Conn_Workshop %in% workshops_in_region)

  make_mm_csv(region_conns, region_name, output_dir)
}

# =============================================================================
# STEP 9 — Summary
# =============================================================================
cat("========================================================\n")
cat("Done! Output files written to:", normalizePath(output_dir), "\n")
cat("========================================================\n\n")
cat("REMINDER [A6]: All connections are encoded as POSITIVE influences\n")
cat("(+0.5 = weak, +1.0 = strong). No negative valence was present in\n")
cat("the source data. Please review and manually flip signs where the\n")
cat("real-world relationship is inhibiting/decreasing.\n\n")
cat("REMINDER [A4]: 'Coffs Harbour' (data spelling) = 'Coffs Harbor'\n")
cat("(user spelling). The NSW output uses the data spelling.\n")