# =============================================================================
# SCRIPT 2: Reduce Aggregated Galveston Model  →  30 Concepts
# Input : Exported_Kumu_Galveston_30April.xlsx  (61 concepts, 105 connections)
# Output: Kumu_Galveston_Reduced30.xlsx         (Kumu format, 2 sheets)
#         MentalModeler_Galveston_Reduced30.csv  (adjacency matrix)
# =============================================================================

# ── Packages ──────────────────────────────────────────────────────────────────
required_pkgs <- c("dplyr", "tidyr", "readxl", "openxlsx")
for (pkg in required_pkgs) {
  if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
}
library(dplyr); library(tidyr); library(readxl); library(openxlsx)


# =============================================================================
# MERGE MAP  (61 original concepts → 30 new concepts)
#
# Format: list( "NEW CONCEPT NAME" = c("original1", "original2", ...) )
#
# Rationale for each group:
#   ECOLOGICAL    – merge behaviourally/ecologically similar drivers
#   HUMAN DIM.    – merge perception/satisfaction/knowledge sub-concepts
#   POLICY/ECON   – merge overlapping regulatory and economic-loss nodes
#   FRM           – merge HMS nodes; merge shark fishery nodes
#   OPERATIONS    – merge gear damage, operational adaptations, shrimping
# =============================================================================

merge_map <- list(
  
  # ── CENTRAL CONCEPT (1) ────────────────────────────────────────────────────
  "Shark Depredation" =
    c("Shark Depredation"),
  
  # ── ECOLOGICAL & BIOLOGICAL (5) ───────────────────────────────────────────
  
  # Shark Populations + Shifting Distributions: both describe shark abundance
  # and spatial dynamics; shifting distributions is an attribute of populations.
  "Shark Populations" =
    c("Shark Populations",
      "Shifting Distributions Of Shark Populations"),
  
  # Shark Behavioral Responses: three concepts describing shark behaviour that
  # increases depredation risk (conditioned learning, vessel-following, hooked
  # fish attracting other sharks). Merged because they share the same causal
  # pathway: behaviour → depredation.
  "Shark Behavioral Responses" =
    c("Shark Learning Behavior",
      "Shark Vessel–Food Association (Attraction And Following Fishing Boats Due To Learned Food Rewards)",
      "Hooked Fish Behavior"),
  
  # Shark Attraction To Fishing Activity: physical attraction of sharks to
  # fishing vessels/activity (recreational boats following shrimpers, shark
  # attraction to vessels). Distinct from Behavioral Responses because it
  # captures the spatial/attractant mechanism rather than learned behaviour.
  "Shark Attraction To Fishing Activity" =
    c("Shark Attraction To Vessels",
      "Recreational Fishing Boats Following Commercial Fishing Vessels. This Attracts Sharks To Shrimper Boats"),
  
  # Habitat & Ecosystem Quality: structural habitat features that modulate
  # fish aggregations and thus depredation encounters. Includes artificial reefs
  # (which can replace removed platforms), platform removal (habitat loss), and
  # dolphin populations (co-predator that affects shark presence/behaviour).
  "Habitat & Ecosystem Quality" =
    c("Habitat Loss (Removal Of Rigs And Platforms)",
      "Increased Number Of Artificial Reefs"),
  
  # Environmental & Climate Drivers: long-term abiotic forcing factors.
  # All three push shark distributions toward fishing areas independently of
  # fishery operations.
  "Environmental & Climate Drivers" =
    c("Climate Change",
      "Water Temperature",
      "Bp Oil Spill Reduced Prey Populations"),
  
  "Dolphin Populations" =
    c("Dolphin Populations"),
  
  # ── HUMAN DIMENSIONS (7) ──────────────────────────────────────────────────
  
  # Public Perception & Media: the social-media / health-trend feedback loop
  # that amplifies public interest in fishing and shapes public opinion about
  # sharks. Health Conscious → Social Media → Public Positive Perception
  # form a tight cluster; merged into single node.
  "Public Perception & Media" =
    c("Public Positive Perception",
      "Increased Fishing Interest In Social Media",
      "Health Conscious",
      "People's Desire To Save Sharks"),
  
  # Fisher & Client Satisfaction: satisfaction outcomes for charter clients
  # and recreational anglers. Client Satisfaction and Fisher Satisfaction both
  # respond to depredation; Interest by Charter Clients is downstream of
  # Client Satisfaction → collapsed into one node.
  "Fisher & Client Satisfaction" =
    c("Client Satisfaction",
      "Fisher Satisfaction With Specific Species Fishing",
      "Interest By Charter Clients"),
  
  # Safety & Social Concerns: personal safety concerns for both crew (on-board)
  # and the public (fear of sharks near shore). Both are human-dimension
  # impacts of depredation.
  "Safety & Social Concerns" =
    c("Crew Safety",
      "Public Fear - Dangling Your Feet"),
  
  # Knowledge & Awareness: epistemic concepts — knowing depredation is
  # happening, and being able to identify the shark species involved.
  # Ability to Identify is a downstream consequence of Knowledge.
  "Knowledge & Awareness" =
    c("Knowledge Of Depredation",
      "Ability To Identify Shark Species"),
  
  # Fishermen's Income: key socioeconomic outcome; kept separate because it
  # has multiple incoming connections from different sectors.
  "Fishermen's Income" =
    c("Fishermen's Income"),
  
  # Recreational Participation: two concepts capturing growth in recreational
  # fishing pressure and engagement in tournaments; both increase shark
  # encounters.
  "Recreational Participation" =
    c("More Recreational Fishermen",
      "Participation In Tournaments - Shark Mutilate Catch Ex: Bigrock In Nc"),
  
  # ── POLICY & ECONOMICS (5) ────────────────────────────────────────────────
  
  # Regulatory Framework: overlapping regulatory instruments — Legislation,
  # Federal Enforcement, and the Shark Finning Law all control or incentivise
  # shark catch/conservation through legal means.
  "Regulatory Framework" =
    c("Legislation",
      "Government Enforcement And Federal Regulations",
      "Shark Finning Law - Illegal To Separate Fins From Body"),
  
  # Shark Conservation & Quota Policy: the two concepts governing the
  # management target for shark populations (quotas reduce populations; shark
  # conservation rebuilds them). Kept distinct from Regulatory Framework
  # because they directly affect Shark Populations.
  "Shark Conservation & Quota Policy" =
    c("Shark Conservation",
      "Increase Shark Quotas"),
  
  # Economic Losses from Depredation: the suite of monetary and time costs
  # directly caused by depredation events. Merged because they all flow from
  # gear/vessel damage and represent the same causal pathway to income loss.
  "Economic Losses from Depredation" =
    c("Economic Loss",
      "Labor Time To Repair Gear",
      "Cost Of Fines From Gear Damage (Shrimpers)",
      "Lack Of Economic Incentives To Overcome The Impacts Of Depredation"),
  
  # Fishing Profitability: kept separate as it is a key outcome node with
  # direct links from Catch and from Shark Depredation.
  "Fishing Profitability" =
    c("Fishing Profitability"),
  
  # Shark Fin Market: unique economic driver of the directed shark fishery;
  # kept separate because it has a distinct causal role (market demand → fishery
  # pressure → shark population reduction → depredation reduction).
  "Shark Fin Market and Demand" =
    c("Shark Fin Market", "Demand For Sharks"),
  
  # ── FISHERIES RESEARCH & MANAGEMENT (4) ───────────────────────────────────
  
  # HMS Management: both HMS nodes represent the same management authority
  # (Highly Migratory Species division of NOAA); one captures proactive actions,
  # the other the lag/slow response—merged because they feed the same outcomes.
  "HMS Management" =
    c("HMS Management Actions",
      "HMS Slow Management Response"),
  
  # Fisheries Management Effectiveness (Reef Fish): kept separate — it captures
  # a cross-sector effect where reef fish management success reduces prey
  # availability for sharks near platforms.
  "Fisheries Management Effectiveness" =
    c("Fisheries Management Effectiveness (Reef Fish)"),
  
  # Commercial Shark Fisheries: the directed and incidental shark harvest. Shark
  # Fishery (Charter sector concept) and Commercial Shark Fisheries both reduce
  # shark populations; merged into one node.
  "Commercial Shark Fisheries" =
    c("Commercial Shark Fisheries",
      "Shark Fishery"),
  
  # Research & Funding: kept separate — it is a management lever with a direct
  # negative effect on depredation (via knowledge production).
  "Research & Funding" =
    c("Research And Funding"),
  
  # ── FISHERIES OPERATIONS & PRACTICES (8) ──────────────────────────────────
  
  # Gear & Vessel Damage: both are direct physical damage outcomes of
  # depredation events; Vessel Damage is a severe form of Gear Damage.
  "Gear & Vessel Damage" =
    c("Damage To Gear",
      "Damage To Vessels By The Sharks Frenzy"),
  
  # Gear Costs: kept separate because it is an economic metric driven by
  # multiple upstream nodes (Gear Damage, Vessel Damage).
  "Gear Costs" =
    c("Gear Costs"),
  
  # Fishing Behavioral Adaptations: operational changes fishers make in
  # response to depredation — changing lures, locations, techniques, and trip
  # types. All four reduce (or alter) shark-fishery interactions.
  "Fishing Behavioral Adaptations" =
    c("Switch From Bait To Lures/Jigs",
      "Change Fishing Locations",
      "Effort In Cleaning And Processing Sharks",
      "Catch And Release Trips"),
  
  # Shrimping Operations: the commercial shrimping sector. Efficiency captures
  # productivity; Number of Boats determines total fleet effort; BRDs reduce
  # bycatch and incidental shark interactions. All three co-determine the
  # shrimping-sector's vulnerability to depredation.
  "Shrimping Operations" =
    c("Number of Shrimp Boats",
      "Bycatch Reduction Devices (Shrimp Fishery)",
      "Efficiency Of Commercial Fishery"),
  
  # Fishing Time: kept separate — it is a key intermediate between gear
  # repairs (time lost) and fishery productivity.
  "Fishing Time" =
    c("Fishing Time"),
  
  # Catch In Commercial Fishing: kept separate — direct productivity output
  # with clear links to Income and Profitability.
  "Catch In Commercial Fishing" =
    c("Catch In Commercial Fishing"),
  
  # Recreational Fishing Hotspots: spatial concentration of recreational effort
  # (hotspots) and species switching. Both increase encounter probability with
  # sharks.
  "Recreational Fishing Hotspots" =
    c("More Frequent Fishing Spots (Hotspots)",
      "Changing Species Targeted By Recreational Fishing"),
  
  # Insurance & Charter Risks: kept separate — unique to the charter sector;
  # captures the insurance/liability dimension of depredation.
  "Insurance & Charter Risks" =
    c("Insurance Risks For Charters")
)

# Verify: exactly 30 new concepts
stopifnot(length(merge_map) == 30)
cat(paste0("Merge map: ", length(merge_map), " new concepts defined.\n\n"))

# Print merge table for reference
cat("CONCEPT MERGE SUMMARY\n")
cat(strrep("=", 70), "\n")
for (new_name in names(merge_map)) {
  old <- merge_map[[new_name]]
  if (length(old) == 1 && old == new_name) {
    cat(sprintf("  [KEEP]  %s\n", new_name))
  } else {
    cat(sprintf("  [MERGE] %s\n          <- %s\n",
                new_name, paste(old, collapse = "\n          <- ")))
  }
}
cat("\n")


# =============================================================================
# CATEGORY ASSIGNMENTS FOR NEW CONCEPTS
# =============================================================================

new_categories <- c(
  "Shark Depredation"                     = "Central Concept",
  "Shark Populations"                     = "Ecological & Biological Factors",
  "Shark Behavioral Responses"            = "Ecological & Biological Factors",
  "Shark Attraction To Fishing Activity"  = "Ecological & Biological Factors",
  "Habitat & Ecosystem Quality"           = "Ecological & Biological Factors",
  "Environmental & Climate Drivers"       = "Ecological & Biological Factors",
  "Public Perception & Media"             = "Human Dimensions",
  "Fisher & Client Satisfaction"          = "Human Dimensions",
  "Safety & Social Concerns"              = "Human Dimensions",
  "Knowledge & Awareness"                 = "Human Dimensions",
  "Shark Demand & Conservation Attitudes" = "Human Dimensions",
  "Fishermen's Income"                    = "Human Dimensions",
  "Recreational Participation"            = "Human Dimensions",
  "Regulatory Framework"                  = "Policy & Economics",
  "Shark Conservation & Quota Policy"     = "Policy & Economics",
  "Economic Losses from Depredation"      = "Policy & Economics",
  "Fishing Profitability"                 = "Policy & Economics",
  "Shark Fin Market"                      = "Policy & Economics",
  "HMS Management"                        = "Fisheries Research & Management",
  "Fisheries Management Effectiveness"    = "Fisheries Research & Management",
  "Commercial Shark Fisheries"            = "Fisheries Research & Management",
  "Research & Funding"                    = "Fisheries Research & Management",
  "Gear & Vessel Damage"                  = "Fisheries Operations & Practices",
  "Gear Costs"                            = "Fisheries Operations & Practices",
  "Fishing Behavioral Adaptations"        = "Fisheries Operations & Practices",
  "Shrimping Operations"                  = "Fisheries Operations & Practices",
  "Fishing Time"                          = "Fisheries Operations & Practices",
  "Catch In Commercial Fishing"           = "Fisheries Operations & Practices",
  "Recreational Fishing Hotspots"         = "Fisheries Operations & Practices",
  "Insurance & Charter Risks"             = "Fisheries Operations & Practices"
)


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

# Build reverse lookup: old concept name → new concept name
build_reverse <- function(mmap) {
  rev_map <- character(0)
  for (new_name in names(mmap)) {
    for (old_name in mmap[[new_name]]) {
      rev_map[old_name] <- new_name
    }
  }
  rev_map
}

# Union Tags strings (e.g. "Commercial|Charter" + "Recreational" → "Charter|Commercial|Recreational")
union_tags <- function(tag_vec) {
  tags <- unlist(strsplit(tag_vec[!is.na(tag_vec) & tag_vec != ""], "\\|"))
  paste(sort(unique(trimws(tags))), collapse = "|")
}

# Round weight to nearest 0.5 step {-1, -0.5, 0, 0.5, 1}
round_half <- function(x) round(x * 2) / 2

# Derive Influence Type, Strength, and Strength Label from weight
weight_labels <- function(w) {
  data.frame(
    Influence_Type = case_when(w > 0 ~ "Positive", w < 0 ~ "Negative", TRUE ~ NA_character_),
    Strength       = case_when(
      w ==  1.0 ~ 2L, w ==  0.5 ~ 1L,
      w == -0.5 ~ -1L, w == -1.0 ~ -2L, TRUE ~ 0L),
    Strength_Label = case_when(
      w ==  1.0 ~ "Strong Positive",   w ==  0.5 ~ "Moderate Positive",
      w == -0.5 ~ "Moderate Negative", w == -1.0 ~ "Strong Negative",
      TRUE ~ NA_character_),
    stringsAsFactors = FALSE
  )
}


# =============================================================================
# LOAD ORIGINAL KUMU MODEL
# =============================================================================

cat("Reading Exported_Kumu_Galveston_30April.xlsx ...\n")

kumu_el    <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Exported_Kumu_Galveston_30April.xlsx",  sheet = "Elements")
kumu_conn <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Exported_Kumu_Galveston_30April.xlsx",  sheet = "Connections")

rev_map <- build_reverse(merge_map)

# Quick sanity check: flag original concepts not in merge map
orig_concepts <- kumu_el$Label
missing <- orig_concepts[!orig_concepts %in% names(rev_map)]
if (length(missing) > 0) {
  cat("WARNING — these original concepts are NOT in the merge map:\n")
  cat(paste(" ", missing, collapse = "\n"), "\n")
} else {
  cat("  All", length(orig_concepts), "original concepts mapped successfully.\n")
}


# =============================================================================
# BUILD NEW ELEMENTS SHEET
# =============================================================================

cat("\nBuilding new Elements sheet...\n")

# For each original element, record its new name and Tags
kumu_el$NewLabel <- rev_map[kumu_el$Label]

# Collect tags by new concept name
new_elements_tags <- kumu_el %>%
  group_by(NewLabel) %>%
  summarise(Tags = union_tags(Tags), .groups = "drop") %>%
  rename(Label = NewLabel)

# Descriptions summarising what was merged
merge_descriptions <- sapply(names(merge_map), function(new_name) {
  old <- merge_map[[new_name]]
  if (length(old) == 1 && old == new_name) return(NA_character_)
  paste0("Merged from: ", paste(old, collapse = "; "))
})

new_elements <- data.frame(
  Label       = names(merge_map),
  Type        = NA,
  Tags        = new_elements_tags$Tags[match(names(merge_map), new_elements_tags$Label)],
  Description = merge_descriptions,
  Category    = new_categories[names(merge_map)],
  stringsAsFactors = FALSE, row.names = NULL
)

cat(sprintf("  New Elements: %d concepts\n", nrow(new_elements)))


# =============================================================================
# BUILD NEW CONNECTIONS SHEET
# =============================================================================

cat("Building new Connections sheet...\n")

# Parse original weights
kumu_conn$weight_num <- suppressWarnings(as.numeric(kumu_conn$Original_MM_Weight))

# Remap From and To
kumu_conn$NewFrom <- rev_map[kumu_conn$From]
kumu_conn$NewTo   <- rev_map[kumu_conn$To]

# Remove rows where source or target not in map (shouldn't happen)
conn_mapped <- kumu_conn %>%
  filter(!is.na(NewFrom), !is.na(NewTo)) %>%
  # Remove self-loops created by merging
  filter(NewFrom != NewTo)

cat(sprintf("  Connections after removing self-loops: %d (from %d original)\n",
            nrow(conn_mapped), nrow(kumu_conn)))

# For each (NewFrom, NewTo, Tags) group, average the weight and round to 0.5
# Tags are preserved per sector: if two original edges both had Tags="Charter",
# they collapse to one; edges from different sectors stay separate rows.
conn_resolved <- conn_mapped %>%
  group_by(NewFrom, NewTo, Tags) %>%
  summarise(mean_weight = mean(weight_num, na.rm = TRUE), .groups = "drop") %>%
  mutate(resolved_weight = round_half(mean_weight)) %>%
  # Remove effectively-zero connections (opposing signs that cancel)
  filter(abs(resolved_weight) >= 0.5)

cat(sprintf("  Connections after weight resolution: %d\n", nrow(conn_resolved)))

# Build label columns for Kumu
wl <- weight_labels(conn_resolved$resolved_weight)

new_connections <- data.frame(
  From             = conn_resolved$NewFrom,
  To               = conn_resolved$NewTo,
  Direction        = "directed",
  Label            = NA,
  Type             = NA,
  Tags             = conn_resolved$Tags,
  Description      = NA,
  Category         = NA,
  Influence_Type   = wl$Influence_Type,
  Original_MM_Weight = as.character(conn_resolved$resolved_weight),
  Strength         = wl$Strength,
  Strength_Label   = wl$Strength_Label,
  stringsAsFactors = FALSE, row.names = NULL
)

# Rename columns to match Kumu export format
names(new_connections)[names(new_connections) == "Influence_Type"]  <- "Influence Type"
names(new_connections)[names(new_connections) == "Strength_Label"]  <- "Strength Label"

cat(sprintf("  Final connection count: %d\n", nrow(new_connections)))


# =============================================================================
# EXPORT: KUMU XLSX
# =============================================================================

cat("\nExporting Kumu_Galveston_Reduced30.xlsx ...\n")

wb <- createWorkbook()
addWorksheet(wb, "Elements")
addWorksheet(wb, "Connections")

# Style header rows
header_style <- createStyle(fontColour = "#FFFFFF", fgFill = "#2F4F4F",
                            textDecoration = "bold", halign = "center",
                            border = "Bottom", borderColour = "#000000")

writeData(wb, "Elements",     new_elements,     startRow = 1, startCol = 1)
writeData(wb, "Connections",  new_connections,  startRow = 1, startCol = 1)
addStyle(wb, "Elements",    header_style, rows = 1,
         cols = seq_len(ncol(new_elements)),    gridExpand = TRUE)
addStyle(wb, "Connections", header_style, rows = 1,
         cols = seq_len(ncol(new_connections)), gridExpand = TRUE)
setColWidths(wb, "Elements",    cols = 1:ncol(new_elements),    widths = "auto")
setColWidths(wb, "Connections", cols = 1:ncol(new_connections), widths = "auto")

saveWorkbook(wb, "Kumu_Galveston_Reduced30.xlsx", overwrite = TRUE)
cat("  Saved: Kumu_Galveston_Reduced30.xlsx\n")


# =============================================================================
# EXPORT: MENTAL MODELER CSV (adjacency matrix)
#
# Format:  rows = From concept, columns = To concept, values = mean weight
#          across all sector rows for that (From, To) pair.
#          Empty cell = no connection.  Headers: blank first cell, then concept names.
# =============================================================================

cat("Exporting MentalModeler_Galveston_Reduced30.csv ...\n")

# For Mental Modeler use the mean weight across sectors
mm_edges <- conn_mapped %>%
  group_by(NewFrom, NewTo) %>%
  summarise(mm_weight = round_half(mean(weight_num, na.rm = TRUE)), .groups = "drop") %>%
  filter(abs(mm_weight) >= 0.5)

# Create 30×30 matrix
concept_order <- names(merge_map)     # preserves the logical ordering above
n <- length(concept_order)

mm_mat <- matrix(NA_character_, nrow = n, ncol = n,
                 dimnames = list(concept_order, concept_order))

for (i in seq_len(nrow(mm_edges))) {
  r <- mm_edges$NewFrom[i]
  c <- mm_edges$NewTo[i]
  if (r %in% concept_order && c %in% concept_order) {
    mm_mat[r, c] <- as.character(mm_edges$mm_weight[i])
  }
}

# Replace NA with empty string (Mental Modeler convention)
mm_mat[is.na(mm_mat)] <- ""

# Write as CSV with blank header in position 1
mm_df <- as.data.frame(mm_mat, stringsAsFactors = FALSE)
mm_df <- cbind(data.frame(Concept = rownames(mm_df),
                          stringsAsFactors = FALSE), mm_df)
rownames(mm_df) <- NULL

write.csv(mm_df, "MentalModeler_Galveston_Reduced30.csv",
          row.names = FALSE, na = "")
cat("  Saved: MentalModeler_Galveston_Reduced30.csv\n")


# =============================================================================
# SUMMARY REPORT
# =============================================================================

cat("\n", strrep("=", 65), "\n")
cat("REDUCTION SUMMARY\n")
cat(strrep("=", 65), "\n")
cat(sprintf("  Original concepts  : %d\n", nrow(kumu_el)))
cat(sprintf("  Reduced concepts   : %d\n", nrow(new_elements)))
cat(sprintf("  Concepts removed   : %d\n", nrow(kumu_el) - nrow(new_elements)))
cat(sprintf("  Original connections: %d\n", nrow(kumu_conn)))
cat(sprintf("  Self-loops removed  : %d\n", nrow(kumu_conn) - nrow(conn_mapped)))
cat(sprintf("  After deduplication : %d\n", nrow(new_connections)))
cat("\n")

cat("CONCEPTS BY NEW CATEGORY\n")
cat(strrep("-", 45), "\n")
print(table(new_categories), zero.print = ".")

cat("\nKEY SELF-LOOPS REMOVED (due to merging):\n")
self_loops <- kumu_conn %>%
  mutate(NewFrom = rev_map[From], NewTo = rev_map[To]) %>%
  filter(!is.na(NewFrom), !is.na(NewTo), NewFrom == NewTo) %>%
  select(From, To, Tags, Original_MM_Weight)
print(self_loops, row.names = FALSE)

cat("\n✓ Script 2 complete.\n")
cat("  Output files:\n")
cat("    Kumu_Galveston_Reduced30.xlsx        (Kumu import format)\n")
cat("    MentalModeler_Galveston_Reduced30.csv (Mental Modeler adjacency matrix)\n")