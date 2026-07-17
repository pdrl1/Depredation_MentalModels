# =============================================================================
# SCRIPT 1: Centrality Analysis — Galveston Shark Depredation Mental Models
# Models: Recreational, Charter, Commercial (CSV)
#         Aggregated full (61 concepts) and Aggregated reduced (30 concepts) (XLSX)
# Metrics: In-degree, Out-degree, Total Degree, Normalised Degree Centrality
# Outputs: centrality tables (CSV), bar charts, scatter plots, heatmap,
#          top-5 tables per group (Ecological / Management / Fisheries Behavior)
# =============================================================================

# ── Packages ──────────────────────────────────────────────────────────────────
required_pkgs <- c("igraph", "ggplot2", "dplyr", "tidyr",
                   "readxl", "ggrepel", "RColorBrewer", "openxlsx")
for (pkg in required_pkgs) {
  if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
}
library(igraph); library(ggplot2); library(dplyr); library(tidyr)
library(readxl); library(ggrepel); library(RColorBrewer); library(openxlsx)


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

# Read Mental Modeler adjacency-matrix CSV
read_mm_csv <- function(filepath) {
  raw <- read.csv(filepath, header = TRUE, check.names = FALSE,
                  row.names = 1, stringsAsFactors = FALSE)
  raw[raw == "" | is.na(raw)] <- 0
  mat <- as.matrix(raw)
  storage.mode(mat) <- "numeric"
  mat[is.na(mat)] <- 0
  return(mat)
}

# Build directed igraph from adjacency matrix (non-zero = edge)
mat_to_igraph <- function(mat) {
  graph_from_adjacency_matrix(mat != 0, mode = "directed", weighted = NULL)
}

# Build directed igraph from a Kumu XLSX (Elements + Connections sheets)
kumu_to_igraph <- function(el_df, conn_df) {
  edges <- conn_df %>%
    select(From, To) %>%
    filter(!is.na(From), !is.na(To))
  nodes <- unique(c(el_df$Label, edges$From, edges$To))
  nodes <- nodes[!is.na(nodes)]
  graph_from_data_frame(edges, directed = TRUE,
                        vertices = data.frame(name = nodes))
}

# Calculate in-, out-, total-degree and normalised degree centrality
calc_centrality <- function(g, model_name) {
  n           <- vcount(g)
  indegree    <- degree(g, mode = "in")
  outdegree   <- degree(g, mode = "out")
  totaldegree <- degree(g, mode = "all")
  
  df <- data.frame(
    Model            = model_name,
    Concept          = V(g)$name,
    InDegree         = indegree,
    OutDegree        = outdegree,
    TotalDegree      = totaldegree,
    InDegree_Norm    = round(indegree    / max(n - 1, 1), 3),
    OutDegree_Norm   = round(outdegree   / max(n - 1, 1), 3),
    DegreeCentrality = round(totaldegree / max(2 * (n - 1), 1), 3),
    stringsAsFactors = FALSE, row.names = NULL
  )
  df <- df %>% arrange(desc(TotalDegree))
  df$Rank <- seq_len(nrow(df))
  df[, c("Rank", setdiff(names(df), "Rank"))]
}

# Map Kumu Category column → analysis Group
kumu_cat_to_group <- function(category_vec) {
  dplyr::case_when(
    category_vec == "Ecological & Biological Factors"              ~ "Ecological",
    category_vec %in% c("Fisheries Research & Management",
                        "Policy & Economics")                      ~ "Management",
    category_vec %in% c("Fisheries Operations & Practices",
                        "Human Dimensions")                        ~ "Fisheries Behavior",
    category_vec == "Central Concept"                              ~ "Central Concept",
    TRUE                                                            ~ "Fisheries Behavior"
  )
}

# Shorten long strings for plot labels
shorten <- function(x, n = 38) {
  ifelse(nchar(x) > n, paste0(substr(x, 1, n - 3), "..."), x)
}


# =============================================================================
# GROUP ASSIGNMENT FOR SECTOR CSVs
# (Pattern-based — used only for Recreational / Charter / Commercial,
#  which have no Category column)
# =============================================================================

assign_group <- function(concept_name) {
  cn <- tolower(trimws(concept_name))
  
  if (grepl("^shark depredation$", cn)) return("Central Concept")
  
  ecological <- c(
    "shark population", "shifting distribution", "habitat loss", "removal of rigs",
    "climate change", "water temperature", "oil spill", "prey population",
    "learning behavior", "hooked fish", "food association", "vessel.food association",
    "shark attraction", "dolphin", "artificial reef", "boats following commercial",
    "shrimper boat"
  )
  
  management <- c(
    "legislation", "government enforcement", "federal regulation",
    "hms management", "hms slow", "fisheries management effectiveness",
    "shark conservation", "increase shark quota", "shark fin",
    "shark finning", "research and funding", "commercial shark fisheri",
    "shark fishery", "lack of economic incentive"
  )
  
  if (any(sapply(ecological, function(p) grepl(p, cn, fixed = FALSE)))) return("Ecological")
  if (any(sapply(management, function(p) grepl(p, cn, fixed = FALSE)))) return("Management")
  return("Fisheries Behavior")
}

# ── Colour palettes (consistent across all plots) ─────────────────────────────
group_colors <- c(
  "Ecological"          = "#2E8B57",
  "Management"          = "#4169E1",
  "Fisheries Behavior"  = "#D2691E",
  "Central Concept"     = "#DC143C"
)

model_colors <- c(
  "Recreational"        = "#E41A1C",
  "Charter"             = "#377EB8",
  "Commercial"          = "#4DAF4A",
  "Aggregated (Full)"   = "#984EA3",
  "Aggregated (Reduced)"= "#FF7F00"
)


# =============================================================================
# 1. READ & PROCESS INDIVIDUAL SECTOR MODELS (CSV)
# =============================================================================

cat("\n── Reading individual sector models ──\n")

rec_mat  <- read_mm_csv("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/MM/Recreational_FinalModel_22April.csv")
char_mat <- read_mm_csv("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/MM/Charter_FinalModel_22April.csv")
comm_mat <- read_mm_csv("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/MM/Commercial_FinalModel_22April.csv")

rec_cent  <- calc_centrality(mat_to_igraph(rec_mat),  "Recreational")
char_cent <- calc_centrality(mat_to_igraph(char_mat), "Charter")
comm_cent <- calc_centrality(mat_to_igraph(comm_mat), "Commercial")

# Group assignment via pattern matching (no category column in CSV)
rec_cent$Group  <- sapply(rec_cent$Concept,  assign_group)
char_cent$Group <- sapply(char_cent$Concept, assign_group)
comm_cent$Group <- sapply(comm_cent$Concept, assign_group)


# =============================================================================
# 2. READ & PROCESS AGGREGATED KUMU MODEL — FULL (61 concepts)
# =============================================================================

cat("── Reading aggregated full model (61 concepts) ──\n")

kumu_el_full   <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Exported_Kumu_Galveston_30April.xlsx", sheet = "Elements")
kumu_conn_full <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Exported_Kumu_Galveston_30April.xlsx", sheet = "Connections")

g_full  <- kumu_to_igraph(kumu_el_full, kumu_conn_full)
agg_full_cent <- calc_centrality(g_full, "Aggregated (Full)")

# Attach category from Elements sheet and derive group
cat_lookup_full <- setNames(kumu_el_full$Category, kumu_el_full$Label)
agg_full_cent$Category <- cat_lookup_full[agg_full_cent$Concept]
agg_full_cent$Category[is.na(agg_full_cent$Category)] <- "Unknown"
agg_full_cent$Group <- kumu_cat_to_group(agg_full_cent$Category)


# =============================================================================
# 3. READ & PROCESS AGGREGATED KUMU MODEL — REDUCED (30 concepts)
# =============================================================================

cat("── Reading aggregated reduced model (30 concepts) ──\n")

kumu_el_red   <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu_Galveston_Reduced30.xlsx", sheet = "Elements")
kumu_conn_red <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu_Galveston_Reduced30.xlsx", sheet = "Connections")

g_red  <- kumu_to_igraph(kumu_el_red, kumu_conn_red)
agg_red_cent <- calc_centrality(g_red, "Aggregated (Reduced)")

# Attach category from Elements sheet and derive group
cat_lookup_red <- setNames(kumu_el_red$Category, kumu_el_red$Label)
agg_red_cent$Category <- cat_lookup_red[agg_red_cent$Concept]
agg_red_cent$Category[is.na(agg_red_cent$Category)] <- "Unknown"
agg_red_cent$Group <- kumu_cat_to_group(agg_red_cent$Category)


# =============================================================================
# 4. PRINT CENTRALITY RANKING TABLES (all 5 models)
# =============================================================================

print_table <- function(df, title) {
  cat(paste0("\n", strrep("=", 65), "\n"))
  cat(paste0(title, "\n"))
  cat(paste0(strrep("=", 65), "\n"))
  cols <- intersect(c("Rank", "Concept", "Category", "Group",
                      "InDegree", "OutDegree", "TotalDegree",
                      "InDegree_Norm", "OutDegree_Norm", "DegreeCentrality"),
                    names(df))
  print(df[, cols], row.names = FALSE)
}

print_table(rec_cent,       "RECREATIONAL MODEL — CENTRALITY RANKINGS")
print_table(char_cent,      "CHARTER MODEL — CENTRALITY RANKINGS")
print_table(comm_cent,      "COMMERCIAL MODEL — CENTRALITY RANKINGS")
print_table(agg_full_cent,  "AGGREGATED FULL MODEL (61) — CENTRALITY RANKINGS")
print_table(agg_red_cent,   "AGGREGATED REDUCED MODEL (30) — CENTRALITY RANKINGS")


# =============================================================================
# 5. TOP 5 CONCEPTS PER GROUP PER MODEL
#    Groups: Ecological | Management | Fisheries Behavior
# =============================================================================

top5_by_group <- function(cent_df, model_name) {
  cat(paste0("\n\n", strrep("-", 65), "\n"))
  cat(paste0("TOP 5 CONCEPTS BY GROUP — ", toupper(model_name), "\n"))
  cat(paste0(strrep("-", 65), "\n"))
  
  groups <- c("Ecological", "Management", "Fisheries Behavior")
  result <- list()
  
  for (grp in groups) {
    sub <- cent_df %>%
      filter(Group == grp) %>%
      arrange(desc(TotalDegree)) %>%
      head(5) %>%
      mutate(Group = grp, Model = model_name, Rank_in_group = row_number()) %>%
      select(Model, Group, Rank_in_group, Concept,
             InDegree, OutDegree, TotalDegree, DegreeCentrality)
    
    cat(paste0("\n  ► ", grp, "\n"))
    print(sub, row.names = FALSE)
    result[[grp]] <- sub
  }
  bind_rows(result)
}

top5_rec      <- top5_by_group(rec_cent,      "Recreational")
top5_char     <- top5_by_group(char_cent,     "Charter")
top5_comm     <- top5_by_group(comm_cent,     "Commercial")
top5_agg_full <- top5_by_group(agg_full_cent, "Aggregated (Full)")
top5_agg_red  <- top5_by_group(agg_red_cent,  "Aggregated (Reduced)")

top5_all <- bind_rows(top5_rec, top5_char, top5_comm,
                      top5_agg_full, top5_agg_red)


# =============================================================================
# 6. EXPORT CENTRALITY TABLES TO CSV
# =============================================================================

cat("\n── Exporting CSVs ──\n")

write.csv(rec_cent,
          "centrality_recreational.csv",       row.names = FALSE)
write.csv(char_cent,
          "centrality_charter.csv",             row.names = FALSE)
write.csv(comm_cent,
          "centrality_commercial.csv",          row.names = FALSE)
write.csv(agg_full_cent %>%
            select(Rank, Concept, Category, Group,
                   InDegree, OutDegree, TotalDegree,
                   InDegree_Norm, OutDegree_Norm, DegreeCentrality),
          "centrality_aggregated_full.csv",     row.names = FALSE)
write.csv(agg_red_cent %>%
            select(Rank, Concept, Category, Group,
                   InDegree, OutDegree, TotalDegree,
                   InDegree_Norm, OutDegree_Norm, DegreeCentrality),
          "centrality_aggregated_reduced.csv",  row.names = FALSE)
write.csv(top5_all,
          "top5_concepts_by_group_all_models.csv", row.names = FALSE)

cat("  All CSVs exported.\n")

cat("\n✓ Script 1 complete.\n")
cat("\n  CENTRALITY TABLES (CSV):\n")
