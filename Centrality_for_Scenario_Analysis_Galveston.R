# =============================================================================
# SCRIPT 1: Centrality Analysis — Galveston Shark Depredation Mental Models
# Models: Recreational, Charter, Commercial (Mental Modeler CSV)
#         Aggregated full (61 concepts) and Aggregated reduced (30 concepts) (Kumu XLSX)
#
# DEGREE CENTRALITY METHOD — matches Goal 1 and Goal 2 exactly:
#   • WEIGHTED directed graph (edge weight = signed causal strength)
#   • Indegree_i  = Σ_j |w_ji|   (sum of ABSOLUTE weights of incoming edges)
#   • Outdegree_i = Σ_j |w_ij|   (sum of ABSOLUTE weights of outgoing edges)
#   • TotalDegree = Indegree + Outdegree
#   • NO normalisation (Goal 1 / Goal 2 report raw weighted degree)
#   • Repeated From→To pairs are collapsed to one edge by averaging signed
#     strength (deduplicate_edges rule), before summing.
#
#   This replaces the previous UNWEIGHTED, normalised degree() calculation,
#   which counted edges and ignored link strength — not comparable to Goal 1/2.
#
# Outputs: centrality tables (CSV), top-5 tables per group.
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
# CONFIGURATION
# =============================================================================
# Mental Modeler CSVs are on the ±0.5 / ±1.0 scale; Goal 1/2 (and the Kumu
# exports) use the integer ±1 / ±2 scale. Multiply the CSV weights by 2 so all
# five models share the SAME integer scale and their weighted degrees are
# directly comparable. Set to 1 if your CSVs already contain ±1 / ±2 values.
MM_SCALE  <- 2
KUMU_SCALE <- 1   # Kumu 'Strength' column is already integer ±1 / ±2


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

# Read Mental Modeler adjacency-matrix CSV (rows = From, cols = To)
read_mm_csv <- function(filepath) {
  raw <- read.csv(filepath, header = TRUE, check.names = FALSE,
                  row.names = 1, stringsAsFactors = FALSE)
  raw[raw == "" | is.na(raw)] <- 0
  mat <- as.matrix(raw)
  storage.mode(mat) <- "numeric"
  mat[is.na(mat)] <- 0
  return(mat)
}

# ---- Weighted degree — IDENTICAL to compute_degrees() in Goal 1 / Goal 2 -----
#   Indegree_i  = Σ_j |w_ji|   Outdegree_i = Σ_j |w_ij|   Degree = In + Out
compute_degrees <- function(g) {
  edge_df <- igraph::as_data_frame(g, what = "edges")
  in_deg <- edge_df %>%
    group_by(to) %>%
    summarise(indegree = sum(abs(weight)), .groups = "drop") %>%
    rename(name = to)
  out_deg <- edge_df %>%
    group_by(from) %>%
    summarise(outdegree = sum(abs(weight)), .groups = "drop") %>%
    rename(name = from)
  data.frame(name = V(g)$name) %>%
    left_join(in_deg,  by = "name") %>%
    left_join(out_deg, by = "name") %>%
    mutate(
      indegree  = tidyr::replace_na(indegree,  0),
      outdegree = tidyr::replace_na(outdegree, 0),
      degree    = indegree + outdegree
    )
}

# Build a WEIGHTED directed igraph from a Mental Modeler adjacency matrix.
# Non-zero cell (i, j) = directed edge From row i To col j, weight = signed value.
build_graph_from_matrix <- function(mat, scale = 1) {
  mat <- mat * scale
  idx <- which(mat != 0, arr.ind = TRUE)
  all_nodes <- unique(c(rownames(mat), colnames(mat)))
  if (nrow(idx) == 0) {
    return(graph_from_data_frame(
      data.frame(from = character(0), to = character(0),
                 weight = numeric(0), polarity = numeric(0)),
      directed = TRUE, vertices = data.frame(name = all_nodes)))
  }
  edges <- data.frame(
    from     = rownames(mat)[idx[, 1]],
    to       = colnames(mat)[idx[, 2]],
    weight   = mat[idx],
    stringsAsFactors = FALSE
  )
  edges$polarity <- sign(edges$weight)
  graph_from_data_frame(edges, directed = TRUE,
                        vertices = data.frame(name = all_nodes))
}

# Build a WEIGHTED directed igraph from a Kumu XLSX (Elements + Connections).
# Uses the integer 'Strength' column; repeated From→To pairs are averaged
# (mean signed strength) exactly like deduplicate_edges() in Goal 1 / Goal 2.
build_graph_from_kumu <- function(el_df, conn_df, scale = 1) {
  conn <- conn_df %>%
    mutate(strength_num = suppressWarnings(as.numeric(Strength)) * scale) %>%
    filter(!is.na(From), !is.na(To), !is.na(strength_num))
  edges <- conn %>%
    group_by(From, To) %>%
    summarise(
      weight   = mean(strength_num, na.rm = TRUE),
      polarity = sign(mean(strength_num, na.rm = TRUE)),
      .groups  = "drop"
    ) %>%
    filter(!is.na(weight)) %>%
    rename(from = From, to = To)
  all_nodes <- unique(c(el_df$Label, edges$from, edges$to))
  all_nodes <- all_nodes[!is.na(all_nodes)]
  graph_from_data_frame(edges, directed = TRUE,
                        vertices = data.frame(name = all_nodes))
}

# Weighted centrality table for one model (weighted degree, unnormalised)
calc_centrality <- function(g, model_name) {
  nd <- compute_degrees(g)
  df <- data.frame(
    Model       = model_name,
    Concept     = nd$name,
    InDegree    = nd$indegree,    # Σ|incoming w|
    OutDegree   = nd$outdegree,   # Σ|outgoing w|
    TotalDegree = nd$degree,      # In + Out
    stringsAsFactors = FALSE, row.names = NULL
  )
  df <- df %>% arrange(desc(TotalDegree))
  df$Rank <- seq_len(nrow(df))
  df[, c("Rank", setdiff(names(df), "Rank"))]
}

# --- OPTIONAL: normalised degree centrality (NOT used by Goal 1/2) -----------
# If you also want the Freeman-normalised versions, uncomment and call after
# calc_centrality():
# add_norm_degree <- function(df, g) {
#   n <- vcount(g)
#   df %>% mutate(
#     InDegree_Norm    = round(InDegree    / max(n - 1, 1), 3),
#     OutDegree_Norm   = round(OutDegree   / max(n - 1, 1), 3),
#     DegreeCentrality = round(TotalDegree / max(2 * (n - 1), 1), 3))
# }

# Print unique edge-weight magnitudes so the scale can be verified
report_scale <- function(g, label) {
  w <- igraph::as_data_frame(g, what = "edges")$weight
  cat(sprintf("  %-22s unique |weights|: %s\n", label,
              paste(sort(unique(abs(w))), collapse = ", ")))
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
# 1. READ & PROCESS INDIVIDUAL SECTOR MODELS (Mental Modeler CSV)
# =============================================================================

cat("\n── Reading individual sector models ──\n")

rec_mat  <- read_mm_csv("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/MM/Recreational_FinalModel_22April.csv")
char_mat <- read_mm_csv("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/MM/Charter_FinalModel_22April.csv")
comm_mat <- read_mm_csv("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/MM/Commercial_FinalModel_22April.csv")

g_rec  <- build_graph_from_matrix(rec_mat,  scale = MM_SCALE)
g_char <- build_graph_from_matrix(char_mat, scale = MM_SCALE)
g_comm <- build_graph_from_matrix(comm_mat, scale = MM_SCALE)

cat("Edge-weight scale check (expect integer 1, 2 after MM_SCALE):\n")
report_scale(g_rec,  "Recreational")
report_scale(g_char, "Charter")
report_scale(g_comm, "Commercial")

rec_cent  <- calc_centrality(g_rec,  "Recreational")
char_cent <- calc_centrality(g_char, "Charter")
comm_cent <- calc_centrality(g_comm, "Commercial")

# Group assignment via pattern matching (no category column in CSV)
rec_cent$Group  <- sapply(rec_cent$Concept,  assign_group)
char_cent$Group <- sapply(char_cent$Concept, assign_group)
comm_cent$Group <- sapply(comm_cent$Concept, assign_group)


# =============================================================================
# 2. READ & PROCESS AGGREGATED KUMU MODEL — FULL (61 concepts)
# =============================================================================

cat("\n── Reading aggregated full model (61 concepts) ──\n")

kumu_el_full   <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Exported_Kumu_Galveston_30April.xlsx", sheet = "Elements")
kumu_conn_full <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Exported_Kumu_Galveston_30April.xlsx", sheet = "Connections")

g_full <- build_graph_from_kumu(kumu_el_full, kumu_conn_full, scale = KUMU_SCALE)
report_scale(g_full, "Aggregated (Full)")
agg_full_cent <- calc_centrality(g_full, "Aggregated (Full)")

# Attach category from Elements sheet and derive group
cat_lookup_full <- setNames(kumu_el_full$Category, kumu_el_full$Label)
agg_full_cent$Category <- cat_lookup_full[agg_full_cent$Concept]
agg_full_cent$Category[is.na(agg_full_cent$Category)] <- "Unknown"
agg_full_cent$Group <- kumu_cat_to_group(agg_full_cent$Category)


# =============================================================================
# 3. READ & PROCESS AGGREGATED KUMU MODEL — REDUCED (30 concepts)
# =============================================================================

cat("\n── Reading aggregated reduced model (30 concepts) ──\n")

kumu_el_red   <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu_Galveston_Reduced30.xlsx", sheet = "Elements")
kumu_conn_red <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu_Galveston_Reduced30.xlsx", sheet = "Connections")

g_red <- build_graph_from_kumu(kumu_el_red, kumu_conn_red, scale = KUMU_SCALE)
report_scale(g_red, "Aggregated (Reduced)")
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
                      "InDegree", "OutDegree", "TotalDegree"),
                    names(df))
  disp <- df[, cols]
  num_cols <- intersect(c("InDegree", "OutDegree", "TotalDegree"), names(disp))
  disp[num_cols] <- lapply(disp[num_cols], function(x) round(x, 3))
  print(disp, row.names = FALSE)
}

print_table(rec_cent,       "RECREATIONAL MODEL — CENTRALITY RANKINGS (weighted degree)")
print_table(char_cent,      "CHARTER MODEL — CENTRALITY RANKINGS (weighted degree)")
print_table(comm_cent,      "COMMERCIAL MODEL — CENTRALITY RANKINGS (weighted degree)")
print_table(agg_full_cent,  "AGGREGATED FULL MODEL (61) — CENTRALITY RANKINGS (weighted degree)")
print_table(agg_red_cent,   "AGGREGATED REDUCED MODEL (30) — CENTRALITY RANKINGS (weighted degree)")


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
             InDegree, OutDegree, TotalDegree)
    
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
# 5b. TOP 5 PER GROUP, EXCLUDING RECEIVERS
#     Groups: Ecological | Management | Fisheries Behavior
#     Receiver = outdegree = 0 AND indegree > 0 (pure outcome, no outgoing
#     links). Within each group, receivers are removed and the remaining
#     concepts are ranked by weighted TotalDegree.
# =============================================================================

top5_by_group_no_receivers <- function(cent_df, model_name) {
  cat(paste0("\n\n", strrep("-", 65), "\n"))
  cat(paste0("TOP 5 BY GROUP (RECEIVERS REMOVED) — ", toupper(model_name), "\n"))
  cat(paste0(strrep("-", 65), "\n"))
  
  groups <- c("Ecological", "Management", "Fisheries Behavior")
  result <- list()
  
  for (grp in groups) {
    sub <- cent_df %>%
      filter(Group == grp) %>%
      filter(!(OutDegree == 0 & InDegree > 0)) %>%   # drop receivers
      arrange(desc(TotalDegree)) %>%
      head(5) %>%
      mutate(Group = grp, Model = model_name, Rank_in_group = row_number()) %>%
      select(Model, Group, Rank_in_group, Concept,
             InDegree, OutDegree, TotalDegree)
    
    cat(paste0("\n  ► ", grp, "\n"))
    disp <- sub
    num_cols <- c("InDegree", "OutDegree", "TotalDegree")
    disp[num_cols] <- lapply(disp[num_cols], function(x) round(x, 3))
    print(disp, row.names = FALSE)
    result[[grp]] <- sub
  }
  bind_rows(result)
}

top5nr_rec      <- top5_by_group_no_receivers(rec_cent,      "Recreational")
top5nr_char     <- top5_by_group_no_receivers(char_cent,     "Charter")
top5nr_comm     <- top5_by_group_no_receivers(comm_cent,     "Commercial")
top5nr_agg_full <- top5_by_group_no_receivers(agg_full_cent, "Aggregated (Full)")
top5nr_agg_red  <- top5_by_group_no_receivers(agg_red_cent,  "Aggregated (Reduced)")

top5_no_receivers_all <- bind_rows(top5nr_rec, top5nr_char, top5nr_comm,
                                   top5nr_agg_full, top5nr_agg_red)
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
                   InDegree, OutDegree, TotalDegree),
          "centrality_aggregated_full.csv",     row.names = FALSE)
write.csv(agg_red_cent %>%
            select(Rank, Concept, Category, Group,
                   InDegree, OutDegree, TotalDegree),
          "centrality_aggregated_reduced.csv",  row.names = FALSE)
write.csv(top5_all,
          "top5_concepts_by_group_all_models.csv", row.names = FALSE)

cat("  All CSVs exported.\n")

cat("\n✓ Script 1 complete (weighted degree, matching Goal 1 / Goal 2).\n")
