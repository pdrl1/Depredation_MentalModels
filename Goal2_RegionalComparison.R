# ============================================================
#  MENTAL MODEL ANALYSIS — GOAL 2
#  Regional scale: Within-workshop sub-region comparison
#
#  Australia sub-regions (from Tags column):
#    Western Australia · North Australia · Queensland · New South Wales
#
#  Gulf Coast USA sub-regions (from Tags column):
#    Alabama · Florida · Louisiana · Mississippi · Texas
#
#  For each sub-region this script computes:
#    — All structural network metrics (same suite as Goal 1)
#    — Concept-level and link-level Jaccard similarity matrices
#    — Polarity conflict detection across sub-regions
#    — Pairwise EDR, LDR, and MDR (Schaffernicht & Groesser 2011)
#      for all combinations within Australia (6 pairs)
#      and within Gulf Coast (10 pairs)
#
#  KEY REFERENCE
#  Schaffernicht, M. & Groesser, S. N. (2011). A comprehensive
#  method for comparing mental models of dynamic systems.
#  European Journal of Operational Research, 210(1), 57–67.
# ============================================================


# ============================================================
# SECTION 0 — PACKAGES
# ============================================================

required_pkgs <- c("readxl", "igraph", "dplyr", "tidyr",
                   "ggplot2", "stringr", "scales", "ggrepel", "purrr")
for (pkg in required_pkgs) {
  if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
}
library(readxl); library(igraph); library(dplyr); library(tidyr)
library(ggplot2); library(stringr); library(scales)
library(ggrepel);  library(purrr)


# ============================================================
# SECTION 1 — DATA LOADING
# ============================================================
# Both files are the Kumu master exports.
# Australia  → Strength column: signed integers ±1 / ±2
# Gulf Coast → Strength column: signed integers ±1 / ±2
#   (back-converted from Mental Modeler scale during standardisation)
# Tags column: pipe-separated sub-region labels,
#   e.g. "Western Australia|Queensland"

AU_FILE <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Australia/Kumu/Exported_Kumu_Australia_7May.xlsx"
AL_FILE <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Alabama/Kumu/Exported_Kumu_Alabama_30April.xlsx"

au_elements    <- read_excel(AU_FILE, sheet = "Elements")
au_connections <- read_excel(AU_FILE, sheet = "Connections")
al_elements    <- read_excel(AL_FILE, sheet = "Elements")
al_connections <- read_excel(AL_FILE, sheet = "Connections")

# Parse numeric strength once, drop rows with no strength
au_connections <- au_connections %>%
  mutate(strength_num = suppressWarnings(as.numeric(Strength))) %>%
  filter(!is.na(From), !is.na(To), !is.na(strength_num))

al_connections <- al_connections %>%
  mutate(strength_num = suppressWarnings(as.numeric(Strength))) %>%
  filter(!is.na(From), !is.na(To), !is.na(strength_num))

cat("Loaded — AU:", nrow(au_elements), "elements |",
    nrow(au_connections), "connections\n")
cat("Loaded — GC:", nrow(al_elements), "elements |",
    nrow(al_connections), "connections\n")

# Element group lookup tables (for coloring prominence plots)
# Australia uses the 'Type' column; Gulf Coast uses the 'Category' column.
au_groups <- au_elements %>%
  select(name = Label, Group = Type) %>%
  distinct() %>%
  filter(!is.na(Group))

gc_groups <- al_elements %>%
  select(name = Label, Group = Category) %>%
  distinct() %>%
  filter(!is.na(Group))

cat("AU groups:", paste(sort(unique(au_groups$Group)), collapse = ", "), "\n")
cat("GC groups:", paste(sort(unique(gc_groups$Group)), collapse = ", "), "\n")


# ============================================================
# SECTION 2 — SUB-REGION DEFINITIONS
# ============================================================
# These strings must match exactly what appears in the Tags column
# of the Kumu export files.

AU_SUBREGIONS <- c("Western Australia", "North Australia",
                   "Queensland",        "New South Wales")

GC_SUBREGIONS <- c("Alabama", "Florida", "Louisiana",
                   "Mississippi", "Texas")


# ============================================================
# SECTION 3 — SUB-REGION FILTERING AND GRAPH BUILDING
# ============================================================

# 3.1 Filter any data frame by a Tags column entry
filter_by_tag <- function(df, tag_name) {
  df %>% filter(!is.na(Tags), str_detect(Tags, fixed(tag_name)))
}

# 3.2 Build a deduplicated edge list for one sub-region.
#     The Kumu master file may contain the same (From, To) pair
#     more than once within a sub-region when different individual
#     workshops inside that region gave it different strengths.
#     We average those strengths — exactly as build_wmat() does
#     in the Network_Analysis_*.R scripts — to obtain one edge weight
#     per (From, To) pair.
build_subregion_edges <- function(all_connections, tag_name) {
  filter_by_tag(all_connections, tag_name) %>%
    group_by(From, To) %>%
    summarise(
      avg_strength = mean(strength_num, na.rm = TRUE),
      polarity     = sign(mean(strength_num, na.rm = TRUE)),
      n_rows       = n(),          # how many raw rows contributed
      .groups      = "drop"
    ) %>%
    filter(!is.na(avg_strength))
}

# 3.3 Build an igraph directed weighted graph for one sub-region.
#     Nodes = elements tagged to this sub-region (from Elements sheet)
#     plus any additional nodes that appear only in edges.
build_subregion_graph <- function(all_elements, all_connections, tag_name) {
  edges       <- build_subregion_edges(all_connections, tag_name)
  tagged_nds  <- filter_by_tag(all_elements, tag_name)$Label
  edge_nds    <- unique(c(edges$From, edges$To))
  all_nds     <- unique(c(tagged_nds, edge_nds))
  
  if (nrow(edges) == 0) {
    warning(sprintf("No edges for sub-region: %s", tag_name))
    return(NULL)
  }
  
  graph_from_data_frame(
    d        = data.frame(from     = edges$From,
                          to       = edges$To,
                          weight   = edges$avg_strength,
                          polarity = edges$polarity),
    vertices = data.frame(name = all_nds),
    directed = TRUE
  )
}

# Build all sub-region graphs
cat("\n=== Building sub-region graphs ===\n")
au_graphs <- setNames(
  lapply(AU_SUBREGIONS, build_subregion_graph,
         all_elements = au_elements, all_connections = au_connections),
  AU_SUBREGIONS)

gc_graphs <- setNames(
  lapply(GC_SUBREGIONS, build_subregion_graph,
         all_elements = al_elements, all_connections = al_connections),
  GC_SUBREGIONS)

for (r in AU_SUBREGIONS) {
  g <- au_graphs[[r]]
  if (!is.null(g))
    cat(sprintf("  AU | %-22s N=%d  E=%d\n", r, vcount(g), ecount(g)))
}
for (r in GC_SUBREGIONS) {
  g <- gc_graphs[[r]]
  if (!is.null(g))
    cat(sprintf("  GC | %-22s N=%d  E=%d\n", r, vcount(g), ecount(g)))
}


# ============================================================
# SECTION 4 — HELPER FUNCTIONS (metric computation)
# ============================================================
# Self-contained copies of the Goal 1 helper functions so this
# script can run independently.

# 4.1 Weighted indegree / outdegree / degree
#     Indegree_i  = Σ_j |w_ji|   (influence RECEIVED by i)
#     Outdegree_i = Σ_j |w_ij|   (influence EXERTED by i)
#     Degree_i    = Indegree_i + Outdegree_i
compute_degrees <- function(g) {
  edge_df <- igraph::as_data_frame(g, what = "edges")
  in_d  <- edge_df %>% group_by(to)   %>%
    summarise(indegree  = sum(abs(weight)), .groups = "drop") %>%
    rename(name = to)
  out_d <- edge_df %>% group_by(from) %>%
    summarise(outdegree = sum(abs(weight)), .groups = "drop") %>%
    rename(name = from)
  data.frame(name = V(g)$name) %>%
    left_join(in_d,  by = "name") %>%
    left_join(out_d, by = "name") %>%
    mutate(indegree  = replace_na(indegree,  0),
           outdegree = replace_na(outdegree, 0),
           degree    = indegree + outdegree)
}

# 4.2 Concept type classification
#     Transmitter: indegree=0, outdegree>0  (pure driver)
#     Receiver:    indegree>0, outdegree=0  (pure outcome)
#     Ordinary:    both > 0                 (mediator)
#     Isolated:    both = 0
classify_concepts <- function(node_df) {
  node_df %>%
    mutate(concept_type = case_when(
      indegree == 0 & outdegree  > 0 ~ "Transmitter",
      indegree  > 0 & outdegree == 0 ~ "Receiver",
      indegree  > 0 & outdegree  > 0 ~ "Ordinary",
      TRUE                           ~ "Isolated"))
}

# 4.3 Simple-cycle finder (depth-first search, Johnson 1975)
#     max_length = 7 caps computation time on large networks.
find_simple_cycles <- function(g, max_length = 7) {
  n <- vcount(g)
  if (n < 2) return(list())
  vnames <- V(g)$name
  adj    <- igraph::as_adj_list(g, mode = "out")
  all_cycles <- list()
  dfs <- function(s, curr, path) {
    if (length(path) > max_length) return()
    for (nb in as.integer(adj[[curr]])) {
      if (nb == s && length(path) >= 2) {
        all_cycles[[length(all_cycles) + 1]] <<- vnames[c(path, s)]
      } else if (nb > s && !(nb %in% path)) {
        dfs(s, nb, c(path, nb))
      }
    }
  }
  for (s in seq_len(n)) dfs(s, s, s)
  all_cycles
}

# 4.4 Classify a single feedback loop as Reinforcing (+1) or Balancing (-1)
#     Product of all edge polarities around the loop.
classify_loop_polarity <- function(g, cycle) {
  pol <- 1
  for (k in seq_len(length(cycle) - 1)) {
    eid <- get.edge.ids(g, c(cycle[k], cycle[k + 1]))
    if (eid == 0) return(NA)
    pol <- pol * sign(E(g)$weight[eid])
  }
  pol
}

# 4.5 Prominence as in Hoffman et al., 2014
compute_prominence <- function(g, node_df) {
  # Eigenvector centrality with tie weights (Hoffman et al., 2014)
  ec <- eigen_centrality(g, directed = TRUE, weights = abs(E(g)$weight))$vector
  
  # Min-max normalize eigenvector centrality
  mx <- max(ec, na.rm = TRUE)
  mn <- min(ec, na.rm = TRUE)
  
  node_df$eigen_centrality <- ec[node_df$name]
  node_df <- node_df %>%
    mutate(
      norm_eigen   = ifelse(mx == mn, 1, (eigen_centrality - mn) / (mx - mn)),
      occurrence   = 1,
      prominence   = (occurrence + norm_eigen) / 2
    )
  return(node_df)
}

# ============================================================
# SECTION 5 — COMPUTE ALL METRICS PER SUB-REGION
# ============================================================

compute_all_metrics <- function(g, model_name, verbose = TRUE) {
  if (is.null(g)) return(NULL)
  
  N       <- vcount(g)
  E_count <- ecount(g)
  density <- if (N > 1) E_count / (N * (N - 1)) else 0
  
  # --- Concept types & R/T ratio ---
  node_df <- compute_degrees(g) %>% classify_concepts()
  n_tx  <- sum(node_df$concept_type == "Transmitter")
  n_rx  <- sum(node_df$concept_type == "Receiver")
  n_ord <- sum(node_df$concept_type == "Ordinary")
  n_iso <- sum(node_df$concept_type == "Isolated")
  RT    <- if (n_tx > 0) n_rx / n_tx else NA_real_
  
  # --- Centrality metrics ---
  # Use inverse absolute weight as distance: stronger links = shorter path
  E(g)$dist_weight <- 1 / (abs(E(g)$weight) + 1e-6)
  btw <- betweenness(g,     weights = E(g)$dist_weight, normalized = TRUE)
  clo <- closeness(g,       weights = E(g)$dist_weight, normalized = TRUE,
                   mode = "out")
  eig <- eigen_centrality(g, weights = abs(E(g)$weight),
                          directed = TRUE)$vector
  node_df$betweenness <- btw
  node_df$closeness   <- clo
  node_df$eigenvector <- eig
  node_df <- compute_prominence(g, node_df)
  
  # --- Link polarity ---
  edge_df <- igraph::as_data_frame(g, what = "edges")
  n_pos   <- sum(edge_df$polarity > 0, na.rm = TRUE)
  n_neg   <- sum(edge_df$polarity < 0, na.rm = TRUE)
  
  # --- Path metrics (unweighted topology) ---
  dist_mat  <- distances(g, weights = NA)
  fin_d     <- dist_mat[is.finite(dist_mat) & dist_mat > 0]
  apl  <- if (length(fin_d) > 0) mean(fin_d)                     else NA_real_
  diam <- if (length(fin_d) > 0) max(dist_mat[is.finite(dist_mat)]) else NA_real_
  
  # --- Clustering coefficients ---
  clust_avg <- transitivity(g, type = "average", isolates = "zero")
  clust_glo <- transitivity(g, type = "global")
  
  # --- Feedback loops ---
  # Circuit rank: CR = E - N + C   (minimum independent loops)
  n_comp    <- components(g, mode = "weak")$no
  circ_rank <- E_count - N + n_comp
  
  # Enumerate and classify simple cycles (≤ 7 steps)
  cycles  <- list()
  n_reinf <- 0; n_bal <- 0
  if (N >= 2 && E_count >= 2) {
    cycles <- find_simple_cycles(g, max_length = 7)
    if (length(cycles) > 0) {
      pols    <- sapply(cycles, classify_loop_polarity, g = g)
      n_reinf <- sum(pols >  0, na.rm = TRUE)
      n_bal   <- sum(pols <  0, na.rm = TRUE)
    }
  }
  
  if (verbose) {
    cat(sprintf("\n── %s ──\n", model_name))
    cat(sprintf("  N=%d  E=%d  D=%.4f  R/T=%s\n",
                N, E_count, density,
                if (is.na(RT)) "undef" else sprintf("%.3f", RT)))
    cat(sprintf("  TX=%d  RX=%d  ORD=%d  ISO=%d\n",
                n_tx, n_rx, n_ord, n_iso))
    cat(sprintf("  APL=%s  Diam=%s  Clust_avg=%.3f  CR=%d\n",
                if (is.na(apl))  "—" else sprintf("%.3f", apl),
                if (is.na(diam)) "—" else as.character(diam),
                clust_avg, circ_rank))
    cat(sprintf("  Cycles≤7=%d  R-loops=%d  B-loops=%d\n",
                length(cycles), n_reinf, n_bal))
    cat(sprintf("  Pos=%d (%.0f%%)  Neg=%d (%.0f%%)\n",
                n_pos, 100*n_pos/max(E_count,1),
                n_neg, 100*n_neg/max(E_count,1)))
    
    # Top-5 concepts by prominence
    top5 <- node_df %>% arrange(desc(prominence)) %>%
      head(5) %>% pull(name)
    cat(sprintf("  Top prominence: %s\n", paste(top5, collapse = " | ")))
  }
  
  invisible(list(
    model_name    = model_name,
    N = N, E = E_count, density = density,
    n_tx = n_tx, n_rx = n_rx, n_ord = n_ord, n_iso = n_iso,
    RT_ratio      = RT,
    n_pos = n_pos, n_neg = n_neg,
    apl = apl, diam = diam,
    clust_avg = clust_avg, clust_glo = clust_glo,
    circ_rank = circ_rank,
    n_cycles      = length(cycles),
    n_reinforcing = n_reinf, n_balancing = n_bal,
    node_df       = node_df,
    cycles        = cycles,
    g             = g
  ))
}

cat("\n\n========== AUSTRALIA — SUB-REGION METRICS ==========\n")
au_metrics <- setNames(
  lapply(AU_SUBREGIONS, function(r)
    compute_all_metrics(au_graphs[[r]], r)),
  AU_SUBREGIONS)

cat("\n\n========== GULF COAST — SUB-REGION METRICS ==========\n")
gc_metrics <- setNames(
  lapply(GC_SUBREGIONS, function(r)
    compute_all_metrics(gc_graphs[[r]], r)),
  GC_SUBREGIONS)



# ============================================================
# SECTION 6b — GALVESTON (NEW WORKSHOP, SPLIT BY FISHERY)
# ============================================================
# Galveston is added to Goal 2 as a third workshop, divided into its three
# fishery groups exactly as Australia is divided into sub-regions:
#   Recreational · Charter · Commercial
# The fishery is recorded in the Tags column (pipe-separated), so the same
# filter_by_tag() → build_subregion_graph() → compute_all_metrics() pipeline
# used for AU and GC applies unchanged.
#
# Galveston is NOT added to au_metrics / gc_metrics and is NOT plotted, so the
# figures below still show only Australia and Gulf Coast. It DOES appear in the
# metrics comparison table (Section 10) via gv_tbl.
#
# Strength scale: integer ±1/±2 (same as AU/GC) — no rescaling applied.
# >>> SET THE FILE PATH BELOW to the Galveston Kumu export before running. <<<

GV_FILE <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Exported_Kumu_Galveston_30April.xlsx"  # <-- EDIT THIS PATH

# Galveston "sub-regions" are the three fishery groups (Tags column values)
GV_SUBREGIONS <- c("Recreational", "Charter", "Commercial")

gv_elements    <- read_excel(GV_FILE, sheet = "Elements")
gv_connections <- read_excel(GV_FILE, sheet = "Connections") %>%
  mutate(strength_num = suppressWarnings(as.numeric(Strength))) %>%
  filter(!is.na(From), !is.na(To), !is.na(strength_num))

cat("\nLoaded — GV:", nrow(gv_elements), "elements |",
    nrow(gv_connections), "connections\n")

# Strength-scale check: expect integer ±1/±2 (not the ±0.5/±1.0 MM scale)
gv_strength_vals <- sort(unique(gv_connections$strength_num))
cat("Galveston unique Strength values:",
    paste(gv_strength_vals, collapse = ", "), "\n")
if (any(abs(gv_strength_vals) > 0 & abs(gv_strength_vals) < 1))
  warning("Galveston Strength contains |value| < 1 (looks like the ±0.5/±1.0 ",
          "Mental Modeler scale). Multiply by 2 to match the integer ±1/±2 ",
          "scale used for the other models before trusting weighted metrics.")

# Build one graph per fishery — identical machinery to the AU/GC sub-regions
cat("\n=== Building Galveston fishery graphs ===\n")
gv_graphs <- setNames(
  lapply(GV_SUBREGIONS, build_subregion_graph,
         all_elements = gv_elements, all_connections = gv_connections),
  GV_SUBREGIONS)

for (r in GV_SUBREGIONS) {
  g <- gv_graphs[[r]]
  if (!is.null(g))
    cat(sprintf("  GV | %-22s N=%d  E=%d\n", r, vcount(g), ecount(g)))
}

cat("\n\n========== GALVESTON — FISHERY METRICS ==========\n")
gv_metrics <- setNames(
  lapply(GV_SUBREGIONS, function(r)
    compute_all_metrics(gv_graphs[[r]], r)),
  GV_SUBREGIONS)

# ============================================================
# SECTION 7 — CONFLICT / DISAGREEMENT ANALYSIS
# ============================================================
# A CONFLICT is defined as: two sub-regions share the same From→To
# connection but assign it OPPOSITE polarity (one positive, one negative).
# These represent genuine perceptual disagreements between groups and are
# of particular interest for management, as the same intervention could
# have opposite predicted effects depending on the sub-region's worldview.
#
# Note: the Data Standardisation Methodology retained such conflicting-
# strength pairs as separate rows in Kumu. This function detects them
# computationally across all sub-region pairs.

find_conflicts <- function(all_connections, subregions) {
  pairs <- combn(subregions, 2, simplify = FALSE)
  results <- lapply(pairs, function(pair) {
    r1 <- pair[1]; r2 <- pair[2]
    
    agg_pol <- function(r) {
      filter_by_tag(all_connections, r) %>%
        mutate(pair_key = paste(From, To, sep = "→"),
               pol      = sign(strength_num)) %>%
        filter(!is.na(pol)) %>%
        group_by(pair_key) %>%
        summarise(pol = sign(mean(pol)), .groups = "drop")
    }
    
    e1 <- agg_pol(r1); e2 <- agg_pol(r2)
    merged <- inner_join(e1, e2, by = "pair_key", suffix = c("_A","_B")) %>%
      filter(pol_A != pol_B)
    # Return NULL when no conflicts to avoid bind_rows type-mismatch on empty frames
    if (nrow(merged) == 0) return(NULL)
    merged %>%
      mutate(region_A  = r1, region_B  = r2,
             polarity_A = dplyr::if_else(pol_A > 0, "Positive", "Negative"),
             polarity_B = dplyr::if_else(pol_B > 0, "Positive", "Negative"))
  })
  results_nonnull <- Filter(Negate(is.null), results)
  if (length(results_nonnull) == 0)
    return(data.frame(region_A=character(), region_B=character(),
                      pair_key=character(), polarity_A=character(),
                      polarity_B=character()))
  bind_rows(results_nonnull)
}

cat("\n\n========== CONFLICTS — AUSTRALIA ==========\n")
au_conflicts <- find_conflicts(au_connections, AU_SUBREGIONS)
if (nrow(au_conflicts) > 0) {
  print(au_conflicts %>%
          select(region_A, region_B, pair_key, polarity_A, polarity_B),
        row.names = FALSE)
} else {
  cat("  No polarity conflicts detected across Australian sub-regions.\n")
}

cat("\n\n========== CONFLICTS — GULF COAST ==========\n")
gc_conflicts <- find_conflicts(al_connections, GC_SUBREGIONS)
if (nrow(gc_conflicts) > 0) {
  print(gc_conflicts %>%
          select(region_A, region_B, pair_key, polarity_A, polarity_B),
        row.names = FALSE)
} else {
  cat("  No polarity conflicts detected across Gulf Coast sub-regions.\n")
}


# ============================================================
# SECTION 8 — EDR FOR ALL PAIRWISE SUB-REGION COMPARISONS
# ============================================================
# EDR (Element Distance Ratio) from Schaffernicht & Groesser (2011) §3.3.
# Applied here between sub-regions WITHIN the same workshop:
#   • Australia: 6 pairs from 4 sub-regions  (4 choose 2)
#   • Gulf Coast: 10 pairs from 5 sub-regions (5 choose 2)
#
# No concept mapping (concept_map = NULL) is required within each
# workshop because terminology was already standardised across sub-regions
# during the data preparation pipeline (see Data Standardisation Methodology).
#
# Parameters (same as Goal 1):
#   b = 2 (maximum absolute strength on ±1/±2 Kumu scale)
#   e = 2 (two polarities), d = 0, a = 1, c = 2

compute_EDR <- function(metrics_A, metrics_B, b = 2, concept_map = NULL) {
  g_A    <- metrics_A$g
  g_B    <- metrics_B$g
  vars_A <- V(g_A)$name
  vars_B <- V(g_B)$name
  
  if (!is.null(concept_map)) {
    vars_A <- ifelse(vars_A %in% names(concept_map),
                     concept_map[vars_A], vars_A)
  }
  
  V_c  <- intersect(vars_A, vars_B)
  V_uA <- setdiff(vars_A, V_c)
  V_uB <- setdiff(vars_B, V_c)
  v_c  <- length(V_c); v_uA <- length(V_uA); v_uB <- length(V_uB)
  v    <- v_c + v_uA + v_uB
  all_vars <- c(V_c, V_uA, V_uB)
  
  # Build signed adjacency matrices over the full union variable space
  build_adj <- function(g, orig_vars, mapped_vars) {
    mat <- matrix(0, v, v, dimnames = list(all_vars, all_vars))
    edf <- igraph::as_data_frame(g, what = "edges")
    for (k in seq_len(nrow(edf))) {
      fr_orig <- edf$from[k]; to_orig <- edf$to[k]
      fr <- if (!is.null(concept_map) && fr_orig %in% names(concept_map))
        concept_map[fr_orig] else fr_orig
      to <- if (!is.null(concept_map) && to_orig %in% names(concept_map))
        concept_map[to_orig] else to_orig
      if (fr %in% all_vars && to %in% all_vars)
        mat[fr, to] <- edf$weight[k]
    }
    mat
  }
  
  mat_A <- build_adj(g_A, V(g_A)$name, vars_A)
  mat_B <- build_adj(g_B, V(g_B)$name, vars_B)
  
  # diff(i,j) = |a_ij − b_ij|  (Eq. 2 in the paper)
  diff_mat <- abs(mat_A - mat_B)
  diag(diff_mat) <- 0
  numerator <- sum(diff_mat)
  
  # Denominator = 4 × v_c(v_c−1) + 2 × [v(v−1) − v_c(v_c−1)]
  denom_cc    <- 4 * v_c * max(v_c - 1, 0)
  denom_other <- 2 * (v * max(v - 1, 0) - v_c * max(v_c - 1, 0))
  denominator <- denom_cc + denom_other
  EDR <- if (denominator == 0) 0 else numerator / denominator
  
  # Common links (From→To pairs present in BOTH, restricted to V_c)
  ea <- igraph::as_data_frame(g_A, "edges") %>%
    filter(from %in% V_c, to %in% V_c) %>%
    mutate(pair = paste(from, to, sep="→"))
  eb <- igraph::as_data_frame(g_B, "edges") %>%
    filter(from %in% V_c, to %in% V_c) %>%
    mutate(pair = paste(from, to, sep="→"))
  common_links <- intersect(ea$pair, eb$pair)
  
  list(EDR = EDR, numerator = numerator, denominator = denominator,
       v_c = v_c, v_uA = v_uA, v_uB = v_uB, v = v,
       V_c = V_c, V_uA = V_uA, V_uB = V_uB,
       n_common_links = length(common_links),
       common_links   = common_links,
       mat_A = mat_A, mat_B = mat_B)
}

# Run all pairwise EDR comparisons for one workshop
run_pairwise_EDR <- function(metrics_list, label) {
  valid_names <- names(Filter(Negate(is.null), metrics_list))
  pairs       <- combn(valid_names, 2, simplify = FALSE)
  cat(sprintf("\n\n========== EDR — %s (%d pairs) ==========\n",
              label, length(pairs)))
  lapply(pairs, function(pair) {
    r1 <- pair[1]; r2 <- pair[2]
    res <- tryCatch(
      compute_EDR(metrics_list[[r1]], metrics_list[[r2]]),
      error = function(e) { cat("  ERROR:", conditionMessage(e), "\n"); NULL })
    if (!is.null(res))
      cat(sprintf("  %-22s vs %-22s | EDR=%.4f  Vc=%d  CommonLinks=%d\n",
                  r1, r2, res$EDR, res$v_c, res$n_common_links))
    list(r1 = r1, r2 = r2, result = res)
  })
}

au_edr_list <- run_pairwise_EDR(au_metrics, "AUSTRALIA")
gc_edr_list <- run_pairwise_EDR(gc_metrics, "GULF COAST")


# ============================================================
# SECTION 9 — SILS + LDR + MDR FOR ALL PAIRWISE COMPARISONS
# ============================================================
# Following Schaffernicht & Groesser (2011) §3.4–3.5.
# Steps:
#   (a) SILS approximation: shortest independent loop set
#   (b) Semantic loop matching via Jaccard similarity (threshold 0.4)
#   (c) LDR(m,n) = g·ldd + i·lpold + j·EDR_loop   (equal weights 1/3)
#   (d) MDR = mean of all LDR values (matched + unmatched loops get LDR=1)

get_sils <- function(cycles) {
  if (length(cycles) == 0) return(list())
  cycles_sorted  <- cycles[order(sapply(cycles, length))]
  all_cycle_vars <- unique(unlist(lapply(cycles, function(cyc) cyc[-length(cyc)])))
  covered <- character(0); sils <- list()
  for (cyc in cycles_sorted) {
    cyc_vars <- cyc[-length(cyc)]
    if (any(!(cyc_vars %in% covered))) {
      sils    <- c(sils, list(cyc))
      covered <- unique(c(covered, cyc_vars))
    }
    if (all(all_cycle_vars %in% covered)) break
  }
  sils
}

compute_LDR_MDR <- function(mA, mB, edr_res, w_g=1/3, w_i=1/3, w_j=1/3) {
  V_c      <- edr_res$V_c
  sils_A   <- get_sils(mA$cycles)
  sils_B   <- get_sils(mB$cycles)
  nA <- length(sils_A); nB <- length(sils_B)
  
  # Edge case: no loops in either model
  if (nA == 0 && nB == 0)
    return(list(MDR = 0, ldr_values = numeric(0),
                n_matched = 0, n_unmatched = 0))
  
  # All loops unique when one model has none
  if (nA == 0 || nB == 0) {
    n_un <- nA + nB
    return(list(MDR = 1, ldr_values = rep(1, n_un),
                n_matched = 0, n_unmatched = n_un))
  }
  
  # Loop variable sets (restricted to V_c for cross-model matching)
  lv_A <- lapply(sils_A, function(cyc) intersect(cyc[-length(cyc)], V_c))
  lv_B <- lapply(sils_B, function(cyc) intersect(cyc[-length(cyc)], V_c))
  
  # Jaccard matrix for loop pairing
  jac <- matrix(0, nA, nB)
  for (ia in seq_len(nA)) for (ib in seq_len(nB)) {
    i <- length(intersect(lv_A[[ia]], lv_B[[ib]]))
    u <- length(union(lv_A[[ia]],     lv_B[[ib]]))
    jac[ia, ib] <- if (u == 0) 0 else i / u
  }
  
  # Greedy matching (threshold 0.4)
  mA_idx <- integer(0); mB_idx <- integer(0); pairs <- list()
  jc <- jac
  repeat {
    bv <- max(jc, na.rm = TRUE)
    if (bv < 0.4) break
    bp <- which(jc == bv, arr.ind = TRUE)[1, ]
    pairs  <- c(pairs, list(bp))
    mA_idx <- c(mA_idx, bp[1]); mB_idx <- c(mB_idx, bp[2])
    jc[bp[1], ] <- -Inf; jc[, bp[2]] <- -Inf
  }
  
  # LDR for each matched pair
  ldr_vals <- sapply(pairs, function(bp) {
    ia <- bp[1]; ib <- bp[2]
    pol_A <- classify_loop_polarity(mA$g, sils_A[[ia]])
    pol_B <- classify_loop_polarity(mB$g, sils_B[[ib]])
    lpold <- if (is.na(pol_A) | is.na(pol_B)) 0 else as.numeric(pol_A != pol_B)
    ldd   <- 0  # no delay information in the data
    
    # Sub-EDR for the matched loop pair
    lnodes_A <- sils_A[[ia]][-length(sils_A[[ia]])]
    lnodes_B <- sils_B[[ib]][-length(sils_B[[ib]])]
    sg_A <- induced_subgraph(mA$g, lnodes_A[lnodes_A %in% V(mA$g)$name])
    sg_B <- induced_subgraph(mB$g, lnodes_B[lnodes_B %in% V(mB$g)$name])
    edr_loop <- tryCatch(
      compute_EDR(list(g = sg_A), list(g = sg_B))$EDR,
      error = function(e) 1)
    
    w_g * ldd + w_i * lpold + w_j * edr_loop
  })
  
  # Unmatched loops → LDR = 1 (maximum distance)
  n_un     <- length(setdiff(seq_len(nA), mA_idx)) +
    length(setdiff(seq_len(nB), mB_idx))
  ldr_all  <- c(ldr_vals, rep(1, n_un))
  MDR      <- if (length(ldr_all) > 0) mean(ldr_all) else NA_real_
  
  list(MDR = MDR, ldr_values = ldr_all,
       n_matched = length(pairs), n_unmatched = n_un)
}

# Run all pairwise LDR/MDR for one workshop
run_pairwise_LDR_MDR <- function(metrics_list, edr_list, label) {
  cat(sprintf("\n\n========== LDR / MDR — %s ==========\n", label))
  lapply(seq_along(edr_list), function(k) {
    r1  <- edr_list[[k]]$r1; r2 <- edr_list[[k]]$r2
    edr <- edr_list[[k]]$result
    if (is.null(edr) || is.na(edr$EDR)) return(NULL)
    res <- tryCatch(
      compute_LDR_MDR(metrics_list[[r1]], metrics_list[[r2]], edr),
      error = function(e) { cat("  ERROR:", conditionMessage(e), "\n"); NULL })
    if (!is.null(res))
      cat(sprintf("  %-22s vs %-22s | EDR=%.4f  MDR=%s  matched=%d  unmatched=%d\n",
                  r1, r2, edr$EDR,
                  if (is.na(res$MDR)) "NA" else sprintf("%.4f", res$MDR),
                  res$n_matched, res$n_unmatched))
    list(r1=r1, r2=r2, EDR=edr$EDR,
         MDR=if(!is.null(res)) res$MDR else NA_real_,
         v_c=edr$v_c, n_common_links=edr$n_common_links)
  })
}

au_mdr_list <- run_pairwise_LDR_MDR(au_metrics, au_edr_list, "AUSTRALIA")
gc_mdr_list <- run_pairwise_LDR_MDR(gc_metrics, gc_edr_list, "GULF COAST")


# ============================================================
# SECTION 10 — SUMMARY TABLES
# ============================================================

# 10.1 Structural metrics table
make_metrics_table <- function(metrics_list, label) {
  cat(sprintf("\n\n========== METRICS TABLE — %s ==========\n", label))
  tbl <- bind_rows(lapply(names(metrics_list), function(r) {
    m <- metrics_list[[r]]
    if (is.null(m)) return(NULL)
    data.frame(
      SubRegion    = r,
      N            = m$N,
      E            = m$E,
      Density      = round(m$density,   4),
      Transmitters = m$n_tx,
      Receivers    = m$n_rx,
      Ordinary     = m$n_ord,
      RT_Ratio     = round(ifelse(is.na(m$RT_ratio), 0, m$RT_ratio), 3),
      Pos_pct      = round(100 * m$n_pos / max(m$E, 1), 1),
      Neg_pct      = round(100 * m$n_neg / max(m$E, 1), 1),
      APL          = round(ifelse(is.na(m$apl), 0, m$apl), 3),
      Diameter     = ifelse(is.na(m$diam), 0, m$diam),
      Clust_avg    = round(m$clust_avg, 3),
      Circ_Rank    = m$circ_rank,
      Cycles_le7   = m$n_cycles,
      R_loops      = m$n_reinforcing,
      B_loops      = m$n_balancing
    )
  }))
  print(tbl, row.names = FALSE)
  invisible(tbl)
}

au_tbl <- make_metrics_table(au_metrics, "AUSTRALIA SUB-REGIONS")
gc_tbl <- make_metrics_table(gc_metrics, "GULF COAST SUB-REGIONS")
gv_tbl <- make_metrics_table(gv_metrics, "GALVESTON FISHERIES")

# Combined comparison table across all workshops (Australia, Gulf Coast, Galveston)
comparison_table_g2 <- bind_rows(
  au_tbl %>% mutate(Workshop = "Australia"),
  gc_tbl %>% mutate(Workshop = "Gulf Coast"),
  gv_tbl %>% mutate(Workshop = "Galveston")
) %>%
  relocate(Workshop)
cat("\n\n========== COMBINED METRICS TABLE — ALL WORKSHOPS ==========\n")
print(comparison_table_g2, row.names = FALSE)


# ============================================================
# SECTION 11 — VISUALISATIONS
# ============================================================

# Helper: heatmap for square matrices


plot_heatmap <- function(mat, title, low_col = "white", high_col = "#0072B2",
                         out_file = NULL) {
  df <- as.data.frame(as.table(mat)) %>%
    setNames(c("Row","Col","Value")) %>%
    filter(!is.na(Value))
  p <- ggplot(df, aes(x = Col, y = Row, fill = Value)) +
    geom_tile(colour = "white", linewidth = 0.5) +
    geom_text(aes(label = round(Value, 2)), size = 3.5, colour = "black") +
    scale_fill_gradient(
      low = low_col, high = high_col,
      limits = c(0, 1), name = "Jaccard",
      guide = guide_colorbar(
        barwidth = 1.2,      # width of the color bar (in "lines" units)
        barheight = 8,       # height/length of the color bar
        title.vjust = 1      # nudges title relative to the bar
      )
    ) +
    labs(title = title, x = NULL, y = NULL) +
    theme_bw(base_size = 11) +
    theme(
      axis.text.x = element_text(angle = 0, hjust = 0.5),
      axis.text.y = element_text(angle = 90, hjust = 0.5),
      legend.title = element_text(margin = margin(b = 10))  # space below title, above the "1"
    )
  if (!is.null(out_file)) ggsave(out_file, p, width = 6, height = 5)
  p
}

# 11.1 Concept Jaccard heatmaps

p_au_cjac <- plot_heatmap(au_cjac, "Australia - Elements")
p_gc_cjac <- plot_heatmap(gc_cjac, "Gulf Coast - Elements")

png(filename = "Jaccard_Cncept_AU_subregions.png", 
    width = 7, height = 5, units = "in", # Set size in inches
    res = 1200)  
p_au_cjac
dev.off()

png(filename = "Jaccard_Cncept_GC_subregions.png", 
    width = 7, height = 5, units = "in", # Set size in inches
    res = 1200)  
p_gc_cjac
dev.off()


# 11.2 Link Jaccard heatmaps

p_au_ljac<-plot_heatmap(au_ljac, "Australia - Links")
p_gc_ljac<-plot_heatmap(gc_ljac, "Gulf Coast - Links")

png(filename = "Jaccard_Link_AU_subregions.png", 
    width = 7, height = 5, units = "in", # Set size in inches
    res = 1200)  
p_au_ljac
dev.off()

png(filename = "Jaccard_Link_GC_subregions.png", 
    width = 7, height = 5, units = "in", # Set size in inches
    res = 1200)  
p_gc_ljac
dev.off()


plot_heatmap(au_ljac, high_col = "#009E73",
             out_file = "Jaccard_Link_Australia.pdf")
plot_heatmap(gc_ljac, "Link Jaccard — Gulf Coast sub-regions",
             high_col = "#009E73",
             out_file = "Jaccard_Link_GulfCoast.pdf")

# 11.3 EDR heatmaps
plot_heatmap(au_dist$EDR, "EDR — Australia sub-regions",
             high_col = "#D55E00",
             out_file = "EDR_Australia.pdf")
plot_heatmap(gc_dist$EDR, "EDR — Gulf Coast sub-regions",
             high_col = "#D55E00",
             out_file = "EDR_GulfCoast.pdf")

# 11.4 MDR heatmaps
plot_heatmap(au_dist$MDR, "MDR — Australia sub-regions",
             high_col = "#CC79A7",
             out_file = "MDR_Australia.pdf")
plot_heatmap(gc_dist$MDR, "MDR — Gulf Coast sub-regions",
             high_col = "#CC79A7",
             out_file = "MDR_GulfCoast.pdf")

# 11.5 Structural metrics bar charts (faceted)
plot_structural <- function(tbl, label) {
  tbl %>%
    select(SubRegion, N, E, Transmitters, Receivers,
           Ordinary, R_loops, B_loops) %>%
    pivot_longer(-SubRegion, names_to = "Metric", values_to = "Value") %>%
    ggplot(aes(x = SubRegion, y = Value, fill = SubRegion)) +
    geom_col(show.legend = FALSE) +
    facet_wrap(~Metric, scales = "free_y") +
    labs(title = paste("Structural metrics —", label), x = NULL, y = NULL) +
    theme_bw(base_size = 10) +
    theme(axis.text.x = element_text(angle = 35, hjust = 1))
}


ggsave("Structural_AU_subregions.pdf",
       plot_structural(au_tbl, "Australia sub-regions"),
       width = 12, height = 8)
ggsave("Structural_GC_subregions.pdf",
       plot_structural(gc_tbl, "Gulf Coast sub-regions"),
       width = 12, height = 8)

# 11.6 Top-10 concepts by prominence per sub-region — ranked dot (lollipop) chart
# Colored by concept group (Ecological, Fisheries, Human Dimensions, etc.)
library(ggnewscale)
GROUP_COLOURS <- c(
  "Central Concept"                   = "#FBD73F",
  "Ecological & Biological Factors"   = "#9AD354",
  "Human Dimensions"                  = "#F8895E",
  "Fisheries Operations & Practices"  = "#E382BA",
  "Fisheries Research & Management"   = "#8695C2",
  "Policy & Economics"                = "#5EB99B",
  "Other"                             = "grey60"
)

# Fixed display order for groups (top of plot -> bottom, since axis is reversed below)
GROUP_ORDER <- c(
  "Central Concept",
  "Ecological & Biological Factors",
  "Human Dimensions",
  "Fisheries Operations & Practices",
  "Fisheries Research & Management",
  "Policy & Economics",
  "Other"
)

plot_prominence <- function(metrics_list, label, groups_df = NULL, top_n = 10) {
  
  group_var <- if (!is.null(groups_df)) "Group" else "concept_type"
  
  plot_data <- bind_rows(lapply(names(metrics_list), function(r) {
    m <- metrics_list[[r]]
    if (is.null(m)) return(NULL)
    df <- m$node_df %>% arrange(desc(prominence)) %>% head(top_n) %>%
      mutate(subregion = r)
    if (!is.null(groups_df))
      df <- df %>% left_join(groups_df, by = "name") %>%
      mutate(Group = replace_na(Group, "Other"))
    else
      df <- df %>% mutate(Group = concept_type)
    df
  }))
  
  group_colours_used <- if (!is.null(groups_df)) GROUP_COLOURS else
    c("Transmitter" = "#E69F00", "Receiver" = "#56B4E9",
      "Ordinary"    = "#009E73", "Isolated"  = "grey60")
  
  # Restrict the fixed order to whatever groups are actually present (Other always sticks around)
  group_levels_present <- GROUP_ORDER[GROUP_ORDER %in% unique(plot_data$Group)]
  
  # One Group value per concept (in case a concept appears with the same Group across subregions)
  concept_groups <- plot_data %>%
    group_by(name) %>%
    summarise(Group     = first(.data[[group_var]]),
              mean_prom = mean(prominence, na.rm = TRUE),
              .groups   = "drop") %>%
    mutate(Group = factor(Group, levels = group_levels_present)) %>%
    arrange(Group, desc(mean_prom))
  
  # Final concept order: grouped first (Central Concept -> ... -> Policy & Economics -> Other),
  # ranked by mean prominence within each group
  concept_order <- concept_groups$name
  
  plot_data <- plot_data %>%
    mutate(name = factor(name, levels = rev(concept_order)))
  
  # Colour for each y-axis label, in the same order as the factor levels (rev(concept_order))
  label_colours <- group_colours_used[
    concept_groups$Group[match(rev(concept_order), concept_groups$name)]
  ]
  label_colours[is.na(label_colours)] <- "grey20"
  
  # Group lookup per concept, used only to drive the invisible legend-generating layer below
  group_lookup <- setNames(as.character(concept_groups$Group), concept_groups$name)
  
  plot_data <- plot_data %>%
    mutate(Group = factor(group_lookup[as.character(name)], levels = group_levels_present))
  
  ggplot(plot_data, aes(x = subregion, y = name)) +
    
    geom_tile(aes(fill = prominence), colour = "white", linewidth = 0.8) +
    geom_text(aes(label = round(prominence, 2)),
              size = 2.5, colour = "grey20") +
    scale_fill_gradientn(
      colours  = c("#f7fbff", "#9ecae1", "#2171b5"),
      na.value = "grey92",
      name     = "Prominence",
      guide    = guide_colourbar(position = "right")   # keep Prominence legend on the right
    ) +
    
    # --- Switch scales so Group gets its own legend, independent of the Prominence fill ---
    new_scale_fill() +
    
    # Invisible layer: exists only so the Group colours get a legend at the bottom.
    # size = 0 / alpha = 0 keeps it from showing up on the tiles themselves.
    geom_point(aes(fill = Group), shape = 22, size = 0, alpha = 0, stroke = 0) +
    scale_fill_manual(
      values = group_colours_used, name = NULL, drop = FALSE,
      guide  = guide_legend(position = "bottom", override.aes = list(size = 5, alpha = 1))
    ) +
    
    scale_x_discrete(expand = expansion(add = c(0.6, 0.3))) +
    labs(title = paste("Top", top_n, "concepts by Prominence —", label),
         x = NULL, y = NULL) +
    theme_minimal(base_size = 9) +
    theme(
      axis.text.x       = element_text(angle = 0, hjust = 0.5),
      axis.text.y       = element_text(size = 8, face = "bold", colour = label_colours),
      panel.grid        = element_blank(),
      legend.title      = element_text(size = 8),
      legend.text       = element_text(size = 7),
      plot.title        = element_text(size = 10, face = "bold")
    )
}

prom_Aus<-plot_prominence(au_metrics, "Australia",  groups_df = au_groups)
prom_Gc<-plot_prominence(gc_metrics, "Gulf Coast", groups_df = gc_groups)

ggsave("Prominence_AU_subregions.pdf",
       plot_prominence(au_metrics, "Australia",  groups_df = au_groups),
       width = 10, height = 8)
ggsave("Prominence_GC_subregions.pdf",
       plot_prominence(gc_metrics, "Gulf Coast", groups_df = gc_groups),
       width = 10, height = 8)

png(filename = "Prominence_AU_subregions.png", 
    width = 10, height = 8, units = "in", # Set size in inches
    res = 1200)  
prom_Aus
dev.off()

png(filename = "Prominence_GC_subregions.png", 
    width = 10, height = 8, units = "in", # Set size in inches
    res = 1200)  
prom_Gc
dev.off()

# 11.7 Conflicts table plot (if any conflicts exist)
if (nrow(au_conflicts) > 0) {
  p_conf_au <- au_conflicts %>%
    select(region_A, region_B, pair_key, polarity_A, polarity_B) %>%
    ggplot(aes(x = pair_key, y = paste(region_A, "vs", region_B))) +
    geom_tile(aes(fill = polarity_A), colour = "white") +
    geom_text(aes(label = paste(polarity_A, "/", polarity_B)), size = 2.8) +
    labs(title = "Polarity conflicts — Australia",
         x = "Connection (From→To)", y = NULL, fill = "Polarity in A") +
    theme_bw(base_size = 9) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
  ggsave("Conflicts_Australia.pdf", p_conf_au, width = 12, height = 5)
}

if (nrow(gc_conflicts) > 0) {
  p_conf_gc <- gc_conflicts %>%
    select(region_A, region_B, pair_key, polarity_A, polarity_B) %>%
    ggplot(aes(x = pair_key, y = paste(region_A, "vs", region_B))) +
    geom_tile(aes(fill = polarity_A), colour = "white") +
    geom_text(aes(label = paste(polarity_A, "/", polarity_B)), size = 2.8) +
    labs(title = "Polarity conflicts — Gulf Coast",
         x = "Connection (From→To)", y = NULL, fill = "Polarity in A") +
    theme_bw(base_size = 9) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
  ggsave("Conflicts_GulfCoast.pdf", p_conf_gc, width = 12, height = 5)
}

cat("\n\nGoal 2 analysis complete. All outputs saved.\n")
