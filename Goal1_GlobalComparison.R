# ============================================================
#  MENTAL MODEL ANALYSIS — GOAL 1
#  Global comparison: Australia vs Gulf Coast USA (Alabama)
#  Schaffernicht & Groesser (2011) comparison method applied
# ============================================================
#
#  DATA SOURCES
#  - Exported_Kumu_Australia_7May.xlsx   → full Australian model
#  - Exported_Kumu_Alabama_30April.xlsx  → full Gulf Coast model
#    (Tags: Alabama, Florida, Louisiana, Mississippi, Texas)
#
#  KEY REFERENCE
#  Schaffernicht, M. & Groesser, S. N. (2011). A comprehensive method
#  for comparing mental models of dynamic systems.
#  European Journal of Operational Research, 210(1), 57–67.
#  → Section 3 of the paper is fully implemented here (EDR, LDR, MDR).
#
#  PAPER SECTION ↔ CODE SECTION MAP
#  Paper §3.1 (distance ratio approach)        → SECTION 4
#  Paper §3.2 (incorporating delays)           → noted in §3.9 below
#  Paper §3.3 (Element Distance Ratio, EDR)    → SECTION 5
#  Paper §3.4 (Loop Distance Ratio, LDR)       → SECTION 6
#  Paper §3.5 (Model Distance Ratio, MDR)      → SECTION 7
# ============================================================


# ============================================================
# SECTION 0 — INSTALL AND LOAD PACKAGES
# ============================================================

required_pkgs <- c("readxl", "igraph", "dplyr", "tidyr", "ggplot2",
                   "stringr", "knitr", "scales")

for (pkg in required_pkgs) {
  if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
}

library(readxl)
library(igraph)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)
library(scales)


# ============================================================
# SECTION 1 — DATA LOADING
# ============================================================
# Adjust the paths below to where your files are located.

au_elements    <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Australia/Kumu/Exported_Kumu_Australia_7May.xlsx",  sheet = "Elements")
au_connections <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Australia/Kumu/Exported_Kumu_Australia_7May.xlsx",  sheet = "Connections")

al_elements    <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Alabama/Kumu/Exported_Kumu_Alabama_30April.xlsx", sheet = "Elements")
al_connections <- read_excel("~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Alabama/Kumu/Exported_Kumu_Alabama_30April.xlsx", sheet = "Connections")

cat("Raw rows — Australia elements:", nrow(au_elements),
    " | Australia connections:", nrow(au_connections), "\n")
cat("Raw rows — Alabama elements:", nrow(al_elements),
    " | Alabama connections:", nrow(al_connections), "\n")


# ============================================================
# SECTION 2 — PREPROCESSING
# ============================================================
# The 'Influence Type' column records Positive / Negative polarity.
# The 'Strength' column records a numerical weight (e.g. 2, 1, -1, -2).
# Some From→To pairs appear on MULTIPLE rows because different sub-regions
# contributed the same link. We DEDUPLICATE by taking the mean strength
# across all rows sharing the same From→To pair (consistent with the
# METRICS document treatment).

# ---- 2.1 Parse polarity and numeric strength ----

parse_connections <- function(conn_df) {
  conn_df %>%
    mutate(
      # Normalise column names that differ slightly between files
      influence = coalesce(`Influence Type`, Direction),
      # Parse signed numeric strength
      strength_num = suppressWarnings(as.numeric(Strength))
    ) %>%
    filter(!is.na(From), !is.na(To), !is.na(strength_num))
}

au_conn <- parse_connections(au_connections)
al_conn <- parse_connections(al_connections)

# ---- 2.2 Deduplicate edges: average strength for repeated From→To pairs ----
# When the same From→To pair appears with different sub-region tags it gets
# multiple rows. We collapse to one edge per (From, To) pair and average
# the signed strength. The resulting polarity is the sign of the mean.

deduplicate_edges <- function(conn_df) {
  conn_df %>%
    group_by(From, To) %>%
    summarise(
      avg_strength = mean(strength_num, na.rm = TRUE),
      polarity     = sign(mean(strength_num, na.rm = TRUE)), # +1 or -1
      n_regions    = n(),
      .groups      = "drop"
    ) %>%
    filter(!is.na(avg_strength))
}

au_edges <- deduplicate_edges(au_conn)
al_edges <- deduplicate_edges(al_conn)

# Node lists (all unique concept labels from the Elements sheet)
au_nodes <- unique(au_elements$Label)
al_nodes <- unique(al_elements$Label)

cat("\nAfter deduplication:\n")
cat("  Australia — nodes:", length(au_nodes), " | edges:", nrow(au_edges), "\n")
cat("  Alabama   — nodes:", length(al_nodes), " | edges:", nrow(al_edges), "\n")


# ============================================================
# SECTION 3 — BUILD IGRAPH OBJECTS
# ============================================================
# We build weighted directed graphs.
# Edge weight = avg_strength (signed: positive or negative).
# Edge polarity = sign(avg_strength).

build_graph <- function(nodes, edges) {
  # Ensure all nodes from elements are included even if isolated
  all_nodes_in_edges <- unique(c(edges$From, edges$To))
  extra_nodes <- setdiff(nodes, all_nodes_in_edges)
  
  g <- graph_from_data_frame(
    d = data.frame(from     = edges$From,
                   to       = edges$To,
                   weight   = edges$avg_strength,
                   polarity = edges$polarity),
    vertices = data.frame(name = c(all_nodes_in_edges, extra_nodes)),
    directed = TRUE
  )
  g
}

g_au <- build_graph(au_nodes, au_edges)
g_al <- build_graph(al_nodes, al_edges)


# ============================================================
# SECTION 4 — HELPER FUNCTIONS
# ============================================================

# ---- 4.1 Weighted degree metrics ----
# Indegree of node i  = Σ |weight| of all incoming edges
# Outdegree of node i = Σ |weight| of all outgoing edges
# Degree centrality   = indegree + outdegree
#
# Using absolute values so both positive and negative links contribute
# to how embedded a concept is in the system.

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
      indegree  = replace_na(indegree,  0),
      outdegree = replace_na(outdegree, 0),
      degree    = indegree + outdegree
    )
}

# ---- 4.2 Concept type classification ----
# Transmitter: outdegree > 0, indegree = 0  — pure external driver
# Receiver:    indegree > 0, outdegree = 0  — pure terminal outcome
# Ordinary:    both > 0                     — mediating concept
# Isolated:    both = 0

classify_concepts <- function(node_df) {
  node_df %>%
    mutate(concept_type = case_when(
      indegree == 0 & outdegree  > 0 ~ "Transmitter",
      indegree  > 0 & outdegree == 0 ~ "Receiver",
      indegree  > 0 & outdegree  > 0 ~ "Ordinary",
      TRUE                           ~ "Isolated"
    ))
}

# ---- 4.3 Simple-cycle finder using DFS (Johnson-style) ----
# Returns a list of cycles (each cycle is a character vector of node names).
# max_length caps computation time for large networks.
# Cycles are reported as the ordered sequence of nodes.

find_simple_cycles <- function(g, max_length = 7) {
  n       <- vcount(g)
  if (n < 2) return(list())        # guard (matches Goal 2)
  vnames  <- V(g)$name
  # Build adjacency list (index-based)
  adj     <- igraph::as_adj_list(g, mode = "out")
  
  all_cycles <- list()
  
  # DFS from each starting node s (only go to nodes with index >= s
  # to avoid counting the same cycle multiple times in different rotations)
  dfs <- function(s_idx, curr_idx, path_idx) {
    if (length(path_idx) > max_length) return()
    
    nbrs <- as.integer(adj[[curr_idx]])
    for (nb in nbrs) {
      if (nb == s_idx && length(path_idx) >= 2) {
        # Closed cycle found
        all_cycles[[length(all_cycles) + 1]] <<-
          vnames[c(path_idx, s_idx)]
      } else if (nb > s_idx && !(nb %in% path_idx)) {
        # Extend path (only to higher-index nodes to avoid duplicates)
        dfs(s_idx, nb, c(path_idx, nb))
      }
    }
  }
  
  for (s in seq_len(n)) {
    dfs(s, s, s)
  }
  
  all_cycles
}

# ---- 4.4 Classify loop polarity ----
# A feedback loop is REINFORCING (positive / R-loop) if the product
# of all edge polarities around the loop equals +1.
# It is BALANCING (negative / B-loop) if the product equals -1.

classify_loop_polarity <- function(g, cycle) {
  # cycle is a character vector of node names including the return to start
  pol <- 1
  for (k in seq_len(length(cycle) - 1)) {
    eid <- get.edge.ids(g, c(cycle[k], cycle[k + 1]))
    if (eid == 0) return(NA)  # edge not found
    pol <- pol * sign(E(g)$weight[eid])
  }
  pol
}

# ---- 4.5 Prominence ----
# Prominence combines how widespread a concept is (occurrence across
# sub-models) with how central it is (normalized degree).
# For a FULL aggregated model, every concept has occurrence = 1.
# Hoffman et al., 2014

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
# SECTION 5 — COMPUTE ALL METRICS PER MODEL
# ============================================================
# This function computes every metric needed for Goal 1 and prints
# a structured report. It returns a named list for downstream comparison.

compute_all_metrics <- function(g, model_name) {
  
  cat("\n", strrep("=", 65), "\n")
  cat("  MODEL:", model_name, "\n")
  cat(strrep("=", 65), "\n")
  
  # ----------------------------------------------------------------
  # 5.1  BASIC STRUCTURAL PROPERTIES
  # ----------------------------------------------------------------
  N       <- vcount(g)
  E_count <- ecount(g)
  
  # Link density: fraction of all possible directed edges that exist.
  # Formula: D = E / [N × (N − 1)]
  # Interpretation: a dense model implies more perceived interconnections.
  density <- if (N > 1) E_count / (N * (N - 1)) else 0   # guard (matches Goal 2)
  
  cat("\n--- A. Basic structure ---\n")
  cat("  Concepts (N):      ", N, "\n")
  cat("  Connections (E):   ", E_count, "\n")
  cat("  Link density D = E/[N(N-1)]: ", round(density, 5), "\n")
  
  # ----------------------------------------------------------------
  # 5.2  CONCEPT TYPES & R/T RATIO
  # ----------------------------------------------------------------
  # Following Jones et al. (2011) and the FCM literature:
  # Transmitters carry system inputs; Receivers capture system outputs.
  # R/T = 0 → pure driver framing; high R/T → outcome-focused framing.
  
  node_df <- compute_degrees(g) %>% classify_concepts()

  n_tx  <- sum(node_df$concept_type == "Transmitter")
  n_rx  <- sum(node_df$concept_type == "Receiver")
  n_ord <- sum(node_df$concept_type == "Ordinary")
  n_iso <- sum(node_df$concept_type == "Isolated")
  
  RT_ratio <- if (n_tx > 0) n_rx / n_tx else NA_real_   # NA when undefined (matches Goal 2)
  
  cat("\n--- B. Concept types ---\n")
  cat("  Transmitters (pure drivers, indegree = 0): ", n_tx,  "\n")
  cat("  Receivers    (pure outcomes, outdegree = 0):", n_rx,  "\n")
  cat("  Ordinary     (both in and out links):      ", n_ord, "\n")
  cat("  Isolated     (no links at all):            ", n_iso, "\n")
  cat("  R/T ratio  = R / T =", n_rx, "/", n_tx, "=",
      round(RT_ratio, 3), "\n")
  
  cat("\n  Transmitters:\n")
  tx_names <- node_df$name[node_df$concept_type == "Transmitter"]
  cat(paste0("    ", tx_names), sep = "\n"); cat("\n")
  
  cat("  Receivers:\n")
  rx_names <- node_df$name[node_df$concept_type == "Receiver"]
  cat(paste0("    ", rx_names), sep = "\n"); cat("\n")
  
  # ----------------------------------------------------------------
  # 5.3  OUTDEGREE / INDEGREE / DEGREE CENTRALITY
  # ----------------------------------------------------------------
  # Outdegree_i = Σ_j |w_{ij}|   (influence exerted BY concept i)
  # Indegree_i  = Σ_j |w_{ji}|   (influence received BY concept i)
  # Degree_i    = Outdegree_i + Indegree_i  (total embeddedness)
  #
  # These use absolute-value weights so a strong negative link and a
  # strong positive link both contribute equally to structural position.
  
  cat("\n--- C. Degree centrality (weighted, top 10) ---\n")
  cat("  Top 10 by OUTDEGREE (strongest drivers):\n")
  print(node_df %>% arrange(desc(outdegree)) %>% head(10) %>%
          select(name, outdegree, indegree, degree, concept_type),
        row.names = FALSE)
  
  cat("\n  Top 10 by INDEGREE (most-influenced concepts):\n")
  print(node_df %>% arrange(desc(indegree)) %>% head(10) %>%
          select(name, indegree, outdegree, degree, concept_type),
        row.names = FALSE)
  
  # ----------------------------------------------------------------
  # 5.4  BETWEENNESS, CLOSENESS, EIGENVECTOR CENTRALITY
  # ----------------------------------------------------------------
  # Betweenness: fraction of all shortest paths (between any two nodes)
  #   that pass through a given node. High betweenness = structural bridge.
  #   Formula (Freeman 1977): CB(v) = Σ_{s≠v≠t} σ(s,t|v) / σ(s,t)
  #   where σ(s,t) = number of shortest s→t paths, σ(s,t|v) = those
  #   passing through v.
  #
  # Closeness: reciprocal of mean shortest path to all reachable nodes.
  #   CC(v) = (n-1) / Σ_u d(v,u)
  #   High closeness = information spreads quickly FROM this node.
  #
  # Eigenvector: a node's importance weighted by its neighbors' importance.
  #   CE(v) = (1/λ) Σ_u A_{uv} CE(u)  where λ is the leading eigenvalue.
  #   High eigenvector = connected to well-connected nodes.
  #
  # For path-based metrics we use inverse |weight| as distance so that
  # stronger links represent SHORTER conceptual distances.
  
  E(g)$dist_weight <- 1 / (abs(E(g)$weight) + 1e-6)  # avoid div/0
  
  btw <- betweenness(g, weights = E(g)$dist_weight, normalized = TRUE)
  clo <- closeness(g,  weights = E(g)$dist_weight, normalized = TRUE,
                   mode = "out")
  eig <- eigen_centrality(g, weights = abs(E(g)$weight),
                          directed = TRUE)$vector
  
  node_df$betweenness <- btw
  node_df$closeness   <- clo
  node_df$eigenvector <- eig
  
  cat("\n--- D. Centrality (top 10 by betweenness) ---\n")
  print(node_df %>% arrange(desc(betweenness)) %>% head(10) %>%
          select(name, betweenness, closeness, eigenvector, degree),
        row.names = FALSE)
  
  # ----------------------------------------------------------------
  # 5.5  PROMINENCE
  # ----------------------------------------------------------------

  node_df <- compute_prominence(g, node_df)
  
  cat("\n--- E. Prominence (top 10) ---\n")
  print(node_df %>% arrange(desc(prominence)) %>% head(10) %>%
          select(name, prominence, degree, concept_type),
        row.names = FALSE)
  
  # ----------------------------------------------------------------
  # 5.6  LINK POLARITY
  # ----------------------------------------------------------------
  # Count positive (reinforcing) vs. negative (inhibiting) causal links.
  # The ratio characterises whether the model is predominantly driven
  # by reinforcing (+) or balancing (−) relationships.
  
  edge_df <- igraph::as_data_frame(g, what = "edges")
  n_pos <- sum(edge_df$polarity > 0, na.rm = TRUE)
  n_neg <- sum(edge_df$polarity < 0, na.rm = TRUE)
  
  cat("\n--- F. Link polarity ---\n")
  cat("  Positive links:", n_pos, "(", round(n_pos / E_count * 100, 1), "% )\n")
  cat("  Negative links:", n_neg, "(", round(n_neg / E_count * 100, 1), "% )\n")
  cat("  Pos/Neg ratio: ", round(n_pos / max(n_neg, 1), 2), "\n")
  
  # ----------------------------------------------------------------
  # 5.7  AVERAGE PATH LENGTH AND DIAMETER
  # ----------------------------------------------------------------
  # Average path length (APL): mean of all pairwise shortest distances.
  #   Small APL → information and impacts spread quickly.
  # Diameter: longest shortest path = maximum number of steps separating
  #   the two most distant concepts. Large diameter → long causal chains.
  # Both are computed on the unweighted topology (hop count).
  
  dist_mat  <- distances(g, weights = NA)
  fin_dists <- dist_mat[is.finite(dist_mat) & dist_mat > 0]
  apl       <- if (length(fin_dists) > 0) mean(fin_dists) else NA_real_          # guard (matches Goal 2)
  diam      <- if (length(fin_dists) > 0) max(dist_mat[is.finite(dist_mat)]) else NA_real_
  
  cat("\n--- G. Path metrics ---\n")
  cat("  Average path length (APL):", round(apl,  3), "\n")
  cat("  Diameter:                 ", diam, "\n")
  
  # ----------------------------------------------------------------
  # 5.8  CLUSTERING COEFFICIENT
  # ----------------------------------------------------------------
  # The average (local) clustering coefficient measures how densely
  # interconnected a node's neighbourhood is.
  # C_i = (number of directed triangles through i) /
  #         (k_i × (k_i − 1))  where k_i is degree of node i.
  # High average C → tight local clusters; low C → chain-like structure.
  
  clust_avg <- transitivity(g, type = "average",  isolates = "zero")
  clust_glo <- transitivity(g, type = "global")
  
  cat("\n--- H. Clustering ---\n")
  cat("  Average local clustering coefficient:", round(clust_avg, 4), "\n")
  cat("  Global clustering coefficient:       ", round(clust_glo, 4), "\n")
  
  # ----------------------------------------------------------------
  # 5.9  FEEDBACK LOOPS
  # ----------------------------------------------------------------
  # Circuit rank (cyclomatic number):
  #   CR = E − N + C   where C = number of weakly connected components.
  # Interpretation: CR is the minimum number of independent feedback
  # loops; it quantifies the model's dynamic complexity.
  #
  # We also find all ACTUAL simple cycles up to max_length = 7 steps
  # and classify each as Reinforcing (R-loop, product of polarities = +1)
  # or Balancing (B-loop, product of polarities = −1).
  #
  # NOTE on delays (Schaffernicht & Groesser §3.2):
  # The original method distinguishes delayed from non-delayed links.
  # Our Kumu data does not record explicit delays, so all links are
  # treated as non-delayed (b = 1 effectively), consistent with standard
  # FCM/CLD practice.
  
  n_comp    <- components(g, mode = "weak")$no
  circ_rank <- E_count - N + n_comp
  
  cat("\n--- I. Feedback loops ---\n")
  cat("  Circuit rank (min independent loops):", circ_rank, "\n")
  cat("  Finding actual simple cycles (length ≤ 7) ...\n")
  
  cycles <- if (N >= 2 && E_count >= 2) find_simple_cycles(g, max_length = 7) else list()  # guard (matches Goal 2)
  cat("  Simple cycles found (length ≤ 7):   ", length(cycles), "\n")
  
  n_reinf <- 0; n_bal <- 0; n_na <- 0
  if (length(cycles) > 0) {
    pols <- sapply(cycles, classify_loop_polarity, g = g)
    n_reinf <- sum(pols >  0, na.rm = TRUE)
    n_bal   <- sum(pols <  0, na.rm = TRUE)
    n_na    <- sum(is.na(pols))
    cat("    Reinforcing (R-loops, positive product):", n_reinf, "\n")
    cat("    Balancing   (B-loops, negative product):", n_bal,   "\n")
    if (n_na > 0) cat("    Undetermined:", n_na, "\n")
  }
  
  # ----------------------------------------------------------------
  # RETURN EVERYTHING FOR DOWNSTREAM COMPARISON
  # ----------------------------------------------------------------
  invisible(list(
    model_name   = model_name,
    N = N, E = E_count,
    density      = density,
    n_tx = n_tx, n_rx = n_rx, n_ord = n_ord,
    RT_ratio     = RT_ratio,
    n_pos = n_pos, n_neg = n_neg,
    apl = apl, diam = diam,
    clust_avg = clust_avg, clust_glo = clust_glo,
    circ_rank = circ_rank,
    n_cycles     = length(cycles),
    n_reinforcing = n_reinf, n_balancing = n_bal,
    node_df      = node_df,
    cycles       = cycles,
    g            = g
  ))
}

# ---- RUN ----
metrics_au <- compute_all_metrics(g_au, "Australia")
metrics_al <- compute_all_metrics(g_al, "US Gulf Coast")



# ============================================================
# SECTION 6 — COMPARATIVE SUMMARY TABLE
# ============================================================

comparison_table <- data.frame(
  Metric                = c(
    "Concepts (N)",
    "Connections (E)",
    "Link Density D",
    "Transmitters",
    "Receivers",
    "Ordinary",
    "R/T Ratio",
    "Positive links (%)",
    "Negative links (%)",
    "Average Path Length",
    "Diameter",
    "Avg Clustering Coeff.",
    "Circuit Rank",
    "Simple Cycles (≤7)",
    "Reinforcing loops",
    "Balancing loops"
  ),
  Australia  = c(
    metrics_au$N, metrics_au$E, round(metrics_au$density, 5),
    metrics_au$n_tx, metrics_au$n_rx, metrics_au$n_ord,
    round(metrics_au$RT_ratio, 3),
    round(metrics_au$n_pos / metrics_au$E * 100, 1),
    round(metrics_au$n_neg / metrics_au$E * 100, 1),
    round(metrics_au$apl, 3), metrics_au$diam,
    round(metrics_au$clust_avg, 4),
    metrics_au$circ_rank, metrics_au$n_cycles,
    metrics_au$n_reinforcing, metrics_au$n_balancing
  ),
  Alabama_GulfCoast = c(
    metrics_al$N, metrics_al$E, round(metrics_al$density, 5),
    metrics_al$n_tx, metrics_al$n_rx, metrics_al$n_ord,
    round(metrics_al$RT_ratio, 3),
    round(metrics_al$n_pos / metrics_al$E * 100, 1),
    round(metrics_al$n_neg / metrics_al$E * 100, 1),
    round(metrics_al$apl, 3), metrics_al$diam,
    round(metrics_al$clust_avg, 4),
    metrics_al$circ_rank, metrics_al$n_cycles,
    metrics_al$n_reinforcing, metrics_al$n_balancing
  )
)

cat("\n\n", strrep("=", 65), "\n")
cat("  COMPARATIVE SUMMARY TABLE\n")
cat(strrep("=", 65), "\n")
print(comparison_table, row.names = FALSE)


# ============================================================
# SECTION 7 — GALVESTON (THIRD STANDALONE MODEL)
# ============================================================
# Galveston is added ONLY for the metric-calculation process, as a third
# standalone model computed with exactly the same pipeline as Australia and
# Alabama (parse → deduplicate_edges → build_graph → compute_all_metrics).
# It is deliberately NOT added to comparison_table, full_df_g1, or any plot,
# so every figure below still shows only Australia and Gulf Coast.
#
# Strength scale: integer ±1/±2 (same as AU/GC) — no rescaling applied.
# >>> SET THE FILE PATH BELOW to the Galveston Kumu export before running. <<<

GV_FILE <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Exported_Kumu_Galveston_30April.xlsx"  # <-- EDIT THIS PATH

gv_elements    <- read_excel(GV_FILE, sheet = "Elements")
gv_connections <- read_excel(GV_FILE, sheet = "Connections")

cat("Raw rows — Galveston elements:", nrow(gv_elements),
    " | Galveston connections:", nrow(gv_connections), "\n")

# Strength-scale check: confirm integer ±1/±2 (no ±0.5/±1.0 Mental Modeler values)
gv_strength_vals <- sort(unique(suppressWarnings(as.numeric(gv_connections$Strength))))
cat("Galveston unique Strength values:",
    paste(gv_strength_vals, collapse = ", "), "\n")
if (any(abs(gv_strength_vals) > 0 & abs(gv_strength_vals) < 1, na.rm = TRUE))
  warning("Galveston Strength contains |value| < 1 (looks like the ±0.5/±1.0 ",
          "Mental Modeler scale). Multiply by 2 to match the integer ±1/±2 ",
          "scale used for Australia and Gulf Coast before trusting weighted metrics.")

gv_conn  <- parse_connections(gv_connections)
gv_edges <- deduplicate_edges(gv_conn)
gv_nodes <- unique(gv_elements$Label)

cat("After deduplication:\n")
cat("  Galveston — nodes:", length(gv_nodes), " | edges:", nrow(gv_edges), "\n")

g_gv       <- build_graph(gv_nodes, gv_edges)
metrics_gv <- compute_all_metrics(g_gv, "Galveston")


# ============================================================
# SECTION 8 — VISUALISATIONS
# ============================================================

# ---- 8.1 Indegree vs Outdegree scatter (driver-outcome map) ----
# Concepts above the diagonal (indegree = outdegree line) are net drivers;
# those below are net receivers. Labels shown for the top-20 by degree.

plot_io_scatter <- function(metrics_obj, top_n = 20) {
  nd <- metrics_obj$node_df
  top_labels <- nd %>% arrange(desc(degree)) %>% head(top_n)
  
  ggplot(nd, aes(x = indegree, y = outdegree, colour = concept_type)) +
    geom_point(alpha = 0.6, size = 2.5) +
    geom_abline(slope = 1, intercept = 0, linetype = "dashed",
                colour = "grey40") +
    ggrepel::geom_text_repel(
      data = top_labels,
      aes(label = str_wrap(name, 25)),
      size = 2.8, max.overlaps = 15, show.legend = FALSE
    ) +
    scale_colour_manual(values = c("Transmitter" = "#D55E00",
                                   "Receiver"    = "#56B4E9",
                                   "Ordinary"    = "#009E73",
                                   "Isolated"    = "grey80")) + #color blind friendly palette
    labs(
      title   = paste0("Driver–Outcome Map: ", metrics_obj$model_name),
      x       = "Indegree (influence RECEIVED, Σ|w|)",
      y       = "Outdegree (influence EXERTED, Σ|w|)",
      colour  = "Concept type",
      caption = "Dashed line: equal driver and outcome weight"
    ) +
    theme_bw(base_size = 12)
}

# Install ggrepel if needed
if (!requireNamespace("ggrepel", quietly = TRUE)) install.packages("ggrepel")
library(ggrepel)

p_au <- plot_io_scatter(metrics_au, top_n = 15)
p_al <- plot_io_scatter(metrics_al, top_n = 15)

ggsave("IO_scatter_Australia.pdf",   p_au, width = 10, height = 8)
ggsave("IO_scatter_Alabama.pdf",     p_al, width = 10, height = 8)


# ---- 8.2 Comparative bar chart: key structural metrics ----

long_comp <- comparison_table %>%
  filter(Metric %in% c("Concepts (N)", "Connections (E)",
                       "Transmitters", "Receivers", "Circuit Rank",
                       "Reinforcing loops", "Balancing loops")) %>%
  pivot_longer(cols = c(Australia, Alabama_GulfCoast),
               names_to = "Model", values_to = "Value") %>%
  mutate(Value = as.numeric(Value))

p_bar <- ggplot(long_comp, aes(x = Metric, y = Value, fill = Model)) +
  geom_col(position = "dodge") +
  coord_flip() +
  scale_fill_manual(values = c("Australia" = "#0072B2",
                               "Alabama_GulfCoast" = "#D55E00")) +
  labs(title = "Structural comparison: Australia vs Gulf Coast USA",
       x = NULL, y = "Value", fill = "Region") +
  theme_bw(base_size = 12)

ggsave("Structural_comparison_barplot.pdf", p_bar, width = 10, height = 6)


# ---- 8.3 Betweenness centrality bar charts (top 15) ----

plot_betweenness <- function(metrics_obj, top_n = 15) {
  metrics_obj$node_df %>%
    arrange(desc(betweenness)) %>%
    head(top_n) %>%
    mutate(name = str_wrap(name, 30),
           name = factor(name, levels = rev(name))) %>%
    ggplot(aes(x = name, y = betweenness, fill = concept_type)) +
    geom_col() +
    coord_flip() +
    scale_colour_manual(values = c("Transmitter" = "#D55E00",
                                   "Receiver"    = "#56B4E9",
                                   "Ordinary"    = "#009E73",
                                   "Isolated"    = "grey80")) + #color blind friendly palette
    labs(title = paste0("Betweenness centrality (top ", top_n, "): ",
                        metrics_obj$model_name),
         x = NULL, y = "Normalised betweenness", fill = "Type") +
    theme_bw(base_size = 11)
}

plot_betweenness(metrics_au)
plot_betweenness(metrics_al)

ggsave("Betweenness_Australia.pdf",
       plot_betweenness(metrics_au), width = 10, height = 7)
ggsave("Betweenness_Alabama.pdf",
       plot_betweenness(metrics_al), width = 10, height = 7)


# ---- 8.4 Prominence bar chart (top 15) ----

plot_prominence <- function(metrics_obj, top_n = 15) {
  metrics_obj$node_df %>%
    arrange(desc(prominence)) %>%
    head(top_n) %>%
    mutate(name = str_wrap(name, 30),
           name = factor(name, levels = rev(name))) %>%
    ggplot(aes(x = name, y = prominence, fill = concept_type)) +
    geom_col() +
    coord_flip() +
    scale_colour_manual(values = c("Transmitter" = "#D55E00",
                                   "Receiver"    = "#56B4E9",
                                   "Ordinary"    = "#009E73",
                                   "Isolated"    = "grey80")) + #color blind friendly palette
    labs(title = paste0("Prominence: ", metrics_obj$model_name),
         x = NULL, y = "Prominence score", fill = "Type") +
    theme_bw(base_size = 11)
}

plot_prominence(metrics_au)
plot_prominence(metrics_al)

ggsave("Prominence_Australia.pdf",
       plot_prominence(metrics_au), width = 10, height = 7)
ggsave("Prominence_Alabama.pdf",
       plot_prominence(metrics_al), width = 10, height = 7)

cat("\n\nAll plots saved. Analysis complete.\n")


# ---- 8.5 BUILD COMBINED NODE TABLE FROM GOAL 1 METRICS ----

full_df_g1 <- bind_rows(
  metrics_au$node_df %>% mutate(Region = metrics_au$model_name),
  metrics_al$node_df %>% mutate(Region = metrics_al$model_name)
) %>%
  rename(
    Concept     = name,
    Type        = concept_type,
    Degree      = degree,
    Indegree    = indegree,
    Outdegree   = outdegree,
    Closeness   = closeness,
    Betweenness = betweenness
  )


# Shared plot theme
theme_fcm <- function(base_size = 10) {
  theme_bw(base_size = base_size) %+replace%
    theme(
      strip.text = element_text(
        face = "bold",
        size = base_size + 1,
        margin = margin(b = 6)),
      strip.background = element_rect(colour = NA),
      axis.text.y      = element_text(size = base_size - 2),
      axis.text.x      = element_text(size = base_size - 2),
      panel.grid.minor = element_blank(),
      legend.position  = "bottom",
      legend.title     = element_text(size = base_size - 1),
      plot.title       = element_text(face = "bold", size = base_size + 2, margin = margin(b = 4)),
      plot.subtitle    = element_text(size = base_size - 1, margin=margin(b=8))
    )
}



#  ------- 8.6. PLOT G1-6 — INDEGREE vs OUTDEGREE -------

cat("Plot G1-6: Indegree vs Outdegree...\n")

# ---- Region order: USGC left, Australia right ----
full_df_g1 <- full_df_g1 %>%
  mutate(Region = factor(
    recode(Region,
           "Alabama (Gulf Coast USA)" = "US Gulf Coast (USGC)",
           "Australia"                = "Australia"),
    levels = c("US Gulf Coast (USGC)", "Australia")
  ))

# ---- Threshold diagnostic: Degree with p-quantile cutoff per region ----

p_keep <- 0.80   # keep top 20% (>= 80th percentile). Keep nodes above the 80th percentile.

thr_g1 <- full_df_g1 %>%
  group_by(Region) %>%
  summarise(thresh = quantile(Degree, p_keep, na.rm = TRUE), .groups = "drop")

diag_df <- full_df_g1 %>%
  left_join(thr_g1, by = "Region") %>%
  group_by(Region) %>%
  mutate(
    rank_deg = rank(-Degree, ties.method = "first"),
    keep     = Degree >= thresh
  ) %>%
  ungroup()

p_thresh <- ggplot(diag_df,
                   aes(x = reorder(Concept, Degree), y = Degree, colour = keep)) +
  geom_segment(aes(xend = Concept, y = 0, yend = Degree),
               colour = "grey80", linewidth = 0.3) +
  geom_point(size = 1.8) +
  geom_hline(data = thr_g1, aes(yintercept = thresh),
             linetype = "dashed", colour = "red", linewidth = 0.5) +
  geom_text(data = thr_g1,
            aes(x = 1, y = thresh,
                label = paste0("p", round(p_keep*100), " = ", round(thresh, 1))),
            inherit.aes = FALSE, hjust = 0, vjust = -0.5,
            size = 3, colour = "red") +
  coord_flip() +
  scale_colour_manual(values = c("TRUE" = "#009E73", "FALSE" = "grey65"),
                      labels = c("TRUE" = "Labeled (kept)", "FALSE" = "Not labeled"),
                      name   = NULL) +
  facet_wrap(~ Region, nrow = 1, scales = "free_y") +
  labs(
    title = "Degree distribution and labeling threshold",
    x     = NULL,
    y     = "Degree (total connectivity)"
  ) +
  theme_fcm()

#Values of Depredation
full_df_g1 %>%
  filter(Concept == "Depredation") %>%
  select(Region, Indegree, Outdegree, Degree) %>%
  arrange(Region) %>%
  print()


# ---- Panel tags (a)/(b) ----
region_labs <- c(
  "US Gulf Coast (USGC)" = "US Gulf Coast (USGC)",
  "Australia"            = "Australia"
)

# ---- Drop Depredation so axes rescale to the rest of the concepts ----
plot_df_g1 <- full_df_g1 %>% filter(Concept != "Depredation")


# ---- Labels: only concepts at/above the p80 Degree threshold per region ----

label_df_g1 <- plot_df_g1 %>%
  group_by(Region) %>%
  filter(Degree >= quantile(Degree, p_keep, na.rm = TRUE)) %>%
  #concepts on above the 80 percentile by Degree (indegree+outdegree) in each panel gets a label box
  ungroup()

# recompute 1:1 annotation position on the reduced data
gm <- with(plot_df_g1,
           min(max(Indegree, na.rm = TRUE), max(Outdegree, na.rm = TRUE)))

#Plot
p_g1_6 <- ggplot(plot_df_g1,
                 aes(x = Indegree, y = Outdegree,
                     colour = Type, size = Degree)) +
  
  geom_abline(slope = 1, intercept = 0,
              linetype = "dashed", colour = "grey60", linewidth = 0.5) +
  
  geom_point(alpha = 0.6, stroke = 0.2) +
  
  annotate("text", x = 0.80 * gm, y = 0.80 * gm, label = "1:1 line",
           angle = 45, vjust = -1, size = 3, colour = "grey45") +
  
  geom_label_repel(
    data          = label_df_g1,
    aes(label     = str_wrap(Concept, 20)),
    size          = 2.6,
    label.padding = unit(0.12, "lines"),
    box.padding   = unit(0.4, "lines"),
    point.padding = unit(0.2, "lines"),
    min.segment.length = 0,
    segment.colour = "grey55",
    segment.size  = 0.3,
    max.overlaps  = Inf,          # don't silently drop labels
    seed          = 42,           # reproducible label positions
    show.legend   = FALSE,
    colour        = "grey15",
    fill          = alpha("white", 0.85)
  ) +
  
  scale_colour_manual(values = c("Transmitter" = "#D55E00",
                                 "Receiver"    = "#56B4E9",
                                 "Ordinary"    = "#009E73", #color blind friendly palette
                                 "Isolated"    = "grey80"), drop = FALSE) +
  
  guides(colour = guide_legend(override.aes = list(size = 4, alpha = 1))) +
  
  scale_size_continuous(range = c(1.5, 6), 
                        guide = "none") + #change here to show legend of dot size. Change to: guide = guide_legend(title = "Degree")
  
  facet_wrap(~ Region, nrow = 1, scales = "fixed", #scales fixed (shared axes) to allow comparison
             labeller = labeller(Region = region_labs)) + 
  
  labs(
    x        = "Indegree",
    y        = "Outdegree",
    colour   = "Concept type"
  ) +
  theme_fcm()+
  theme(
    plot.title    = element_text(size = 14, face = "bold"),
    axis.title    = element_text(size = 11),
    axis.text     = element_text(size = 9),
    strip.text    = element_text(size = 11, face = "bold"),
    legend.title  = element_text(size = 10),
    legend.text   = element_text(size = 9),
    plot.caption  = element_text(size = 8, hjust = 0, colour = "grey30")
  )


# ---- Export: PNG + vector PDF + 600-dpi LZW TIFF ----
ggsave("G1_P6_indegree_vs_outdegree.png", p_g1_6,
       width = 15, height = 7, units = "in", dpi = 1200)

ggsave("G1_P6_indegree_vs_outdegree.tiff", p_g1_6,
       width = 17, height = 8, units = "in", dpi = 1200,
       compression = "lzw")

pdf(filename = "G1_P6_indegree_vs_outdegree.png",
    width = 17, height = 8, units = "in", res = 1200)
p_g1_6
dev.off()



# ----- 8.6B - PLOT G1-6b — CLOSENESS vs BETWEENNESS -------

cat("Plot G1-6b: Closeness vs Betweenness...\n")

# ---- Region order: USGC left, Australia right ----
full_df_g1 <- full_df_g1 %>%
  mutate(Region = factor(
    recode(Region,
           "Alabama (Gulf Coast USA)" = "US Gulf Coast (USGC)",
           "Australia"                = "Australia"),
    levels = c("US Gulf Coast (USGC)", "Australia")
  ))

p_keep_n <- 10

plot_df_cb <- full_df_g1 %>%
  filter(!is.na(Closeness), !is.na(Betweenness)) %>%
  filter(Concept != "Depredation") %>%          # drop Depredation
  group_by(Region) %>%
  mutate(
    keep = rank(-Closeness, ties.method = "min") <= p_keep_n |
      rank(-Betweenness, ties.method = "min") <= p_keep_n,
    lab  = ifelse(keep, str_wrap(Concept, 16), "")
  ) %>%
  ungroup()

# recompute median lines WITHOUT Depredation so quadrants match the plotted points
quad_lines_g1 <- plot_df_cb %>%
  group_by(Region) %>%
  summarise(
    med_closeness   = median(Closeness,   na.rm = TRUE),
    med_betweenness = median(Betweenness, na.rm = TRUE),
    .groups = "drop"
  )

#Depredation values
full_df_g1 %>%
  filter(Concept == "Depredation") %>%
  select(Region, Closeness, Betweenness) %>%
  arrange(Region) %>%
  print()


p_g1_6b <- ggplot(plot_df_cb,
                  aes(x = Closeness, y = Betweenness)) +
  geom_vline(data = quad_lines_g1, aes(xintercept = med_closeness),
             linetype = "dashed", colour = "grey70", linewidth = 0.4) +
  geom_hline(data = quad_lines_g1, aes(yintercept = med_betweenness),
             linetype = "dashed", colour = "grey70", linewidth = 0.4) +
  geom_point(aes(colour = Type, size = Degree), alpha = 0.6, stroke = 0.2) +
  
  geom_label_repel(
    data          = plot_df_cb,
    aes(label     = lab),
    size          = 2.6,
    label.padding = unit(0.12, "lines"),
    box.padding   = unit(0.5, "lines"),
    point.padding = unit(0.4, "lines"),
    min.segment.length = 0,
    segment.colour = "grey55", segment.size = 0.3,
    max.overlaps  = Inf, seed = 42,
    show.legend   = FALSE, colour = "grey15",
    fill          = alpha("white", 0.85)
  ) +
  
  scale_colour_manual(values = c("Transmitter" = "#D55E00",
                                 "Receiver"    = "#56B4E9",
                                 "Ordinary"    = "#009E73",
                                 "Isolated"    = "grey80"), drop = FALSE) +
  guides(colour = guide_legend(override.aes = list(size = 4, alpha = 1))) +
  
  scale_size_continuous(range = c(1.5, 6), guide = "none") +
  scale_x_sqrt() +          # spread the low-closeness cluster
  scale_y_sqrt() +          # spread the near-zero betweenness cluster
  facet_wrap(~ Region, nrow = 1, scales = "fixed") +
  labs(
    x      = "Closeness (normalized, outgoing; higher = faster broadcaster)",
    y      = "Betweenness (normalized; higher = more of a bridge/bottleneck)",
    colour = "Concept type"
    ) +
  theme_fcm() +
  theme(
    plot.title = element_text(size = 14, face = "bold"),
    axis.title = element_text(size = 11), axis.text = element_text(size = 9),
    strip.text = element_text(size = 11, face = "bold"),
    legend.title = element_text(size = 10), legend.text = element_text(size = 9),
    plot.caption = element_text(size = 8, hjust = 0, colour = "grey30")
  )



# ---- Export: PNG + vector PDF + 600-dpi LZW TIFF ----
ggsave("G1_P6b_closeness_vs_betweenness.png", p_g1_6b,
       width = 15, height = 7, units = "in", dpi = 1200)

ggsave("G1_P6b_closeness_vs_betweenness.pdf", p_g1_6b,
       width = 17, height = 8, units = "in", device = cairo_pdf)

ggsave("G1_P6b_closeness_vs_betweenness.tiff", p_g1_6b,
       width = 17, height = 8, units = "in", dpi = 600,
       compression = "lzw")












# ============================================================
# ============== SOME EXTRA ANALYSIS =========================
# ============================================================


# SCHAFFERNICHT & GROESSER (2011): ELEMENT DISTANCE RATIO
# ------------------------------------------------------------

# Paper §3.3 — The EDR is computed using signed adjacency matrices.
#
# PARAMETERS (for system dynamics / FCMs):
#   a = 1   → self-loops excluded
#   b = 2   → maximum link strength (our scale goes from −2 to +2)
#   c = 2   → differences involving unique-model variables are meaningful
#   d = 0   → polarity difference does not depend on link strength
#   e = 2   → two possible polarities (positive, negative)
#
# VARIABLE SETS (following Paper §3, notation):
#   V_A  = set of all variables in model A
#   V_B  = set of all variables in model B
#   V_c  = V_A ∩ V_B  (common variables)
#   V_uA = V_A \ V_c  (unique to A)
#   V_uB = V_B \ V_c  (unique to B)
#   v_c, v_uA, v_uB = cardinalities of the above sets
#
# DIFF FUNCTION (Paper §3.3, Eq. 2):
#   For each ordered pair (i, j) with i ≠ j in the UNION variable space:
#
#   (i)  If i = j:         diff = 0  [no self-loops, a = 1]
#
#   (ii) If either i or j is a unique variable (not in V_c):
#          diff = C(a_ij, b_ij) where C = 1 unless both models have 0
#          (because c = 2, unique-variable links count as differences)
#          In practice: if a link exists in one model, diff = its strength
#          (max = b = 2); if neither has the link, diff = 0.
#
#   (iii) If both i and j are common variables:
#          diff = |a_ij − b_ij|
#          where a_ij, b_ij are SIGNED strengths (0 if no link).
#          Maximum possible = e × b + d = 4 (link +2 in A vs −2 in B).
#
# DENOMINATOR (maximum possible total difference):
#   Computed from all ordered pairs by category:
#
#   Common–common pairs (v_c × (v_c − 1) directed pairs, no self-loops):
#     max diff per pair = e × b + d = 4
#
#   All other pairs (one or both from unique sets):
#     max diff per pair = b = 2  (max strength in one model, 0 in other)
#     Total such pairs = v × (v−1) − v_c × (v_c−1)
#     where v = v_c + v_uA + v_uB
#
#   Denominator = 4 × v_c × (v_c − 1) +
#                 2 × [v × (v−1) − v_c × (v_c−1)]
#
# EDR ranges from 0 (identical) to 1 (completely different).
# Thresholds (paper §3.5): EDR < 0.25 = similar; EDR > 0.25 = dissimilar.

compute_EDR <- function(metrics_A, metrics_B,
                        b = 2, concept_map = NULL) {
  # concept_map (optional): named character vector mapping concept names in
  # model A to equivalent names in model B when they differ linguistically.
  # Example: c("Shark Populations" = "Shark Abundance")
  # If NULL, exact string matching is used.
  
  g_A <- metrics_A$g
  g_B <- metrics_B$g
  
  cat("\n", strrep("=", 65), "\n")
  cat("  ELEMENT DISTANCE RATIO  (Schaffernicht & Groesser 2011, §3.3)\n")
  cat("  Comparing:", metrics_A$model_name, "vs", metrics_B$model_name, "\n")
  cat(strrep("=", 65), "\n")
  
  # ---- Variable sets ----
  vars_A <- V(g_A)$name
  vars_B <- V(g_B)$name
  
  # Apply optional concept mapping to harmonise names
  if (!is.null(concept_map)) {
    vars_A_mapped <- ifelse(vars_A %in% names(concept_map),
                            concept_map[vars_A], vars_A)
  } else {
    vars_A_mapped <- vars_A
  }
  
  V_c   <- intersect(vars_A_mapped, vars_B)   # common variables
  V_uA  <- setdiff(vars_A_mapped, V_c)        # unique to A
  V_uB  <- setdiff(vars_B, V_c)               # unique to B
  v_c   <- length(V_c)
  v_uA  <- length(V_uA)
  v_uB  <- length(V_uB)
  v     <- v_c + v_uA + v_uB
  
  cat("  Common variables   (V_c):   ", v_c,  "\n")
  cat("  Unique to A (V_uA):         ", v_uA, "\n")
  cat("  Unique to B (V_uB):         ", v_uB, "\n")
  cat("  Total variable space (v):   ", v,    "\n")
  
  cat("\n  Common concepts:\n")
  cat(paste0("    ", V_c), sep = "\n"); cat("\n")
  
  # ---- Build signed adjacency matrices ----
  # Both matrices use the FULL variable space (union of V_A and V_B).
  # Entries outside each model's variable set are 0.
  
  build_adj_matrix <- function(g, var_names, model_vars) {
    mat <- matrix(0, nrow = length(var_names), ncol = length(var_names),
                  dimnames = list(var_names, var_names))
    edge_df <- igraph::as_data_frame(g, what = "edges")
    for (i in seq_len(nrow(edge_df))) {
      fr <- edge_df$from[i]
      to <- edge_df$to[i]
      w  <- edge_df$weight[i]
      # Only fill cells where both variables belong to this model
      if (fr %in% model_vars && to %in% model_vars) {
        mat[fr, to] <- w
      }
    }
    mat
  }
  
  all_vars   <- c(V_c, V_uA, V_uB)
  
  # Remap model-A variable names if concept_map provided
  vars_A_orig_to_mapped <- if (!is.null(concept_map)) concept_map else character(0)
  
  # For matrix building, use original names for model A (then relabel)
  mat_A_raw <- build_adj_matrix(g_A, vars_A, vars_A)
  mat_B_raw <- build_adj_matrix(g_B, vars_B, vars_B)
  
  # Place both matrices into the common variable space
  mat_A <- matrix(0, nrow = length(all_vars), ncol = length(all_vars),
                  dimnames = list(all_vars, all_vars))
  mat_B <- matrix(0, nrow = length(all_vars), ncol = length(all_vars),
                  dimnames = list(all_vars, all_vars))
  
  # Fill mat_A: use mapped names for rows/cols
  for (r in rownames(mat_A_raw)) {
    r_mapped <- if (r %in% names(concept_map)) concept_map[r] else r
    for (cc in colnames(mat_A_raw)) {
      c_mapped <- if (cc %in% names(concept_map)) concept_map[cc] else cc
      if (r_mapped %in% all_vars && c_mapped %in% all_vars) {
        mat_A[r_mapped, c_mapped] <- mat_A_raw[r, cc]
      }
    }
  }
  
  # Fill mat_B directly (names already in common space)
  for (r in rownames(mat_B_raw)) {
    for (cc in colnames(mat_B_raw)) {
      if (r %in% all_vars && cc %in% all_vars) {
        mat_B[r, cc] <- mat_B_raw[r, cc]
      }
    }
  }
  
  # ---- Compute diff matrix ----
  # For each ordered pair (i, j), i ≠ j:
  #   diff(i,j) = |mat_A[i,j] − mat_B[i,j]|
  
  diff_mat <- abs(mat_A - mat_B)
  diag(diff_mat) <- 0   # no self-loops (a = 1)
  
  numerator <- sum(diff_mat)
  
  # ---- Denominator ----
  # Maximum possible total difference:
  #   Common–common pairs: max diff = e×b+d = 4, count = v_c×(v_c−1)
  #   All other pairs:     max diff = b = 2,     count = v×(v−1) − v_c×(v_c−1)
  
  denom_cc    <- 4 * v_c * (v_c - 1)
  denom_other <- 2 * (v * (v - 1) - v_c * (v_c - 1))
  denominator <- denom_cc + denom_other
  
  EDR <- numerator / denominator
  
  cat("  Numerator  (actual differences):    ", round(numerator,  2), "\n")
  cat("  Denominator (max possible diff):    ", round(denominator, 2), "\n")
  cat("\n  *** EDR =", round(EDR, 4), "***\n")
  cat("  Interpretation:\n")
  if (EDR < 0.25) {
    cat("  → Models are SIMILAR at the element level (EDR < 0.25)\n")
  } else {
    cat("  → Models are DISSIMILAR at the element level (EDR ≥ 0.25)\n")
  }
  
  # ---- Additional element-level fractions ----
  # These follow Table 4 in the paper and provide interpretable context.
  cat("\n  Element-level comparison fractions (cf. Table 4 in paper):\n")
  cat("    Fraction unique to A:      ", round(v_uA / v, 3), "\n")
  cat("    Fraction unique to B:      ", round(v_uB / v, 3), "\n")
  cat("    Fraction common:           ", round(v_c  / v, 3), "\n")
  
  # Links
  e_A    <- nrow(igraph::as_data_frame(g_A, what = "edges"))
  e_B    <- nrow(igraph::as_data_frame(g_B, what = "edges"))
  # Common links: From→To pairs that exist in BOTH models (among V_c only)
  edge_A_cc <- igraph::as_data_frame(g_A, what = "edges") %>%
    filter(from %in% V_c, to %in% V_c) %>%
    mutate(pair = paste(from, to, sep = "→"))
  edge_B_cc <- igraph::as_data_frame(g_B, what = "edges") %>%
    filter(from %in% V_c, to %in% V_c) %>%
    mutate(pair = paste(from, to, sep = "→"))
  common_links_cc <- intersect(edge_A_cc$pair, edge_B_cc$pair)
  n_common_links  <- length(common_links_cc)
  
  cat("\n  Link fractions:\n")
  cat("    Links in A:                ", e_A, "\n")
  cat("    Links in B:                ", e_B, "\n")
  cat("    Common links (V_c pairs):  ", n_common_links, "\n")
  if (length(common_links_cc) > 0) {
    cat("    Common links list:\n")
    cat(paste0("      ", common_links_cc), sep = "\n"); cat("\n")
  }
  
  invisible(list(
    EDR = EDR, numerator = numerator, denominator = denominator,
    v_c = v_c, v_uA = v_uA, v_uB = v_uB,
    V_c = V_c, V_uA = V_uA, V_uB = V_uB,
    n_common_links = n_common_links,
    common_links   = common_links_cc,
    mat_A = mat_A, mat_B = mat_B, diff_mat = diff_mat
  ))
}



# SCHAFFERNICHT & GROESSER (2011): LOOP DISTANCE RATIO
# ------------------------------------------------------------

# Paper §3.4 — LDR compares individual feedback loops between models.
#
# Each loop is compared along three dimensions (Eq. 3):
#   LDR(m,n) = g × ldd(m,n) + i × lpold(m,n) + j × EDR(m,n)
#   Default weights: g = i = j = 1/3  (equal weighting)
#   where:
#     ldd(m,n)   = delay difference (0 if same delay pattern; 1 if different)
#     lpold(m,n) = loop polarity difference (0 if same polarity; 1 if different)
#     EDR(m,n)   = element distance ratio of the two loops' variable sets
#
# Loops unique to one model receive LDR = 1 (maximum distance).
#
# SILS = Shortest Independent Loop Set (Oliva, 2004):
#   The smallest set of shortest loops that covers all feedback-participating
#   variables. We approximate SILS by taking the minimum cycle basis of the
#   undirected version of the graph (igraph::minimum.spanning.tree approach),
#   then verifying directionality. For large models, we work with all simple
#   cycles up to length 7 (already found in Section 5.9) and select the
#   shortest non-redundant set covering all variables in any cycle.

compute_LDR_MDR <- function(metrics_A, metrics_B, edr_result,
                            w_g = 1/3, w_i = 1/3, w_j = 1/3) {
  
  cat("\n", strrep("=", 65), "\n")
  cat("  LOOP & MODEL DISTANCE RATIOS  (Schaffernicht & Groesser 2011, §3.4–3.5)\n")
  cat("  Comparing:", metrics_A$model_name, "vs", metrics_B$model_name, "\n")
  cat(strrep("=", 65), "\n")
  cat("  Weights: g (delay) =", w_g, "| i (polarity) =", w_i,
      "| j (EDR) =", w_j, "\n\n")
  
  cycles_A <- metrics_A$cycles
  cycles_B <- metrics_B$cycles
  g_A      <- metrics_A$g
  g_B      <- metrics_B$g
  V_c      <- edr_result$V_c
  
  cat("  Cycles in A (≤ 7 steps):", length(cycles_A), "\n")
  cat("  Cycles in B (≤ 7 steps):", length(cycles_B), "\n")
  
  if (length(cycles_A) == 0 || length(cycles_B) == 0) {
    cat("  One or both models have no cycles — MDR cannot be computed.\n")
    return(invisible(NULL))
  }
  
  # ---- SILS approximation ----
  # We build the SILS by taking the shortest cycles first and removing
  # variables already covered, until all feedback-participating variables
  # are covered. This follows Oliva (2004) in spirit.
  
  get_sils <- function(cycles, g) {
    if (length(cycles) == 0) return(list())
    # Sort cycles by length (shortest first)
    cycle_lengths <- sapply(cycles, length)
    sorted_idx    <- order(cycle_lengths)
    cycles_sorted <- cycles[sorted_idx]
    
    # Collect all variables appearing in ANY cycle
    all_cycle_vars <- unique(unlist(lapply(cycles, function(cyc) cyc[-length(cyc)])))
    covered        <- character(0)
    sils           <- list()
    
    for (cyc in cycles_sorted) {
      cyc_vars <- cyc[-length(cyc)]   # remove the repeated start node
      # Add this cycle if it covers at least one new variable
      if (any(!(cyc_vars %in% covered))) {
        sils    <- c(sils, list(cyc))
        covered <- unique(c(covered, cyc_vars))
      }
      if (all(all_cycle_vars %in% covered)) break
    }
    sils
  }
  
  sils_A <- get_sils(cycles_A, g_A)
  sils_B <- get_sils(cycles_B, g_B)
  
  cat("  SILS size — A:", length(sils_A), " | B:", length(sils_B), "\n")
  
  # ---- Semantic matching of loops ----
  # Two loops correspond if they share the same concept nodes (after mapping
  # to common variable names). We match each loop in SILS_A to the loop in
  # SILS_B with the greatest variable overlap.
  
  loop_vars_A <- lapply(sils_A, function(cyc) intersect(cyc[-length(cyc)], V_c))
  loop_vars_B <- lapply(sils_B, function(cyc) intersect(cyc[-length(cyc)], V_c))
  
  n_A <- length(sils_A)
  n_B <- length(sils_B)
  
  # Build a Jaccard-similarity matrix to identify candidate matches
  jaccard_mat <- matrix(0, nrow = n_A, ncol = n_B)
  for (ia in seq_len(n_A)) {
    for (ib in seq_len(n_B)) {
      inter <- length(intersect(loop_vars_A[[ia]], loop_vars_B[[ib]]))
      union <- length(union(loop_vars_A[[ia]], loop_vars_B[[ib]]))
      jaccard_mat[ia, ib] <- ifelse(union == 0, 0, inter / union)
    }
  }
  
  # Greedy matching: pair A-loop and B-loop with highest Jaccard;
  # threshold = 0.4 (must share at least 40% of variable space)
  threshold   <- 0.40
  matched_A   <- integer(0)
  matched_B   <- integer(0)
  match_pairs <- list()   # list of (ia, ib) pairs
  
  jac_copy <- jaccard_mat
  while (TRUE) {
    best_val <- max(jac_copy, na.rm = TRUE)
    if (best_val < threshold) break
    best_pos <- which(jac_copy == best_val, arr.ind = TRUE)[1, ]
    ia <- best_pos[1]; ib <- best_pos[2]
    match_pairs <- c(match_pairs, list(c(ia, ib)))
    matched_A   <- c(matched_A, ia)
    matched_B   <- c(matched_B, ib)
    jac_copy[ia, ] <- -Inf
    jac_copy[, ib] <- -Inf
  }
  
  unmatched_A <- setdiff(seq_len(n_A), matched_A)
  unmatched_B <- setdiff(seq_len(n_B), matched_B)
  
  cat("  Matched loop pairs (Jaccard ≥", threshold, "):", length(match_pairs), "\n")
  cat("  Loops unique to A:", length(unmatched_A),
      "| Unique to B:", length(unmatched_B), "\n")
  
  # ---- Compute LDR for each matched pair ----
  # Recall Eq. 3: LDR(m,n) = g×ldd + i×lpold + j×EDR_loop
  #   ldd    = 0 (no delay information in our data)
  #   lpold  = 0 if same polarity, 1 if different
  #   EDR_loop computed on the sub-graph of each loop
  
  ldr_values <- numeric(0)
  ldr_table  <- data.frame(
    loop_A = character(0), loop_B = character(0),
    polarity_A = character(0), polarity_B = character(0),
    jaccard = numeric(0), ldd = numeric(0),
    lpold = numeric(0), EDR_loop = numeric(0), LDR = numeric(0)
  )
  
  get_loop_polarity_label <- function(g, cyc) {
    p <- classify_loop_polarity(g, cyc)
    if (is.na(p)) "Unknown"
    else if (p > 0) "Reinforcing (+)"
    else "Balancing (−)"
  }
  
  for (pair in match_pairs) {
    ia <- pair[1]; ib <- pair[2]
    cyc_A <- sils_A[[ia]]; cyc_B <- sils_B[[ib]]
    
    pol_A  <- classify_loop_polarity(g_A, cyc_A)
    pol_B  <- classify_loop_polarity(g_B, cyc_B)
    lpold  <- ifelse(is.na(pol_A) | is.na(pol_B), 0,
                     as.numeric(pol_A != pol_B))
    
    # Build mini-graph for each loop to compute EDR_loop
    loop_vars_a <- cyc_A[-length(cyc_A)]
    loop_vars_b <- cyc_B[-length(cyc_B)]
    
    # Sub-graph: only keep edges among this loop's nodes
    sg_A <- induced_subgraph(g_A, loop_vars_a[loop_vars_a %in% V(g_A)$name])
    sg_B <- induced_subgraph(g_B, loop_vars_b[loop_vars_b %in% V(g_B)$name])
    
    # Create minimal metrics objects for EDR sub-computation
    mini_A <- list(g = sg_A, model_name = "LoopA")
    mini_B <- list(g = sg_B, model_name = "LoopB")
    
    edr_loop_res <- tryCatch(
      suppressMessages(
        compute_EDR(mini_A, mini_B, concept_map = NULL)
      ),
      error = function(e) list(EDR = 1)   # if sub-EDR fails, max distance
    )
    EDR_loop <- edr_loop_res$EDR
    
    ldd <- 0  # no delay information
    
    LDR_val <- w_g * ldd + w_i * lpold + w_j * EDR_loop
    
    jac_val <- jaccard_mat[ia, ib]
    
    ldr_values <- c(ldr_values, LDR_val)
    ldr_table <- rbind(ldr_table, data.frame(
      loop_A     = paste(cyc_A, collapse = "→"),
      loop_B     = paste(cyc_B, collapse = "→"),
      polarity_A = get_loop_polarity_label(g_A, cyc_A),
      polarity_B = get_loop_polarity_label(g_B, cyc_B),
      jaccard    = round(jac_val, 3),
      ldd        = ldd,
      lpold      = lpold,
      EDR_loop   = round(EDR_loop, 3),
      LDR        = round(LDR_val, 3)
    ))
  }
  
  # Unmatched loops get LDR = 1
  n_unmatched <- length(unmatched_A) + length(unmatched_B)
  ldr_values  <- c(ldr_values, rep(1, n_unmatched))
  
  # ---- Model Distance Ratio ----
  # MDR = mean of all LDRs (Eq. 4)
  # MDR < 0.25 → similar; MDR > 0.25 → dissimilar
  
  n_total_loops <- length(ldr_values)
  MDR <- mean(ldr_values)
  
  cat("\n  LDR table (matched pairs):\n")
  print(ldr_table[, c("polarity_A", "polarity_B", "jaccard",
                      "ldd", "lpold", "EDR_loop", "LDR")],
        row.names = FALSE)
  
  cat("\n  *** MDR =", round(MDR, 4), "(averaged over", n_total_loops,
      "loops) ***\n")
  if (MDR < 0.25) {
    cat("  → Models are SIMILAR at the loop/dynamic level (MDR < 0.25)\n")
  } else {
    cat("  → Models are DISSIMILAR at the loop/dynamic level (MDR ≥ 0.25)\n")
  }
  
  invisible(list(
    MDR = MDR, ldr_values = ldr_values, ldr_table = ldr_table,
    match_pairs = match_pairs, n_unmatched = n_unmatched,
    sils_A = sils_A, sils_B = sils_B
  ))
}



# OPTIONAL CONCEPT MAPPING
# ------------------------------------------------------------

# The Australia and Alabama models were elicited independently and
# may use different terminology for equivalent concepts.
# Add any semantic equivalences here as a named character vector:
#   names = Australia variable name
#   values = Alabama variable name
# Concepts not listed are matched by exact string only.

concept_map_AU_to_AL <- c(
  "Shark Abundance"      = "Shark Populations",
  "Fisher Satisfaction"  = "Fisher Satisfaction",   # same name
  "Number of Fishers"    = "Number of Fishers",     # same name
  "Fishing Time"         = "Fishing Time",           # same name
  "Fight Time"           = "Fight Time",             # same name
  "Shark Harvest"        = "Shark Harvest",          # same name
  "Shark Habituation"    = "Shark Learning Behavior",
  "Fisher Behaviour Change" = "Fisher Competency",
  "Fish Discards"        = "Discarding Small Fish",
  "Fishing Costs"        = "Economic Loss",
  "Shark Market Demand"  = "Diverse Markets For Shark Products",
  "Fisheries Health"     = "Fisheries Management Effectiveness",
  "Vessel Noise"         = "Boat Signature",
  "Fishing Effort Concentration" = "Fishing Effort",
  "Human Health/Safety"  = "Safety",
  "Shark Conservation"   = "Shark Conservation",
  "Tourism"              = "Tourist Angler Satisfaction",
  "Prey Availability"    = "Prey Populations (Menhaden)"
)



# RUN EDR, LDR, MDR
# ------------------------------------------------------------

edr_result  <- compute_EDR(metrics_au, metrics_al,
                           concept_map = concept_map_AU_to_AL)

ldr_mdr_result <- compute_LDR_MDR(metrics_au, metrics_al,
                                  edr_result = edr_result,
                                  w_g = 1/3, w_i = 1/3, w_j = 1/3)


cat("\n  Schaffernicht & Groesser (2011) distance ratios:\n")
cat("  EDR  =", round(edr_result$EDR, 4),
    "(element level)\n")
if (!is.null(ldr_mdr_result)) {
  cat("  MDR  =", round(ldr_mdr_result$MDR, 4),
      "(loop/dynamic level)\n")
}


