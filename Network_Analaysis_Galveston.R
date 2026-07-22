# ============================================================
# NETWORK STRUCTURE ANALYSIS — GALVESTON (TEXAS)
# ============================================================
# Analyzes the full Galveston FCM and each of its three
# fishing-sector sub-models as individual networks.
#
# LEVELS ANALYZED:
#   1. Galveston (full model — all 3 sectors combined)
#   2. Charter sector
#   3. Commercial sector
#   4. Recreational sector
#
# METRICS:
#   Part A — Concept types:   Transmitter, Receiver, Ordinary
#   Part B — Centrality:      Indegree, Outdegree, Degree,
#                             Betweenness, Closeness,
#                             Eigenvector, Prominence, Katz
#   Part C — Validation:      N concepts & relationships,
#                             R/T ratio, Avg path length,
#                             Diameter, Clustering, Density,
#                             Feedback Loops, Reinforcing Loops,
#                             Balancing Loops
#
# HOW TO RUN:
#   Place this file in the same folder as the Kumu .xlsx and
#   run: source("Network_Analysis_Galveston.R")
# ============================================================

# ── PACKAGES ──────────────────────────────────────────────────
pkgs <- c("readxl", "igraph", "dplyr", "ggplot2", "tidyr",
          "RColorBrewer", "scales", "openxlsx")
for (p in pkgs) {
  if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
  library(p, character.only = TRUE)
}

# ================================================================
# SECTION 0 — PARAMETERS  (only section you need to edit)
# ================================================================

DATA_FILE <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Exported_Kumu_Galveston_30April.xlsx"

# Fishing sectors to analyze individually (must match Tags exactly)
SECTORS   <- c("Charter", "Commercial", "Recreational")

OUT_DIR    <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Network_Outputs"
SAVE_PLOTS <- TRUE
SAVE_CSV   <- TRUE
# Also write Excel (.xlsx) copies of every table. Excel files store real
# NUMERIC cells, so decimals (e.g. R/T = 0.273) display correctly regardless
# of your computer's locale. Plain .csv files store "0.273" as text, and
# Excel set to a comma-decimal locale misreads the dot as a thousands
# separator — turning 0.273 into 273. Open the .xlsx files, not the .csv.
SAVE_XLSX  <- TRUE

# Helper: save a data frame as .csv and/or .xlsx (path given WITHOUT extension)
save_table <- function(df, path_noext) {
  if (SAVE_CSV)  utils::write.csv(df, paste0(path_noext, ".csv"), row.names = FALSE)
  if (SAVE_XLSX) openxlsx::write.xlsx(df, paste0(path_noext, ".xlsx"))
}

# ================================================================
# SECTION 1 — LOAD DATA
# ================================================================

elements    <- read_excel(DATA_FILE, sheet = "Elements")
connections <- read_excel(DATA_FILE, sheet = "Connections")

# Galveston uses Original_MM_Weight (FCM scale):
#  +1.0 = Strong Positive    +0.5 = Moderate Positive
#  -0.5 = Moderate Negative  -1.0 = Strong Negative
connections$fcm_weight <- as.numeric(connections$Original_MM_Weight)
connections <- connections[!is.na(connections$fcm_weight), ]

all_nodes <- elements$Label

if (!dir.exists(OUT_DIR)) dir.create(OUT_DIR, recursive = TRUE)

cat("Data loaded.\n")
cat(sprintf("Full model: %d concepts, %d connections\n",
            length(all_nodes), nrow(connections)))

# ================================================================
# SECTION 2 — HELPER FUNCTIONS
# ================================================================

build_wmat <- function(conn_df, nodes) {
  W <- matrix(0, length(nodes), length(nodes),
              dimnames = list(nodes, nodes))
  agg <- conn_df %>%
    filter(From %in% nodes, To %in% nodes) %>%
    group_by(From, To) %>%
    summarise(w = mean(fcm_weight), .groups = "drop")
  for (i in seq_len(nrow(agg))) W[agg$From[i], agg$To[i]] <- agg$w[i]
  W
}

filter_sector <- function(conn_df, sec) {
  conn_df[!is.na(conn_df$Tags) &
            grepl(sec, conn_df$Tags, fixed = TRUE), ]
}

# ================================================================
# SECTION 3 — CORE ANALYSIS FUNCTION
# ================================================================

run_analysis <- function(W, nodes, label) {
  
  n <- length(nodes)
  
  edge_df <- data.frame(
    from   = rep(nodes, times = n),
    to     = rep(nodes, each  = n),
    weight = as.vector(t(abs(W)))
  ) %>% filter(weight > 0)
  
  if (nrow(edge_df) == 0) {
    message(sprintf("[%s] No edges found — skipping.", label))
    return(NULL)
  }
  
  g <- graph_from_data_frame(edge_df, directed = TRUE, vertices = nodes)
  E(g)$weight <- edge_df$weight
  
  # Distance weights: 1/|w| so stronger links = shorter paths.
  # Used consistently for betweenness and closeness.
  E(g)$dist_weight <- 1 / E(g)$weight
  
  # ── PART A: CONCEPT CLASSIFICATION ──────────────────────────
  #
  # TRANSMITTER: no incoming connections, only outgoing.
  #   This concept is not influenced by any other concept in the
  #   map — it is an external driver or independent variable.
  #   Example: weather conditions, national regulations.
  #   Structural role: ROOT CAUSE / ENTRY POINT.
  #
  # RECEIVER: no outgoing connections, only incoming.
  #   This concept is affected by others but does not cascade
  #   further. It is a final endpoint or impact variable.
  #   Example: fisher wellbeing, economic loss.
  #   Structural role: OUTCOME / ENDPOINT.
  #
  # ORDINARY: has both incoming and outgoing connections.
  #   Mediated by upstream variables and itself mediates
  #   downstream ones. The vast majority of FCM nodes fall here.
  #   Structural role: MEDIATOR / INTERMEDIATE VARIABLE.
  #
  # ISOLATED: no connections at all.
  #   Rare. Possibly a concept added but never linked.
  
  indeg_w  <- colSums(abs(W))
  outdeg_w <- rowSums(abs(W))
  
  type <- dplyr::case_when(
    indeg_w  == 0 & outdeg_w >  0 ~ "Transmitter",
    indeg_w  >  0 & outdeg_w == 0 ~ "Receiver",
    indeg_w  == 0 & outdeg_w == 0 ~ "Isolated",
    TRUE                          ~ "Ordinary"
  )
  
  # ── PART B: CENTRALITY MEASURES ─────────────────────────────
  # All centrality measures use ABSOLUTE edge weights.
  # Signed weights (positive/negative) encode direction of effect,
  # but for structural importance we care about connection
  # STRENGTH, not direction. Signs are preserved separately in
  # the weight matrix for simulation work.
  
  # B1. INDEGREE CENTRALITY
  # Formula: ID(v) = Σ |w_uv| for all u pointing to v
  # Meaning: total strength of forces acting ON this concept.
  # High value → heavily influenced by many parts of the system.
  # Use: identify concepts most responsive to system-wide change.
  indegree <- indeg_w
  
  # B2. OUTDEGREE CENTRALITY
  # Formula: OD(v) = Σ |w_vu| for all u that v points to
  # Meaning: total strength of influence this concept exerts.
  # High value → strong driver of other concepts.
  # Use: identify key leverage points for management intervention.
  outdegree <- outdeg_w
  
  # B3. DEGREE CENTRALITY
  # Formula: D(v) = ID(v) + OD(v)
  # Meaning: total connection weight — both as cause and effect.
  # The most commonly used centrality measure in FCM literature.
  # Use: rank all concepts by overall network importance.
  degree <- indegree + outdegree
  
  # B4. BETWEENNESS CENTRALITY (weighted, normalised)
  # Formula: B(v) = Σ σ(s,t|v) / σ(s,t) over all pairs s≠t
  #   where σ(s,t) = total shortest paths from s to t,
  #         σ(s,t|v) = those passing through v.
  # Path lengths weighted by dist_weight = 1/|weight|
  # (stronger = shorter), following Opsahl et al. (2010).
  # High value → BOTTLENECK; removing it would most disrupt
  # system-wide information flow.
  # Use: identify critical bridge concepts for interventions.
  btw <- tryCatch(
    setNames(betweenness(g, weights = E(g)$dist_weight,
                         directed = TRUE, normalized = TRUE),
             V(g)$name),
    error = function(e) setNames(rep(0, n), nodes)
  )
  
  # B5. CLOSENESS CENTRALITY (outgoing, normalised)
  # igraph outgoing closeness, normalised by (N-1).
  # For each node: (N-1) / sum of shortest-path distances to all
  # reachable nodes. Unreachable pairs are excluded.
  # Distances weighted by dist_weight = 1/|weight|
  # (stronger = shorter), following Opsahl et al. (2010).
  # High value → fast-spreading concept; low value → isolated.
  # Complements betweenness: measures spread speed, not just
  # position on paths.
  clo <- tryCatch(
    setNames(closeness(g, weights = E(g)$dist_weight,
                       normalized = TRUE, mode = "out"),
             V(g)$name),
    error = function(e) setNames(rep(0, n), nodes)
  )
  
  # B6. EIGENVECTOR CENTRALITY
  # Formula: EC(v) = (1/λ) × Σ w_uv × EC(u) for all neighbors u
  # where λ is the largest eigenvalue of the weight matrix.
  # Being linked to other important nodes amplifies your own score.
  # Use: identify concepts embedded in dense high-centrality
  # neighborhoods — system-wide influential concepts.
  # Note: Returns zero for all nodes in purely acyclic graphs
  # (no feedback loops). Use Katz centrality in that case.
  eig <- tryCatch({
    ev <- eigen_centrality(g, directed = TRUE, weights = E(g)$weight)
    setNames(ev$vector, V(g)$name)
  }, error = function(e) setNames(rep(0, n), nodes))
  
  # B7. PROMINENCE (Hoffman et al., 2014)
  # Formula: P(c_i) = [f(c_i) + EC_norm(c_i)] / 2
  # where f(c_i) = 1 for all concepts in a fully aggregated model
  # (each concept appears in every participant's map),
  # and EC_norm is eigenvector centrality min-max normalised to [0,1].
  # Identifies concepts that are both structurally central
  # (high eigenvector) and frequently mentioned by participants.
  ec_vals    <- as.numeric(eig[nodes])
  ec_mn      <- min(ec_vals, na.rm = TRUE)
  ec_mx      <- max(ec_vals, na.rm = TRUE)
  norm_eigen <- if (ec_mx == ec_mn) rep(1, n) else (ec_vals - ec_mn) / (ec_mx - ec_mn)
  prominence <- (1 + norm_eigen) / 2   # occurrence = 1 for all concepts
  
  # B8. KATZ CENTRALITY
  # Formula: KC(v) = Σ_{k=1}^{∞} α^k (Aᵀ)^k · β
  # Simplified: k = β × (I − α·Aᵀ)⁻¹ · 1
  # Parameters:
  #   α = attenuation factor (auto-set to 0.85/spectral_radius
  #       to guarantee convergence; must be < 1/ρ)
  #   β = baseline importance per node (set to 1)
  # Like eigenvector centrality but every node starts with a
  # non-zero base score β. Preferred in small or acyclic
  # sub-group FCMs where eigenvector centrality may fail.
  katz <- tryCatch({
    ev_v <- eigen(t(abs(W)), only.values = TRUE)$values
    rho  <- max(Mod(ev_v))
    alp  <- 0.85 / max(rho, 1e-8)
    k    <- solve(diag(n) - alp * t(abs(W)), rep(1, n))
    k / max(abs(k))
  }, error = function(e) rep(0, n))
  
  concept_df <- data.frame(
    Concept     = nodes,
    Type        = type,
    Indegree    = round(indegree,                        4),
    Outdegree   = round(outdegree,                       4),
    Degree      = round(degree,                          4),
    Betweenness = round(as.numeric(btw[nodes]),          4),
    Closeness   = round(as.numeric(clo[nodes]),          4),
    Eigenvector = round(as.numeric(eig[nodes]),          4),
    Prominence  = round(prominence,                      4),
    Katz        = round(katz,                            4),
    stringsAsFactors = FALSE
  ) %>% arrange(desc(Degree))
  
  # ── PART C: VALIDATION METRICS ──────────────────────────────
  
  # C1. NUMBER OF CONCEPTS AND RELATIONSHIPS
  # N concepts (nodes): how many distinct ideas the group mapped.
  # N connections (edges): how many directed relationships.
  # Baseline comparison: are all sectors contributing similarly?
  n_concepts <- n
  n_edges    <- ecount(g)
  
  # C2. RECEIVER-TRANSMITTER (R/T) RATIO
  # R/T = |Receivers| / |Transmitters|
  # Interpretation:
  #   > 3     → many outcomes, few inputs (complex system view)
  #   1 – 3   → balanced (well-elaborated FCM)
  #   < 1     → many inputs, few outcomes (over-simplified)
  n_recv   <- sum(type == "Receiver")
  n_tran   <- sum(type == "Transmitter")
  rt_ratio <- if (n_tran > 0) round(n_recv / n_tran, 3) else NA
  
  # C3. SHORTEST PATH METRICS
  # Distance d(i→j): minimum sum of dist_weight = 1/|weight|
  # along any directed path from i to j.
  #
  # Average path length: mean across all REACHABLE pairs.
  # Diameter: the longest shortest path.
  # Per Table 1, Average Path Length is computed on UNWEIGHTED distance
  # ("the fewest number of relationships between them"), i.e. a hop count —
  # NOT the 1/|weight| distance used for betweenness/closeness. weights = NA
  # forces igraph to ignore edge weights and count edges. Only reachable,
  # directed pairs are used (Inf/unreachable excluded). Diameter is the
  # longest of those unweighted paths.
  D_unw    <- distances(g, mode = "out", weights = NA)
  finite_d <- D_unw[is.finite(D_unw) & D_unw > 0]
  avg_path <- if (length(finite_d) > 0) round(mean(finite_d), 4) else NA
  diam     <- if (length(finite_d) > 0) round(max(finite_d),  4) else NA
  
  # C4. CLUSTERING COEFFICIENT
  # Local (per node): proportion of a node's direct neighbors
  # that are also directly connected to each other.
  # Average: mean of local clustering across all nodes.
  # Global: proportion of "open triples" that are closed.
  clust_local  <- transitivity(g, type = "local", isolates = "zero")
  clust_avg    <- round(mean(clust_local, na.rm = TRUE), 4)
  clust_global <- round(transitivity(g, type = "global"),  4)
  
  # C5. DENSITY
  # Formula: |edges| / (n × (n−1))
  # Low density = participants were selective in adding links.
  # Density naturally decreases as n grows, so compare only
  # across networks of similar size.
  density_val <- round(ecount(g) / (n * (n - 1)), 4)
  
  # C5b. LINK POLARITY (Table 1)
  # Percentage of causal links that are positive (w > 0) vs negative
  # (w < 0), using the SIGNED aggregated weight matrix W. A pair whose
  # averaged weight is exactly 0 is not an edge and is excluded, so
  # pos + neg always equals the reported edge count.
  pos_links <- sum(W > 0)
  neg_links <- sum(W < 0)
  tot_links <- pos_links + neg_links
  pct_pos   <- if (tot_links > 0) round(100 * pos_links / tot_links, 1) else NA
  pct_neg   <- if (tot_links > 0) round(100 * neg_links / tot_links, 1) else NA
  
  # C6. FEEDBACK LOOPS (up to length 7)
  # Feedback loops are directed cycles in the FCM graph.
  # Loop POLARITY is determined by the product of the signs of
  # edge weights around the cycle, using the original signed W.
  #   Product > 0 → Reinforcing (positive feedback) loop
  #   Product < 0 → Balancing (negative feedback) loop
  # Only cycles of length ≤ 7 are enumerated (computationally
  # tractable; longer loops have negligible dynamic effect).
  cycles_all <- tryCatch(igraph::simple_cycles(g),
                         error = function(e) list())
  cycles     <- Filter(function(x) length(x) <= 7, cycles_all)
  n_cycles   <- length(cycles)
  
  loop_pol <- if (n_cycles > 0) {
    sapply(cycles, function(cyc) {
      v_seq <- V(g)$name[cyc]
      n_c   <- length(v_seq)
      prod(sapply(seq_len(n_c), function(i) {
        j <- if (i == n_c) 1L else i + 1L
        sign(W[v_seq[i], v_seq[j]])
      }))
    })
  } else numeric(0)
  
  n_reinf <- sum(loop_pol > 0)
  n_bal   <- sum(loop_pol < 0)
  
  n_ord  <- sum(type == "Ordinary")
  n_isol <- sum(type == "Isolated")
  
  validation_df <- data.frame(
    Metric = c(
      "Concepts (nodes)",
      "Connections (edges, aggregated)",
      "Transmitters",
      "Receivers",
      "Ordinary",
      "Isolated",
      "R/T Ratio",
      "Average Path Length",
      "Diameter",
      "Average Clustering Coefficient",
      "Global Clustering Coefficient",
      "Density",
      "Positive Links (%)",
      "Negative Links (%)",
      "Feedback Loops (≤7 nodes)",
      "Reinforcing Loops",
      "Balancing Loops"
    ),
    Value = c(
      n_concepts, n_edges,
      n_tran, n_recv, n_ord, n_isol,
      rt_ratio,
      avg_path, diam,
      clust_avg, clust_global,
      density_val,
      pct_pos, pct_neg,
      n_cycles, n_reinf, n_bal
    ),
    stringsAsFactors = FALSE
  )
  
  list(
    label         = label,
    n_nodes       = n,
    n_edges       = n_edges,
    concept_df    = concept_df,
    validation_df = validation_df,
    graph         = g
  )
}

# ================================================================
# SECTION 4 — RUN ANALYSIS: FULL MODEL
# ================================================================

cat("\n")
cat("================================================================\n")
cat("FULL MODEL — GALVESTON (all 3 sectors combined)\n")
cat("================================================================\n")

W_full   <- build_wmat(connections, all_nodes)
res_full <- run_analysis(W_full, all_nodes, "Galveston (full)")

cat("\n--- PART A: Concept Classification ---\n")
cat("Transmitter = pure input; Receiver = pure outcome; Ordinary = mediator\n\n")
print(table(res_full$concept_df$Type))

cat("\n--- Transmitters ---\n")
print(res_full$concept_df[res_full$concept_df$Type == "Transmitter",
                          c("Concept","Outdegree")], row.names = FALSE)

cat("\n--- Receivers ---\n")
print(res_full$concept_df[res_full$concept_df$Type == "Receiver",
                          c("Concept","Indegree")], row.names = FALSE)

cat("\n--- PART B: Centrality — Top 20 by Degree ---\n")
print(head(res_full$concept_df, 20), row.names = FALSE)

cat("\n--- PART C: Validation Metrics ---\n")
print(res_full$validation_df, row.names = FALSE)

save_table(res_full$concept_df,    file.path(OUT_DIR, "GA_full_centrality"))
save_table(res_full$validation_df, file.path(OUT_DIR, "GA_full_validation"))

# ================================================================
# SECTION 5 — RUN ANALYSIS: EACH SECTOR
# ================================================================

cat("\n")
cat("================================================================\n")
cat("SECTOR ANALYSES\n")
cat("================================================================\n")

sub_results <- list()

for (sec in SECTORS) {
  
  cat(sprintf("\n--- Sector: %s ---\n", sec))
  
  conn_s  <- filter_sector(connections, sec)
  nodes_s <- unique(c(conn_s$From, conn_s$To))
  nodes_s <- nodes_s[nodes_s %in% all_nodes]
  
  if (length(nodes_s) < 2) {
    cat(sprintf("  Fewer than 2 nodes in %s — skipping.\n", sec))
    next
  }
  
  W_s   <- build_wmat(conn_s, nodes_s)
  res_s <- run_analysis(W_s, nodes_s, sec)
  if (is.null(res_s)) next
  sub_results[[sec]] <- res_s
  
  cat(sprintf("  Concepts: %d  |  Connections: %d\n",
              res_s$n_nodes, res_s$n_edges))
  cat("  Concept types: ")
  print(table(res_s$concept_df$Type))
  
  cat("  Top 10 by Degree:\n")
  print(head(res_s$concept_df[, c("Concept","Type","Degree",
                                  "Betweenness","Katz")], 10),
        row.names = FALSE)
  
  cat("  Validation:\n")
  print(res_s$validation_df, row.names = FALSE)
  
  safe_name <- gsub(" ", "_", sec)
  save_table(res_s$concept_df,
             file.path(OUT_DIR, sprintf("GA_%s_centrality", safe_name)))
  save_table(res_s$validation_df,
             file.path(OUT_DIR, sprintf("GA_%s_validation", safe_name)))
}

# ================================================================
# SECTION 6 — COMPARISON TABLE
# ================================================================

cat("\n")
cat("================================================================\n")
cat("COMPARISON: FULL MODEL vs SECTORS\n")
cat("================================================================\n")

all_results <- c(list("Galveston (full)" = res_full), sub_results)

compare_df <- do.call(rbind, lapply(names(all_results), function(nm) {
  vdf <- all_results[[nm]]$validation_df
  row <- setNames(as.list(vdf$Value), vdf$Metric)
  data.frame(Level = nm, as.data.frame(row, check.names = FALSE,
                                       stringsAsFactors = FALSE))
}))

key_cols <- c("Level", "Concepts (nodes)", "Connections (edges, aggregated)",
              "Transmitters", "Receivers", "R/T Ratio",
              "Average Path Length", "Diameter",
              "Average Clustering Coefficient", "Density",
              "Positive Links (%)", "Negative Links (%)",
              "Feedback Loops (≤7 nodes)", "Reinforcing Loops", "Balancing Loops")
key_cols <- key_cols[key_cols %in% names(compare_df)]
cat("\nKey metrics:\n")
print(compare_df[, key_cols], row.names = FALSE)

cat("\n--- Top 5 Concepts by Degree per Level ---\n")
for (nm in names(all_results)) {
  top5 <- head(all_results[[nm]]$concept_df$Concept, 5)
  cat(sprintf("  %-22s: %s\n", nm, paste(top5, collapse = ", ")))
}

save_table(compare_df, file.path(OUT_DIR, "GA_comparison_table"))


cat(sprintf("\nAll outputs saved to: %s\n", OUT_DIR))
cat("Analysis complete.\n")