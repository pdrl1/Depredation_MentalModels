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
#                             Eigenvector, Katz
#   Part C — Validation:      N concepts & relationships,
#                             R/T ratio, Avg path length,
#                             Diameter, Clustering, Density
#
# HOW TO RUN:
#   Place this file in the same folder as the Kumu .xlsx and
#   run: source("Network_Analysis_Galveston.R")
# ============================================================

# ── PACKAGES ──────────────────────────────────────────────────
pkgs <- c("readxl", "igraph", "dplyr", "ggplot2", "tidyr",
          "RColorBrewer", "scales")
for (p in pkgs) {
  if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
  library(p, character.only = TRUE)
}

# ================================================================
# SECTION 0 — PARAMETERS  (only section you need to edit)
# ================================================================

DATA_FILE <- "Exported_kumu_Galveston_22April.xlsx"

# Fishing sectors to analyze individually (must match Tags exactly)
SECTORS   <- c("Charter", "Commercial", "Recreational")

OUT_DIR    <- "network_outputs_galveston"
SAVE_PLOTS <- TRUE
SAVE_CSV   <- TRUE

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
  
  # B4. BETWEENNESS CENTRALITY (normalised)
  # Formula: B(v) = Σ σ(s,t|v) / σ(s,t) over all pairs s≠t
  #   where σ(s,t) = total shortest paths from s to t,
  #         σ(s,t|v) = those passing through v.
  # Path lengths weighted by 1/|weight| (stronger = shorter).
  # Meaning: how often this concept lies on the most efficient
  # route between any two other concepts in the network.
  # High value → BOTTLENECK; removing it would most disrupt
  # system-wide information flow.
  # Use: identify critical bridge concepts for interventions.
  btw <- tryCatch(
    setNames(betweenness(g, weights = 1 / E(g)$weight,
                         directed = TRUE, normalized = TRUE),
             V(g)$name),
    error = function(e) setNames(rep(0, n), nodes)
  )
  
  # B5. CLOSENESS CENTRALITY (harmonic form)
  # Formula: HC(v) = (1/n-1) × Σ 1/d(v,u) for all reachable u≠v
  # Standard closeness (1/mean distance) breaks for disconnected
  # graphs; harmonic closeness handles this by ignoring
  # unreachable nodes (they contribute 0 to the sum).
  # Meaning: how quickly this concept can propagate its effects
  # through the network via directed shortest paths.
  # High value → fast-spreading concept; low value → isolated.
  closeness <- sapply(nodes, function(v) {
    d <- distances(g, v = v, to = nodes,
                   weights = 1 / E(g)$weight, mode = "out")
    d <- d[d > 0 & is.finite(d)]
    if (length(d) == 0) return(0)
    mean(1 / d)
  })
  
  # B6. EIGENVECTOR CENTRALITY
  # Formula: EC(v) = (1/λ) × Σ w_uv × EC(u) for all neighbors u
  # where λ is the largest eigenvalue of the weight matrix.
  # Meaning: a concept's importance is proportional to the
  # importance of the concepts connected to it. Being linked to
  # other important nodes amplifies your own score.
  # Use: identify concepts embedded in dense high-centrality
  # neighborhoods — system-wide influential concepts.
  # Note: Returns zero for all nodes in purely acyclic graphs
  # (no feedback loops). Use Katz centrality in that case.
  eig <- tryCatch({
    ev <- eigen_centrality(g, directed = TRUE, weights = E(g)$weight)
    setNames(ev$vector, V(g)$name)
  }, error = function(e) setNames(rep(0, n), nodes))
  
  # B7. KATZ CENTRALITY
  # Formula: KC(v) = Σ_{k=1}^{∞} α^k (Aᵀ)^k · β
  # Simplified: k = β × (I − α·Aᵀ)⁻¹ · 1
  # Parameters:
  #   α = attenuation factor (auto-set to 0.85/spectral_radius
  #       to guarantee convergence; must be < 1/ρ)
  #   β = baseline importance per node (set to 1)
  # Meaning: like eigenvector centrality but every node starts
  # with a non-zero base score β. This prevents the "dead end"
  # problem where nodes with no feedback loops score zero.
  # Use: preferred in small or acyclic sub-group FCMs where
  # eigenvector centrality may fail.
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
    Indegree    = round(indegree,  4),
    Outdegree   = round(outdegree, 4),
    Degree      = round(degree,    4),
    Betweenness = round(as.numeric(btw[nodes]),  4),
    Closeness   = round(closeness,               4),
    Eigenvector = round(as.numeric(eig[nodes]),  4),
    Katz        = round(katz,                    4),
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
  #   ∞ (0 transmitters) → all inputs are mediated through
  #             ordinary concepts; may indicate no clear external
  #             drivers identified by this group.
  n_recv   <- sum(type == "Receiver")
  n_tran   <- sum(type == "Transmitter")
  rt_ratio <- if (n_tran > 0) round(n_recv / n_tran, 3) else NA
  
  # C3. SHORTEST PATH METRICS
  # Distance d(i→j): minimum sum of 1/|weight| along any
  # directed path from i to j. Shorter = more direct influence.
  #
  # Average path length: mean across all REACHABLE pairs.
  #   Small → effects cascade quickly across the whole map.
  #   Large → map is chain-like; influence takes many steps.
  #
  # Diameter: the longest shortest path.
  #   Tells you the worst-case depth of a causal chain.
  #   Diameter much smaller than n can indicate the group did
  #   not fully elaborate multi-step causal reasoning.
  g_dist <- g
  E(g_dist)$weight <- 1 / E(g)$weight
  avg_path <- round(mean_distance(g_dist, directed = TRUE,
                                  unconnected = TRUE), 4)
  diam     <- round(diameter(g_dist, directed = TRUE,
                             unconnected = TRUE), 4)
  
  # C4. CLUSTERING COEFFICIENT
  # Local (per node): proportion of a node's direct neighbors
  # that are also directly connected to each other.
  # Average: mean of local clustering across all nodes.
  # Global: proportion of "open triples" (A→B, B→C) that are
  # closed (A→C also exists), computed globally.
  #
  # In FCM context:
  #   High clustering → participants added many triangle-like
  #   relationships; may indicate excessive connections.
  #   Low clustering  → relationships are more chain-like;
  #   participants were selective → generally desirable.
  clust_local  <- transitivity(g, type = "local", isolates = "zero")
  clust_avg    <- round(mean(clust_local, na.rm = TRUE), 4)
  clust_global <- round(transitivity(g, type = "global"),  4)
  
  # C5. DENSITY
  # Formula: |edges| / (n × (n−1))
  # Maximum possible edges in a directed graph without self-loops
  # is n×(n−1) (every node connects to every other).
  # Low density = participants were selective in adding links.
  # This is desirable: not everything influences everything else.
  # Density naturally decreases as n grows, so compare only
  # across networks of similar size.
  density_val <- round(ecount(g) / (n * (n - 1)), 4)
  
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
      "Density"
    ),
    Value = c(
      n_concepts, n_edges,
      n_tran, n_recv, n_ord, n_isol,
      rt_ratio,
      avg_path, diam,
      clust_avg, clust_global,
      density_val
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

if (SAVE_CSV) {
  write.csv(res_full$concept_df,
            file.path(OUT_DIR, "GA_full_centrality.csv"), row.names = FALSE)
  write.csv(res_full$validation_df,
            file.path(OUT_DIR, "GA_full_validation.csv"), row.names = FALSE)
}

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
  
  if (SAVE_CSV) {
    write.csv(res_s$concept_df,
              file.path(OUT_DIR, sprintf("GA_%s_centrality.csv", sec)),
              row.names = FALSE)
    write.csv(res_s$validation_df,
              file.path(OUT_DIR, sprintf("GA_%s_validation.csv", sec)),
              row.names = FALSE)
  }
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
              "Average Clustering Coefficient", "Density")
key_cols <- key_cols[key_cols %in% names(compare_df)]
cat("\nKey metrics:\n")
print(compare_df[, key_cols], row.names = FALSE)

cat("\n--- Top 5 Concepts by Degree per Level ---\n")
for (nm in names(all_results)) {
  top5 <- head(all_results[[nm]]$concept_df$Concept, 5)
  cat(sprintf("  %-22s: %s\n", nm, paste(top5, collapse = ", ")))
}

if (SAVE_CSV) write.csv(compare_df,
                        file.path(OUT_DIR, "GA_comparison_table.csv"),
                        row.names = FALSE)

# ================================================================
# SECTION 7 — VISUALIZATIONS
# ================================================================

cat("\n--- Generating plots ---\n")

sector_colours <- c(
  "Galveston (full)" = "#1B9E77",
  "Charter"          = "#D95F02",
  "Commercial"       = "#7570B3",
  "Recreational"     = "#E7298A"
)

# ── Plot 1: Degree centrality — full model, top 25 ────────────
p1 <- res_full$concept_df %>%
  head(25) %>%
  ggplot(aes(x = reorder(Concept, Degree), y = Degree, fill = Type)) +
  geom_col() +
  coord_flip() +
  scale_fill_brewer(palette = "Set2") +
  labs(title    = "Degree Centrality — Galveston (Full Model)",
       subtitle = "Top 25 concepts",
       x = NULL, y = "Degree") +
  theme_bw(base_size = 9)

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "01_GA_full_degree.pdf"),
                       p1, width = 10, height = 8)
print(p1)

# ── Plot 2: Top-3 per sector ──────────────────────────────────
top3_df <- do.call(rbind, lapply(names(all_results), function(nm) {
  head(all_results[[nm]]$concept_df[, c("Concept","Degree")], 3) %>%
    mutate(Level = nm, Rank = row_number())
}))

p2 <- ggplot(top3_df, aes(x = Rank, y = Degree,
                          fill = Level, label = Concept)) +
  geom_col(position = "dodge") +
  geom_text(position = position_dodge(0.9), hjust = -0.1,
            size = 2.5, angle = 0) +
  facet_wrap(~Level, scales = "free_y") +
  coord_flip() +
  scale_fill_manual(values = sector_colours) +
  labs(title = "Top 3 by Degree — Galveston and Sectors",
       x = "Rank", y = "Degree Centrality") +
  theme_bw(base_size = 8) +
  theme(legend.position = "none",
        strip.text = element_text(size = 7, face = "bold"))

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "02_GA_sector_top3.pdf"),
                       p2, width = 11, height = 7)
print(p2)

# ── Plot 3: Validation comparison ─────────────────────────────
val_long <- compare_df %>%
  select(Level, Density,
         `Average Clustering Coefficient`, `R/T Ratio`) %>%
  pivot_longer(-Level, names_to = "Metric", values_to = "Value") %>%
  mutate(Value = as.numeric(Value))

p3 <- ggplot(val_long, aes(x = Level, y = Value, fill = Level)) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~Metric, scales = "free_y") +
  scale_fill_manual(values = sector_colours) +
  labs(title = "Validation Metrics — Galveston and Sectors",
       x = NULL, y = "Value") +
  theme_bw(base_size = 9) +
  theme(axis.text.x = element_text(angle = 25, hjust = 1, size = 8),
        strip.text  = element_text(face = "bold"))

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "03_GA_validation_comparison.pdf"),
                       p3, width = 9, height = 5)
print(p3)

# ── Plot 4: Betweenness — full model ──────────────────────────
p4 <- res_full$concept_df %>%
  arrange(desc(Betweenness)) %>%
  head(20) %>%
  ggplot(aes(x = reorder(Concept, Betweenness),
             y = Betweenness, fill = Type)) +
  geom_col() +
  coord_flip() +
  scale_fill_brewer(palette = "Set2") +
  labs(title    = "Betweenness Centrality — Galveston (Full Model)",
       subtitle  = "High value = key bottleneck for information flow",
       x = NULL, y = "Betweenness (normalised)") +
  theme_bw(base_size = 9)

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "04_GA_full_betweenness.pdf"),
                       p4, width = 10, height = 7)
print(p4)

# ── Plot 5: Concept type composition ─────────────────────────
type_df <- do.call(rbind, lapply(names(all_results), function(nm) {
  tab <- table(all_results[[nm]]$concept_df$Type)
  data.frame(Level = nm, Type = names(tab), Count = as.integer(tab))
}))

p5 <- ggplot(type_df, aes(x = Level, y = Count, fill = Type)) +
  geom_col(position = "fill") +
  scale_y_continuous(labels = percent) +
  scale_fill_brewer(palette = "Set2") +
  labs(title = "Concept Type Composition — Galveston and Sectors",
       x = NULL, y = "Proportion") +
  theme_bw(base_size = 9) +
  theme(axis.text.x = element_text(angle = 20, hjust = 1))

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "05_GA_type_composition.pdf"),
                       p5, width = 8, height = 5)
print(p5)

# ── Plot 6: All centrality measures — side by side (full model)
cent_long <- res_full$concept_df %>%
  head(15) %>%
  select(Concept, Degree, Betweenness, Closeness, Eigenvector, Katz) %>%
  pivot_longer(-Concept, names_to = "Measure", values_to = "Score") %>%
  mutate(Measure = factor(Measure,
                          levels = c("Degree","Betweenness","Closeness","Eigenvector","Katz")))

p6 <- ggplot(cent_long,
             aes(x = reorder(Concept, Score), y = Score, fill = Measure)) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~Measure, scales = "free_x", ncol = 5) +
  coord_flip() +
  scale_fill_brewer(palette = "Dark2") +
  labs(title    = "All Centrality Measures — Galveston Full Model (Top 15 by Degree)",
       subtitle = "Each panel uses its own x-axis scale",
       x = NULL, y = "Score") +
  theme_bw(base_size = 8) +
  theme(strip.text = element_text(face = "bold", size = 8))

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "06_GA_all_centrality.pdf"),
                       p6, width = 16, height = 6)
print(p6)

cat(sprintf("\nAll outputs saved to: %s\n", OUT_DIR))
cat("Analysis complete.\n")