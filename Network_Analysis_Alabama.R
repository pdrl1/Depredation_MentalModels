# ============================================================
# NETWORK STRUCTURE ANALYSIS — ALABAMA GULF COAST
# ============================================================
# Analyzes the full Alabama FCM and each of its five
# state-group sub-models as individual networks.
#
# LEVELS ANALYZED:
#   1. Gulf Coast (full model — all 5 state groups combined)
#   2. Alabama (state)
#   3. Mississippi
#   4. Louisiana
#   5. Florida
#   6. Texas
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
#   run: source("Network_Analysis_Alabama.R")
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

DATA_FILE  <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Alabama/Kumu/Exported_Kumu_Alabama_30April.xlsx"

# State groups to analyze individually (must match Tags exactly)
SUBGROUPS  <- c("Alabama", "Mississippi", "Louisiana",
                "Florida",  "Texas")

OUT_DIR    <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Alabama/Network_Outputs"
SAVE_PLOTS <- TRUE
SAVE_CSV   <- TRUE

# ================================================================
# SECTION 1 — LOAD DATA
# ================================================================

elements    <- read_excel(DATA_FILE, sheet = "Elements")
connections <- read_excel(DATA_FILE, sheet = "Connections")

# Alabama uses Original_MM_Weight (already in FCM scale):
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

filter_group <- function(conn_df, grp) {
  conn_df[!is.na(conn_df$Tags) &
            grepl(grp, conn_df$Tags, fixed = TRUE), ]
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
  # TRANSMITTER: has outgoing connections, NO incoming.
  #   The concept influences others but nothing in the model
  #   influences it. Typically an external driver: e.g., weather,
  #   regulatory context, shark life history traits.
  #   → Pure input / root cause.
  #
  # RECEIVER: has incoming connections, NO outgoing.
  #   The concept is affected by others but exerts no further
  #   influence. Typically a final outcome: e.g., fisher income,
  #   wellbeing, ecological population change.
  #   → Pure outcome / endpoint.
  #
  # ORDINARY: has BOTH incoming and outgoing connections.
  #   A mediating variable that is both caused and causal.
  #   Most FCM nodes are ordinary.
  #   → Mediator / intermediate variable.
  #
  # ISOLATED: no connections at all. Rare; may indicate a data
  #   entry issue or a concept that was listed but not linked.
  
  indeg_w  <- colSums(abs(W))
  outdeg_w <- rowSums(abs(W))
  
  type <- dplyr::case_when(
    indeg_w  == 0 & outdeg_w >  0 ~ "Transmitter",
    indeg_w  >  0 & outdeg_w == 0 ~ "Receiver",
    indeg_w  == 0 & outdeg_w == 0 ~ "Isolated",
    TRUE                          ~ "Ordinary"
  )
  
  # ── PART B: CENTRALITY MEASURES ─────────────────────────────
  
  # B1. INDEGREE CENTRALITY
  # = sum of |incoming edge weights|
  # Measures how INFLUENCED a concept is by the rest of the
  # network. High indegree → many or strong forces act on it.
  # Receivers have the highest indegrees among outcome concepts.
  indegree <- indeg_w
  
  # B2. OUTDEGREE CENTRALITY
  # = sum of |outgoing edge weights|
  # Measures how much a concept INFLUENCES others.
  # High outdegree → strong driver of the system.
  # Transmitters have the highest outdegrees among input concepts.
  outdegree <- outdeg_w
  
  # B3. DEGREE CENTRALITY
  # = indegree + outdegree
  # Total connection weight — measures overall embeddedness.
  # The most commonly reported FCM centrality measure.
  # High degree = key concept regardless of causal direction.
  degree <- indegree + outdegree
  
  # B4. BETWEENNESS CENTRALITY
  # = fraction of all shortest paths (between any pair of nodes)
  #   that pass through this node.
  # Weighted version: path length = sum of 1/|weight| along edges
  # (stronger link = shorter = more preferred route).
  # High betweenness → bottleneck / bridge concept.
  # Removing it would most disrupt information flow.
  # Practical use: identify leverage points for management.
  btw <- tryCatch(
    setNames(betweenness(g, weights = 1 / E(g)$weight,
                         directed = TRUE, normalized = TRUE),
             V(g)$name),
    error = function(e) setNames(rep(0, n), nodes)
  )
  
  # B5. CLOSENESS CENTRALITY (harmonic)
  # = mean of (1 / distance) to all reachable nodes
  # Using harmonic mean handles disconnected graphs: unreachable
  # nodes contribute 0 rather than inflating the average.
  # High closeness → concept can reach (or be reached from) all
  # others quickly → fast propagation of its effects.
  # Complements betweenness: measures spread speed, not just
  # position on paths.
  closeness <- sapply(nodes, function(v) {
    d <- distances(g, v = v, to = nodes,
                   weights = 1 / E(g)$weight, mode = "out")
    d <- d[d > 0 & is.finite(d)]
    if (length(d) == 0) return(0)
    mean(1 / d)
  })
  
  # B6. EIGENVECTOR CENTRALITY
  # = importance score that weights a node's connections by the
  #   importance of its neighbors (recursive definition).
  # Solved as the principal eigenvector of the weight matrix.
  # Being connected to OTHER central nodes amplifies your score.
  # Caution: returns all zeros for purely acyclic networks
  # (i.e., no feedback loops). Use Katz in those cases.
  eig <- tryCatch({
    ev <- eigen_centrality(g, directed = TRUE, weights = E(g)$weight)
    setNames(ev$vector, V(g)$name)
  }, error = function(e) setNames(rep(0, n), nodes))
  
  # B7. KATZ CENTRALITY
  # Extends eigenvector centrality by adding a small baseline
  # importance β to every node (default β=1), preventing zero
  # scores in acyclic sub-graphs.
  # α (attenuation factor) is auto-set to 0.85/spectral_radius
  # to guarantee convergence of the series.
  # Best practice: report alongside eigenvector; if they differ
  # substantially, the network has significant acyclic regions.
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
  
  # C1. CONCEPT AND RELATIONSHIP COUNTS
  # Basic size descriptors. Compare across groups to check
  # whether all state sub-groups contributed proportionally.
  n_concepts <- n
  n_edges    <- ecount(g)
  
  # C2. RECEIVER-TRANSMITTER (R/T) RATIO = |Receivers| / |Transmitters|
  # Measures the input-output balance of the map.
  #   R/T >> 1 → many outcomes, few drivers → rich causal chains.
  #   R/T << 1 → many drivers, few outcomes → overly simple.
  #   R/T ≈ 1–3 → well-balanced (typical for good FCMs).
  n_recv   <- sum(type == "Receiver")
  n_tran   <- sum(type == "Transmitter")
  rt_ratio <- if (n_tran > 0) round(n_recv / n_tran, 3) else NA
  
  # C3. SHORTEST PATH METRICS
  # All distances weighted by 1/|weight| (stronger = shorter).
  #
  # Average path length: mean finite pairwise distance.
  #   Small → influence travels quickly across the whole map.
  #
  # Diameter: longest shortest path in the network.
  #   Indicates the maximum number of "hops" in a causal chain.
  #   Very small diameter relative to n may signal that
  #   participants did not elaborate complex causal sequences.
  g_dist <- g
  E(g_dist)$weight <- 1 / E(g)$weight
  avg_path <- round(mean_distance(g_dist, directed = TRUE,
                                  unconnected = TRUE), 4)
  diam     <- round(diameter(g_dist, directed = TRUE,
                             unconnected = TRUE), 4)
  
  # C4. CLUSTERING COEFFICIENT
  # Local clustering: for each node, fraction of its neighbors
  # that are also directly connected to each other.
  # Average local clustering: mean across all nodes.
  # Global clustering: fraction of "open triples" closed
  # into triangles across the entire network.
  # Low clustering = participants added relationships selectively
  # (desirable). High clustering may indicate over-elaboration.
  clust_local  <- transitivity(g, type = "local", isolates = "zero")
  clust_avg    <- round(mean(clust_local, na.rm = TRUE), 4)
  clust_global <- round(transitivity(g, type = "global"),  4)
  
  # C5. DENSITY = |edges| / (n × (n-1))
  # Maximum possible directed edges in a network of n nodes
  # is n×(n-1) (excluding self-loops).
  # Low density → participants were selective (desirable).
  # Density naturally decreases as n grows.
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
cat("FULL MODEL — ALABAMA GULF COAST (all 5 state groups)\n")
cat("================================================================\n")

W_full   <- build_wmat(connections, all_nodes)
res_full <- run_analysis(W_full, all_nodes, "Gulf Coast (full)")

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
            file.path(OUT_DIR, "AL_full_centrality.csv"), row.names = FALSE)
  write.csv(res_full$validation_df,
            file.path(OUT_DIR, "AL_full_validation.csv"), row.names = FALSE)
}

# ================================================================
# SECTION 5 — RUN ANALYSIS: EACH STATE GROUP
# ================================================================

cat("\n")
cat("================================================================\n")
cat("STATE-GROUP ANALYSES\n")
cat("================================================================\n")

sub_results <- list()

for (grp in SUBGROUPS) {
  
  cat(sprintf("\n--- State group: %s ---\n", grp))
  
  conn_g  <- filter_group(connections, grp)
  nodes_g <- unique(c(conn_g$From, conn_g$To))
  nodes_g <- nodes_g[nodes_g %in% all_nodes]
  
  if (length(nodes_g) < 2) {
    cat(sprintf("  Fewer than 2 nodes in %s — skipping.\n", grp))
    next
  }
  
  W_g   <- build_wmat(conn_g, nodes_g)
  res_g <- run_analysis(W_g, nodes_g, grp)
  if (is.null(res_g)) next
  sub_results[[grp]] <- res_g
  
  cat(sprintf("  Concepts: %d  |  Connections: %d\n",
              res_g$n_nodes, res_g$n_edges))
  cat("  Concept types: ")
  print(table(res_g$concept_df$Type))
  
  cat("  Top 10 by Degree:\n")
  print(head(res_g$concept_df[, c("Concept","Type","Degree",
                                  "Betweenness","Katz")], 10),
        row.names = FALSE)
  
  cat("  Validation:\n")
  print(res_g$validation_df, row.names = FALSE)
  
  if (SAVE_CSV) {
    write.csv(res_g$concept_df,
              file.path(OUT_DIR, sprintf("AL_%s_centrality.csv", grp)),
              row.names = FALSE)
    write.csv(res_g$validation_df,
              file.path(OUT_DIR, sprintf("AL_%s_validation.csv", grp)),
              row.names = FALSE)
  }
}

# ================================================================
# SECTION 6 — COMPARISON TABLE
# ================================================================

cat("\n")
cat("================================================================\n")
cat("COMPARISON: FULL MODEL vs STATE GROUPS\n")
cat("================================================================\n")

all_results <- c(list("Gulf Coast (full)" = res_full), sub_results)

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
                        file.path(OUT_DIR, "AL_comparison_table.csv"),
                        row.names = FALSE)

# ================================================================
# SECTION 7 — VISUALIZATIONS
# ================================================================

cat("\n--- Generating plots ---\n")

model_colours <- c(
  "Gulf Coast (full)" = "#1B9E77",
  "Alabama"           = "#D95F02",
  "Mississippi"       = "#7570B3",
  "Louisiana"         = "#E7298A",
  "Florida"           = "#66A61E",
  "Texas"             = "#E6AB02"
)

# ── Plot 1: Degree centrality — full model, top 25 ────────────
p1 <- res_full$concept_df %>%
  head(25) %>%
  ggplot(aes(x = reorder(Concept, Degree), y = Degree, fill = Type)) +
  geom_col() +
  coord_flip() +
  scale_fill_brewer(palette = "Set2") +
  labs(title    = "Degree Centrality — Alabama Gulf Coast (Full Model)",
       subtitle = "Top 25 concepts",
       x = NULL, y = "Degree") +
  theme_bw(base_size = 9)

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "01_AL_full_degree.pdf"),
                       p1, width = 10, height = 8)
print(p1)

# ── Plot 2: Top-3 per level ───────────────────────────────────
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
  scale_fill_manual(values = model_colours) +
  labs(title = "Top 3 by Degree — Alabama and State Groups",
       x = "Rank", y = "Degree Centrality") +
  theme_bw(base_size = 8) +
  theme(legend.position = "none",
        strip.text = element_text(size = 7, face = "bold"))

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "02_AL_subgroup_top3.pdf"),
                       p2, width = 13, height = 8)
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
  scale_fill_manual(values = model_colours) +
  labs(title = "Validation Metrics — Alabama and State Groups",
       x = NULL, y = "Value") +
  theme_bw(base_size = 9) +
  theme(axis.text.x = element_text(angle = 35, hjust = 1, size = 7),
        strip.text  = element_text(face = "bold"))

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "03_AL_validation_comparison.pdf"),
                       p3, width = 10, height = 5)
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
  labs(title    = "Betweenness Centrality — Alabama Gulf Coast (Full Model)",
       subtitle  = "High value = key bottleneck for information flow",
       x = NULL, y = "Betweenness (normalised)") +
  theme_bw(base_size = 9)

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "04_AL_full_betweenness.pdf"),
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
  labs(title = "Concept Type Composition — Alabama and State Groups",
       x = NULL, y = "Proportion") +
  theme_bw(base_size = 9) +
  theme(axis.text.x = element_text(angle = 30, hjust = 1))

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "05_AL_type_composition.pdf"),
                       p5, width = 9, height = 5)
print(p5)

cat(sprintf("\nAll outputs saved to: %s\n", OUT_DIR))
cat("Analysis complete.\n")