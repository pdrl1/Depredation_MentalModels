# ============================================================
# NETWORK STRUCTURE ANALYSIS — AUSTRALIA
# ============================================================
# Analyzes the full Australia FCM and each of its four
# regional sub-models as individual networks.
#
# LEVELS ANALYZED:
#   1. Australia (full model — all 4 regions combined)
#   2. New South Wales
#   3. Queensland
#   4. North Australia
#   5. Western Australia
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
#   run: source("Network_Analysis_Australia.R")
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

DATA_FILE  <- "Exported_Kumu_Australia_30April.xlsx"

# Sub-regions to analyze individually (must match Tags exactly)
SUBREGIONS <- c("New South Wales", "Queensland",
                "North Australia", "Western Australia")

# Output folder for plots and CSV files
OUT_DIR    <- "network_outputs_australia"
SAVE_PLOTS <- TRUE
SAVE_CSV   <- TRUE

# ================================================================
# SECTION 1 — LOAD DATA
# ================================================================

elements    <- read_excel(DATA_FILE, sheet = "Elements")
connections <- read_excel(DATA_FILE, sheet = "Connections")

# Australia uses Strength codes. Map to FCM weights:
#   2 → +1.0 (Strong Positive)    1 → +0.5 (Moderate Positive)
#  -1 → -0.5 (Moderate Negative) -2 → -1.0 (Strong Negative)
str_map <- c("2" = 1.0, "1" = 0.5, "-1" = -0.5, "-2" = -1.0)
connections$fcm_weight <- str_map[as.character(connections$Strength)]
connections <- connections[!is.na(connections$fcm_weight), ]

all_nodes <- elements$Label

if (!dir.exists(OUT_DIR)) dir.create(OUT_DIR, recursive = TRUE)

cat("Data loaded.\n")
cat(sprintf("Full model: %d concepts, %d connections\n",
            length(all_nodes), nrow(connections)))

# ================================================================
# SECTION 2 — HELPER: BUILD WEIGHT MATRIX FROM EDGE LIST
# ================================================================

build_wmat <- function(conn_df, nodes) {
  # Aggregate duplicate From-To pairs (same pair, multiple regions)
  # by taking the mean weight across occurrences.
  W <- matrix(0, length(nodes), length(nodes),
              dimnames = list(nodes, nodes))
  agg <- conn_df %>%
    filter(From %in% nodes, To %in% nodes) %>%
    group_by(From, To) %>%
    summarise(w = mean(fcm_weight), .groups = "drop")
  for (i in seq_len(nrow(agg))) W[agg$From[i], agg$To[i]] <- agg$w[i]
  W
}

# Filter connections to a specific sub-region (Tags column is pipe-separated)
filter_region <- function(conn_df, region) {
  conn_df[!is.na(conn_df$Tags) &
            grepl(region, conn_df$Tags, fixed = TRUE), ]
}

# ================================================================
# SECTION 3 — CORE ANALYSIS FUNCTION
# ================================================================
# This function receives a weight matrix and node list, computes
# every metric below, and returns a structured results object.
# It is called once per level (full model + each sub-region).

run_analysis <- function(W, nodes, label) {
  
  n <- length(nodes)
  
  # Build igraph object.
  # Edge weights = ABSOLUTE values of FCM weights so that
  # stronger connections always count as "closer" in path-based
  # metrics, regardless of whether they are positive or negative.
  edge_df <- data.frame(
    from   = rep(nodes, times = n),
    to     = rep(nodes, each  = n),
    weight = as.vector(t(abs(W)))  # row i → column j = W[i,j]
  ) %>% filter(weight > 0)
  
  if (nrow(edge_df) == 0) {
    message(sprintf("[%s] No edges found — skipping.", label))
    return(NULL)
  }
  
  g <- graph_from_data_frame(edge_df, directed = TRUE, vertices = nodes)
  E(g)$weight <- edge_df$weight
  
  # ── PART A: CONCEPT CLASSIFICATION ──────────────────────────
  #
  # Every node in an FCM falls into one of three structural roles:
  #
  # TRANSMITTER — has outgoing connections but NO incoming ones.
  #   → Pure input variable; influences the system but is not
  #     influenced by anything inside the model. Often represents
  #     external drivers (e.g., climate, policy context).
  #
  # RECEIVER — has incoming connections but NO outgoing ones.
  #   → Pure outcome variable; it is affected by others but exerts
  #     no further influence. Think: final consequences.
  #
  # ORDINARY — has BOTH incoming and outgoing connections.
  #   → Mediating variable; it is both driven and driving.
  #     Most nodes in a typical FCM are ordinary.
  #
  # ISOLATED — no connections at all (rare; usually data error).
  #
  # Note: classification uses WEIGHTED degree (sum of |weights|)
  # so a node with one weak and many strong links is treated
  # proportionally.
  
  indeg_w  <- colSums(abs(W))   # weighted indegree per node
  outdeg_w <- rowSums(abs(W))   # weighted outdegree per node
  
  type <- dplyr::case_when(
    indeg_w  == 0 & outdeg_w >  0 ~ "Transmitter",
    indeg_w  >  0 & outdeg_w == 0 ~ "Receiver",
    indeg_w  == 0 & outdeg_w == 0 ~ "Isolated",
    TRUE                          ~ "Ordinary"
  )
  
  # ── PART B: CENTRALITY MEASURES ─────────────────────────────
  
  # B1. INDEGREE CENTRALITY
  # Definition: sum of the absolute values of all INCOMING edge
  # weights for a node.
  # What it tells you: how strongly a concept is INFLUENCED by
  # others in the network. A high indegree means many (or strong)
  # causal forces push on that concept.
  # FCM relevance: identifies concepts most affected by the system.
  indegree <- indeg_w
  
  # B2. OUTDEGREE CENTRALITY
  # Definition: sum of the absolute values of all OUTGOING edge
  # weights for a node.
  # What it tells you: how strongly a concept INFLUENCES others.
  # A high outdegree means the node drives (or dampens) many
  # other concepts strongly.
  # FCM relevance: identifies the main causal drivers / leverage
  # points in the system.
  outdegree <- outdeg_w
  
  # B3. DEGREE CENTRALITY
  # Definition: indegree + outdegree (total connection weight).
  # What it tells you: overall involvement of a concept in the
  # network — both as a cause and as an effect.
  # FCM relevance: the single most commonly reported metric;
  # ranks concepts by their total embeddedness in the system.
  degree <- indegree + outdegree
  
  # B4. BETWEENNESS CENTRALITY
  # Definition: fraction of all shortest paths between pairs of
  # nodes that pass THROUGH a given node.
  # What it tells you: how critical a node is for TRANSMITTING
  # information across the network. Removing a high-betweenness
  # node would most disrupt the flow of influence.
  # Shortest paths are weighted by 1/|weight| so stronger links
  # are treated as shorter (more direct) paths.
  # FCM relevance: identifies bottleneck concepts — interventions
  # here would cascade broadly.
  btw <- tryCatch(
    setNames(betweenness(g, weights = 1 / E(g)$weight,
                         directed = TRUE, normalized = TRUE),
             V(g)$name),
    error = function(e) setNames(rep(0, n), nodes)
  )
  
  # B5. CLOSENESS CENTRALITY (harmonic)
  # Definition: for each node, the mean of 1/distance to all
  # other REACHABLE nodes (harmonic mean handles disconnected
  # graphs correctly by ignoring unreachable pairs).
  # What it tells you: how quickly a concept can reach all
  # others through the shortest paths — its "broadcasting speed."
  # A high closeness concept can spread its effects rapidly.
  # FCM relevance: complements betweenness by measuring not just
  # passage but also propagation speed.
  closeness <- sapply(nodes, function(v) {
    d <- distances(g, v = v, to = nodes,
                   weights = 1 / E(g)$weight, mode = "out")
    d <- d[d > 0 & is.finite(d)]
    if (length(d) == 0) return(0)
    mean(1 / d)
  })
  
  # B6. EIGENVECTOR CENTRALITY
  # Definition: a node's importance is proportional to the
  # importance of the nodes it is connected to. Solved as the
  # principal eigenvector of the (absolute) weight matrix.
  # What it tells you: being connected to other central nodes
  # amplifies your own centrality — quality over quantity.
  # FCM relevance: identifies concepts with system-wide influence
  # through their high-centrality neighbors. Does NOT work for
  # purely acyclic graphs (all scores → 0), but FCMs typically
  # contain feedback loops.
  eig <- tryCatch({
    ev <- eigen_centrality(g, directed = TRUE, weights = E(g)$weight)
    setNames(ev$vector, V(g)$name)
  }, error = function(e) setNames(rep(0, n), nodes))
  
  # B7. KATZ CENTRALITY
  # Definition: extends eigenvector centrality by giving every
  # node a small baseline importance (β), so nodes with zero
  # eigenvector score still propagate influence.
  # Formula: k = β · (I − α·Aᵀ)⁻¹ · 1
  # α (attenuation) must be < 1/spectral_radius to converge;
  # here α = 0.85/ρ (auto-set).
  # What it tells you: same as eigenvector but robust to nodes
  # with no feedback connections — particularly useful for
  # individual FCMs where feedback loops can be sparse.
  # FCM relevance: preferred over eigenvector when the graph
  # has acyclic sub-regions.
  katz <- tryCatch({
    ev_v <- eigen(t(abs(W)), only.values = TRUE)$values
    rho  <- max(Mod(ev_v))
    alp  <- 0.85 / max(rho, 1e-8)
    k    <- solve(diag(n) - alp * t(abs(W)), rep(1, n))
    k / max(abs(k))
  }, error = function(e) rep(0, n))
  
  # Assemble concept-level table
  concept_df <- data.frame(
    Concept     = nodes,
    Type        = type,
    Indegree    = round(indegree,  4),
    Outdegree   = round(outdegree, 4),
    Degree      = round(degree,    4),
    Betweenness = round(as.numeric(btw[nodes]),       4),
    Closeness   = round(closeness,                    4),
    Eigenvector = round(as.numeric(eig[nodes]),       4),
    Katz        = round(katz,                         4),
    stringsAsFactors = FALSE
  ) %>% arrange(desc(Degree))
  
  # ── PART C: VALIDATION METRICS ──────────────────────────────
  
  # C1. NUMBER OF CONCEPTS AND RELATIONSHIPS
  # The most basic descriptors of FCM size.
  # More concepts = wider scope; more relationships = greater
  # elaboration. These are compared across groups to assess
  # whether all participants contributed equally to the map.
  n_concepts <- n
  n_edges    <- ecount(g)
  
  # C2. RECEIVER-TRANSMITTER (R/T) RATIO
  # R/T = number of receivers / number of transmitters.
  # HIGH R/T: many outcomes from few inputs → complex map with
  #   rich causal chains but potentially few entry points.
  # LOW R/T:  many inputs, few outcomes → over-simplified;
  #   causal relationships not fully elaborated.
  # Moderate R/T (1–3) is generally desirable.
  n_recv <- sum(type == "Receiver")
  n_tran <- sum(type == "Transmitter")
  rt_ratio <- if (n_tran > 0) round(n_recv / n_tran, 3) else NA
  
  # C3. SHORTEST PATH METRICS
  # Distance d(i,j) = length of shortest path from i to j,
  # weighted by 1/|weight| (stronger link = shorter distance).
  #
  # Average path length: mean of all finite pairwise distances.
  #   Small avg path → effects spread quickly through the map.
  #
  # Diameter: the LARGEST shortest path in the network.
  #   It tells you the worst-case number of "steps" for an
  #   influence to travel from one end of the map to the other.
  #   A very small diameter relative to n can indicate that
  #   causal chains are not fully elaborated.
  g_dist <- g
  E(g_dist)$weight <- 1 / E(g)$weight
  avg_path <- round(mean_distance(g_dist, directed = TRUE,
                                  unconnected = TRUE), 4)
  diam     <- round(diameter(g_dist, directed = TRUE,
                             unconnected = TRUE), 4)
  
  # C4. CLUSTERING COEFFICIENT
  # For each node: fraction of its neighbors that are also
  # connected to each other (local transitivity).
  # Average clustering coefficient = mean across all nodes.
  # Global clustering coefficient = fraction of all connected
  # triples that form closed triangles.
  #
  # HIGH clustering → participants added many interconnected
  #   relationships (dense local neighborhoods).
  # LOW clustering  → relationships are more chain-like;
  #   participants were selective.
  # FCM context: clustering of ~0.2–0.4 is typical and desirable.
  clust_local  <- transitivity(g, type = "local", isolates = "zero")
  clust_avg    <- round(mean(clust_local, na.rm = TRUE), 4)
  clust_global <- round(transitivity(g, type = "global"),  4)
  
  # C5. DENSITY
  # Density = actual edges / maximum possible edges.
  # For a directed graph: density = |E| / (n × (n−1))
  # LOW density is desirable in FCMs because it means
  # participants were selective — they did not connect every
  # concept to every other one. Density typically decreases
  # as n grows.
  density_val <- round(ecount(g) / (n * (n - 1)), 4)
  
  # Assemble validation table
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
cat("FULL MODEL — AUSTRALIA (all regions combined)\n")
cat("================================================================\n")

W_full      <- build_wmat(connections, all_nodes)
res_full    <- run_analysis(W_full, all_nodes, "Australia (full)")

cat("\n--- PART A: Concept Classification ---\n")
cat("Definition summary:\n")
cat("  Transmitter = only outgoing links (external drivers)\n")
cat("  Receiver    = only incoming links (outcomes/consequences)\n")
cat("  Ordinary    = both incoming and outgoing links (mediators)\n\n")
print(table(res_full$concept_df$Type))

cat("\n--- Transmitters ---\n")
print(res_full$concept_df[res_full$concept_df$Type == "Transmitter",
                          c("Concept","Outdegree")], row.names = FALSE)

cat("\n--- Receivers ---\n")
print(res_full$concept_df[res_full$concept_df$Type == "Receiver",
                          c("Concept","Indegree")], row.names = FALSE)

cat("\n--- PART B: Centrality Measures — Top 20 by Degree ---\n")
print(head(res_full$concept_df, 20), row.names = FALSE)

cat("\n--- PART C: Validation Metrics ---\n")
print(res_full$validation_df, row.names = FALSE)

if (SAVE_CSV) {
  write.csv(res_full$concept_df,
            file.path(OUT_DIR, "AU_full_centrality.csv"), row.names = FALSE)
  write.csv(res_full$validation_df,
            file.path(OUT_DIR, "AU_full_validation.csv"), row.names = FALSE)
}

# ================================================================
# SECTION 5 — RUN ANALYSIS: EACH SUB-REGION
# ================================================================

cat("\n")
cat("================================================================\n")
cat("SUB-REGION ANALYSES\n")
cat("================================================================\n")

sub_results <- list()

for (reg in SUBREGIONS) {
  
  cat(sprintf("\n--- Region: %s ---\n", reg))
  
  conn_r  <- filter_region(connections, reg)
  nodes_r <- unique(c(conn_r$From, conn_r$To))
  nodes_r <- nodes_r[nodes_r %in% all_nodes]
  
  if (length(nodes_r) < 2) {
    cat(sprintf("  Fewer than 2 nodes in %s — skipping.\n", reg))
    next
  }
  
  W_r   <- build_wmat(conn_r, nodes_r)
  res_r <- run_analysis(W_r, nodes_r, reg)
  
  if (is.null(res_r)) next
  sub_results[[reg]] <- res_r
  
  cat(sprintf("  Concepts: %d  |  Connections: %d\n",
              res_r$n_nodes, res_r$n_edges))
  cat("  Concept types: ")
  print(table(res_r$concept_df$Type))
  
  cat("  Top 10 by Degree Centrality:\n")
  print(head(res_r$concept_df[, c("Concept","Type","Degree",
                                  "Betweenness","Katz")], 10),
        row.names = FALSE)
  
  cat("  Validation:\n")
  print(res_r$validation_df, row.names = FALSE)
  
  if (SAVE_CSV) {
    safe_name <- gsub(" ", "_", reg)
    write.csv(res_r$concept_df,
              file.path(OUT_DIR, sprintf("AU_%s_centrality.csv", safe_name)),
              row.names = FALSE)
    write.csv(res_r$validation_df,
              file.path(OUT_DIR, sprintf("AU_%s_validation.csv", safe_name)),
              row.names = FALSE)
  }
}

# ================================================================
# SECTION 6 — COMPARISON TABLE ACROSS ALL LEVELS
# ================================================================

cat("\n")
cat("================================================================\n")
cat("COMPARISON: FULL MODEL vs SUB-REGIONS\n")
cat("================================================================\n")

# Collect validation metrics for every level in one table
all_results <- c(list("Australia (full)" = res_full), sub_results)

compare_df <- do.call(rbind, lapply(names(all_results), function(nm) {
  vdf <- all_results[[nm]]$validation_df
  row <- setNames(as.list(vdf$Value), vdf$Metric)
  data.frame(Level = nm, as.data.frame(row, stringsAsFactors = FALSE),
             check.names = FALSE, stringsAsFactors = FALSE)
}))

# Select key metrics for the comparison print
key_cols <- c("Level", "Concepts (nodes)", "Connections (edges, aggregated)",
              "Transmitters", "Receivers", "R/T Ratio",
              "Average Path Length", "Diameter",
              "Average Clustering Coefficient", "Density")
key_cols <- key_cols[key_cols %in% names(compare_df)]

cat("\nKey metrics across all levels:\n")
print(compare_df[, key_cols], row.names = FALSE)

# Top-5 by degree per level — useful for quick cross-comparison
cat("\n--- Top 5 Concepts by Degree Centrality per Level ---\n")
for (nm in names(all_results)) {
  top5 <- head(all_results[[nm]]$concept_df$Concept, 5)
  cat(sprintf("  %-30s: %s\n", nm, paste(top5, collapse = ", ")))
}

if (SAVE_CSV) write.csv(compare_df,
                        file.path(OUT_DIR, "AU_comparison_table.csv"),
                        row.names = FALSE)

# ================================================================
# SECTION 7 — VISUALIZATIONS
# ================================================================

cat("\n--- Generating plots ---\n")

model_colours <- c(
  "Australia (full)"  = "#1B9E77",
  "New South Wales"   = "#D95F02",
  "Queensland"        = "#7570B3",
  "North Australia"   = "#E7298A",
  "Western Australia" = "#66A61E"
)

# ── Plot 1: Degree centrality — full model, top 25 ────────────
p1 <- res_full$concept_df %>%
  head(25) %>%
  ggplot(aes(x = reorder(Concept, Degree), y = Degree, fill = Type)) +
  geom_col() +
  coord_flip() +
  scale_fill_brewer(palette = "Set2") +
  labs(
    title    = "Degree Centrality — Australia (Full Model)",
    subtitle = "Top 25 concepts; fill = concept type",
    x = NULL, y = "Degree (sum of |incoming| + |outgoing| weights)"
  ) +
  theme_bw(base_size = 9)

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "01_AU_full_degree.pdf"),
                       p1, width = 10, height = 8)
print(p1)

# ── Plot 2: Centrality comparison across sub-regions ─────────
# For each sub-region, top-3 concepts by degree
top3_df <- do.call(rbind, lapply(names(all_results), function(nm) {
  head(all_results[[nm]]$concept_df[, c("Concept","Degree")], 3) %>%
    mutate(Level = nm, Rank = row_number())
}))

p2 <- ggplot(top3_df, aes(x = Rank, y = Degree,
                          fill = Level, label = Concept)) +
  geom_col(position = "dodge") +
  geom_text(position = position_dodge(0.9), hjust = -0.1,
            size = 2.8, angle = 0) +
  facet_wrap(~Level, scales = "free_y") +
  coord_flip() +
  scale_fill_manual(values = model_colours) +
  labs(title = "Top 3 Concepts by Degree — Australia and Sub-Regions",
       x = "Rank", y = "Degree Centrality") +
  theme_bw(base_size = 8) +
  theme(legend.position = "none",
        strip.text = element_text(size = 7, face = "bold"))

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "02_AU_subregion_top3.pdf"),
                       p2, width = 12, height = 8)
print(p2)

# ── Plot 3: Validation metrics bar chart ──────────────────────
val_long <- compare_df %>%
  select(Level,
         Density,
         `Average Clustering Coefficient`,
         `R/T Ratio`) %>%
  pivot_longer(-Level, names_to = "Metric", values_to = "Value") %>%
  mutate(Value = as.numeric(Value))

p3 <- ggplot(val_long, aes(x = Level, y = Value, fill = Level)) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~Metric, scales = "free_y") +
  scale_fill_manual(values = model_colours) +
  labs(title = "Validation Metrics — Australia and Sub-Regions",
       x = NULL, y = "Value") +
  theme_bw(base_size = 9) +
  theme(axis.text.x = element_text(angle = 35, hjust = 1, size = 7),
        strip.text  = element_text(face = "bold"))

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "03_AU_validation_comparison.pdf"),
                       p3, width = 10, height = 5)
print(p3)

# ── Plot 4: Betweenness — full model, top 20 ──────────────────
p4 <- res_full$concept_df %>%
  arrange(desc(Betweenness)) %>%
  head(20) %>%
  ggplot(aes(x = reorder(Concept, Betweenness),
             y = Betweenness, fill = Type)) +
  geom_col() +
  coord_flip() +
  scale_fill_brewer(palette = "Set2") +
  labs(title    = "Betweenness Centrality — Australia (Full Model)",
       subtitle  = "Top 20 concepts; high betweenness = key bridge/bottleneck",
       x = NULL, y = "Betweenness (normalised)") +
  theme_bw(base_size = 9)

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "04_AU_full_betweenness.pdf"),
                       p4, width = 10, height = 7)
print(p4)

# ── Plot 5: Concept type breakdown per level ──────────────────
type_df <- do.call(rbind, lapply(names(all_results), function(nm) {
  tab <- table(all_results[[nm]]$concept_df$Type)
  data.frame(Level = nm, Type = names(tab), Count = as.integer(tab))
}))

p5 <- ggplot(type_df, aes(x = Level, y = Count, fill = Type)) +
  geom_col(position = "fill") +
  scale_y_continuous(labels = percent) +
  scale_fill_brewer(palette = "Set2") +
  labs(title = "Concept Type Composition per Level — Australia",
       x = NULL, y = "Proportion") +
  theme_bw(base_size = 9) +
  theme(axis.text.x = element_text(angle = 30, hjust = 1))

if (SAVE_PLOTS) ggsave(file.path(OUT_DIR, "05_AU_type_composition.pdf"),
                       p5, width = 9, height = 5)
print(p5)

cat(sprintf("\nAll outputs saved to: %s\n", OUT_DIR))
cat("Analysis complete.\n")
