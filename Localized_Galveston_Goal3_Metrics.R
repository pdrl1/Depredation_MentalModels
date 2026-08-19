# ============================================================
#  MENTAL MODEL ANALYSIS — GALVESTON
#  Galveston community model, by fishery sector
#
#  This script analyses Galveston at its own scale:
#    1) The AGGREGATED Galveston community model (all workshops
#       pooled together, tag-independent) — including prominence.
#    2) Each of the Galveston FISHERY SECTORS separately (from the
#       Tags column): Recreational · Charter · Commercial.
#
#  For every unit (aggregate + each fishery sector) this script
#  computes:
#    — All structural network metrics (degree, concept types,
#      centrality, path metrics, clustering, feedback loops)
#    — Prominence (Hoffman et al., 2014)
#
#  FIGURES:
#    — Indegree vs Outdegree  (one figure, faceted by unit,
#      aggregate model included)
#    — Closeness vs Betweenness (one figure, faceted by unit,
#      aggregate model included)
#    — Prominence ranking heatmap (one figure, one column per
#      unit, aggregate model included as its own column)
#
#  This script supersedes the Galveston portion of the previous
#  Goal1_GlobalComparison.R (aggregate model, metrics-only) and
#  Goal2_RegionalComparison.R (fishery-level breakdown, metrics-only)
#  scripts, which did not previously produce Galveston figures.
#
#  Strength scale: integer +-1/+-2 (same as Australia/Gulf Coast) --
#  no rescaling applied. The script checks this assumption below.
# ============================================================


# ============================================================
# SECTION 0 — PACKAGES
# ============================================================

required_pkgs <- c("readxl", "igraph", "dplyr", "tidyr", "ggplot2",
                   "stringr", "scales", "ggrepel", "purrr",
                   "tidytext", "ggnewscale")
for (pkg in required_pkgs) {
  if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
}
library(readxl); library(igraph); library(dplyr); library(tidyr)
library(ggplot2); library(stringr); library(scales)
library(ggrepel); library(purrr); library(tidytext); library(ggnewscale)


# ============================================================
# SECTION 1 — DATA LOADING
# ============================================================
# Kumu master export for Galveston.
# Tags column: pipe-separated fishery-sector labels,
#   e.g. "Recreational|Charter"
# >>> SET THE FILE PATH BELOW to the Galveston Kumu export before running. <<<

GV_FILE <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Exported_Kumu_Galveston_30April.xlsx"  

gv_elements    <- read_excel(GV_FILE, sheet = "Elements")
gv_connections <- read_excel(GV_FILE, sheet = "Connections")

# Parse numeric strength once, drop rows with no strength
gv_connections <- gv_connections %>%
  mutate(strength_num = suppressWarnings(as.numeric(Strength))) %>%
  filter(!is.na(From), !is.na(To), !is.na(strength_num))

cat("Loaded -- Galveston:", nrow(gv_elements), "elements |",
    nrow(gv_connections), "connections\n")

# Strength-scale check: expect integer +-1/+-2 (not the +-0.5/+-1.0 MM scale)
gv_strength_vals <- sort(unique(gv_connections$strength_num))
cat("Galveston unique Strength values:",
    paste(gv_strength_vals, collapse = ", "), "\n")
if (any(abs(gv_strength_vals) > 0 & abs(gv_strength_vals) < 1))
  warning("Galveston Strength contains |value| < 1 (looks like the +-0.5/+-1.0 ",
          "Mental Modeler scale). Multiply by 2 to match the integer +-1/+-2 ",
          "scale used for Australia and Gulf Coast before trusting weighted metrics.")

# Element group lookup table (for colouring plots).
# Uses 'Category' if present (matches the Gulf Coast convention), else 'Type'
# (matches the Australia convention). Adjust the column name below if your
# Galveston export uses a different group column.
gv_group_col <- if ("Category" %in% names(gv_elements)) "Category" else "Type"
gv_groups <- gv_elements %>%
  select(name = Label, Group = all_of(gv_group_col)) %>%
  distinct() %>%
  filter(!is.na(Group))

cat("GV groups (from '", gv_group_col, "' column):",
    paste(sort(unique(gv_groups$Group)), collapse = ", "), "\n", sep = "")


# ============================================================
# SECTION 2 — UNIT DEFINITIONS
# ============================================================
# The "aggregate" unit = the full pooled Galveston model (no Tags
# filtering). The fishery-sector units mirror the Tags column exactly.

GV_SECTORS <- c("Recreational", "Charter", "Commercial")

AGGREGATE_LABEL <- "Galveston Community Model"
GV_UNITS <- c(AGGREGATE_LABEL, GV_SECTORS)


# ============================================================
# SECTION 3 — GRAPH BUILDING
# ============================================================

# 3.1 Filter any data frame by a Tags column entry.
#     tag_name = NULL means "no filtering" -> the aggregate model.
filter_by_tag <- function(df, tag_name) {
  if (is.null(tag_name)) return(df)
  df %>% filter(!is.na(Tags), str_detect(Tags, fixed(tag_name)))
}

# 3.2 Build a deduplicated edge list for one unit (aggregate or sector).
#     The Kumu master file may contain the same (From, To) pair more
#     than once (different individual workshops gave it different
#     strengths). We average those strengths to obtain one edge
#     weight per (From, To) pair -- same treatment used throughout
#     the previous Goal1/Goal2 scripts.
build_unit_edges <- function(all_connections, tag_name = NULL) {
  filter_by_tag(all_connections, tag_name) %>%
    group_by(From, To) %>%
    summarise(
      avg_strength = mean(strength_num, na.rm = TRUE),
      polarity     = sign(mean(strength_num, na.rm = TRUE)),
      n_rows       = n(),
      .groups      = "drop"
    ) %>%
    filter(!is.na(avg_strength))
}

# 3.3 Build an igraph directed weighted graph for one unit.
#     Nodes = elements belonging to this unit (Elements sheet, or all
#     elements for the aggregate) plus any additional nodes that
#     appear only in edges.
build_unit_graph <- function(all_elements, all_connections, tag_name = NULL) {
  edges      <- build_unit_edges(all_connections, tag_name)
  tagged_nds <- if (is.null(tag_name)) {
    unique(all_elements$Label)
  } else {
    filter_by_tag(all_elements, tag_name)$Label
  }
  edge_nds <- unique(c(edges$From, edges$To))
  all_nds  <- unique(c(tagged_nds, edge_nds))
  
  if (nrow(edges) == 0) {
    warning(sprintf("No edges for unit: %s",
                    if (is.null(tag_name)) "AGGREGATE" else tag_name))
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

cat("\n=== Building Galveston unit graphs (aggregate + fishery sectors) ===\n")
gv_graphs <- setNames(
  c(list(build_unit_graph(gv_elements, gv_connections, tag_name = NULL)),
    lapply(GV_SECTORS, build_unit_graph,
           all_elements = gv_elements, all_connections = gv_connections)),
  GV_UNITS)

for (u in GV_UNITS) {
  g <- gv_graphs[[u]]
  if (!is.null(g))
    cat(sprintf("  %-24s N=%d  E=%d\n", u, vcount(g), ecount(g)))
}


# ============================================================
# SECTION 4 — HELPER FUNCTIONS (metric computation)
# ============================================================
# Self-contained so this script can run independently.

# 4.1 Weighted indegree / outdegree / degree
#     Indegree_i  = Sum_j |w_ji|   (influence RECEIVED by i)
#     Outdegree_i = Sum_j |w_ij|   (influence EXERTED by i)
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
#     Eigenvector centrality computed on the UNDIRECTED graph, because
#     directed eigenvector centrality is 0 for acyclic (DAG) sector
#     networks (a real risk once we split into small sectors). Using
#     the undirected version everywhere -- aggregate AND sectors --
#     keeps prominence comparable across every unit in this script.
compute_prominence <- function(g, node_df) {
  er <- eigen_centrality(g, directed = FALSE,
                         weights = abs(E(g)$weight), scale = TRUE)
  
  ec_df <- tibble(name = names(er$vector),
                  eigen_centrality = as.numeric(er$vector))
  
  mn <- min(ec_df$eigen_centrality, na.rm = TRUE)
  mx <- max(ec_df$eigen_centrality, na.rm = TRUE)
  
  node_df %>%
    select(-any_of(c("eigen_centrality", "norm_eigen", "occurrence", "prominence"))) %>%
    left_join(ec_df, by = "name") %>%
    mutate(
      norm_eigen = if (mx == mn) 0 else (eigen_centrality - mn) / (mx - mn),
      occurrence = 1,
      prominence = (occurrence + norm_eigen) / 2
    )
}


# ============================================================
# SECTION 5 — COMPUTE ALL METRICS PER UNIT
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
  
  # Enumerate and classify simple cycles (<= 7 steps)
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
    cat(sprintf("\n-- %s --\n", model_name))
    cat(sprintf("  N=%d  E=%d  D=%.4f  R/T=%s\n",
                N, E_count, density,
                if (is.na(RT)) "undef" else sprintf("%.3f", RT)))
    cat(sprintf("  TX=%d  RX=%d  ORD=%d  ISO=%d\n",
                n_tx, n_rx, n_ord, n_iso))
    cat(sprintf("  APL=%s  Diam=%s  Clust_avg=%.3f  CR=%d\n",
                if (is.na(apl))  "-" else sprintf("%.3f", apl),
                if (is.na(diam)) "-" else as.character(diam),
                clust_avg, circ_rank))
    cat(sprintf("  Cycles<=7=%d  R-loops=%d  B-loops=%d\n",
                length(cycles), n_reinf, n_bal))
    cat(sprintf("  Pos=%d (%.0f%%)  Neg=%d (%.0f%%)\n",
                n_pos, 100*n_pos/max(E_count,1),
                n_neg, 100*n_neg/max(E_count,1)))
    
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

cat("\n\n========== GALVESTON -- METRICS (AGGREGATE + FISHERY SECTORS) ==========\n")
gv_metrics <- setNames(
  lapply(GV_UNITS, function(u) compute_all_metrics(gv_graphs[[u]], u)),
  GV_UNITS)


# ============================================================
# SECTION 6 — SUMMARY METRICS TABLE
# ============================================================

make_metrics_table <- function(metrics_list, label) {
  cat(sprintf("\n\n========== METRICS TABLE -- %s ==========\n", label))
  tbl <- bind_rows(lapply(names(metrics_list), function(u) {
    m <- metrics_list[[u]]
    if (is.null(m)) return(NULL)
    data.frame(
      Unit         = u,
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

gv_tbl <- make_metrics_table(gv_metrics, "GALVESTON (AGGREGATE + FISHERY SECTORS)")


# ============================================================
# SECTION 7 — VISUALISATIONS
# ============================================================

# Fixed display order for groups (top of plot -> bottom on reversed axis)
GROUP_ORDER <- c(
  "Central Concept",
  "Ecological & Biological Factors",
  "Human Dimensions",
  "Fisheries Operations & Practices",
  "Fisheries Research & Management",
  "Policy & Economics",
  "Other"
)

GROUP_COLOURS <- c(
  "Central Concept"                   = "#FBD73F",
  "Ecological & Biological Factors"   = "#9AD354",
  "Human Dimensions"                  = "#F8895E",
  "Fisheries Operations & Practices"  = "#E382BA",
  "Fisheries Research & Management"   = "#8695C2",
  "Policy & Economics"                = "#5EB99B",
  "Other"                             = "grey60"
)

TYPE_COLOURS <- c("Transmitter" = "#D55E00",
                  "Receiver"    = "#56B4E9",
                  "Ordinary"    = "#009E73",
                  "Isolated"    = "grey80")

# Shared plot theme
theme_fcm <- function(base_size = 14) {
  theme_bw(base_size = base_size) %+replace%
    theme(
      strip.text = element_text(face = "bold", size = base_size + 1,
                                margin = margin(b = 6)),
      strip.background = element_rect(colour = NA),
      axis.text.y      = element_text(size = base_size - 2),
      axis.text.x      = element_text(size = base_size - 2),
      panel.grid.minor = element_blank(),
      legend.position  = "bottom",
      legend.title     = element_text(size = base_size - 1),
      plot.title       = element_text(face = "bold", size = base_size + 2, margin = margin(b = 4)),
      plot.subtitle    = element_text(size = base_size - 1, margin = margin(b = 8))
    )
}

# ---- Build a combined node table across ALL units (aggregate + sectors) ----
full_df_gv <- bind_rows(lapply(names(gv_metrics), function(u) {
  m <- gv_metrics[[u]]; if (is.null(m)) return(NULL)
  m$node_df %>% mutate(Unit = u)
})) %>%
  mutate(Unit = factor(Unit, levels = GV_UNITS)) %>%
  rename(Concept = name, Type = concept_type, Degree = degree,
         Indegree = indegree, Outdegree = outdegree,
         Closeness = closeness, Betweenness = betweenness)


# ---- 7.1 INDEGREE vs OUTDEGREE (aggregated model only) ----

cat("Plot: Indegree vs Outdegree (Galveston, all units)...\n")

agg_df_gv <- full_df_gv %>%
  filter(Unit == AGGREGATE_LABEL) %>%
  mutate(Type = factor(Type, levels = names(TYPE_COLOURS)))

p_keep <- 0.80   # label the top 20% (>= 80th percentile) of Degree

IO_XLIM<-c(0,22)
IO_YLIM<-c(0,20)

label_df_gv <- agg_df_gv %>%
  filter(Degree >= quantile(Degree, p_keep, na.rm = TRUE))

# Place the "1:1 line" label near the far end of the diagonal (within the
# visible axis range), where the point cloud is sparsest, rather than at a
# data-derived position -- this keeps it from landing on top of/behind the
# concept labels and geom_label_repel boxes.
io_diag_max   <- min(IO_XLIM[2], IO_YLIM[2])
io_label_pos  <- 0.25 * io_diag_max

p_gv_io <- ggplot(agg_df_gv, aes(x = Indegree, y = Outdegree,
                                 colour = Type, size = Degree)) +
  geom_abline(slope = 1, intercept = 0,
              linetype = "dashed", colour = "grey60", linewidth = 0.5) +
  geom_point(alpha = 0.6, stroke = 0.2) +
  geom_label_repel(
    data          = label_df_gv,
    aes(label     = str_wrap(Concept, 20)),
    size          = 3.2,
    label.padding = unit(0.12, "lines"),
    box.padding   = unit(0.4, "lines"),
    point.padding = unit(0.2, "lines"),
    min.segment.length = 0,
    segment.colour = "grey55", segment.size = 0.3,
    max.overlaps  = Inf, seed = 42, show.legend = FALSE,
    colour        = "grey15", fill = alpha("white", 0.85)
  ) +
  # Drawn LAST (on top of every other layer, incl. geom_label_repel boxes)
  # so it can never be buried. Un-rotated (angle=0) + solid black + a
  # white background box -- the least ambiguous, most version-proof way
  # to render a label in ggplot2 (rotated geom_label boxes are unreliable
  # across ggplot2 versions, which is the likely reason it was invisible
  # before even after being repositioned).
  annotate("text", x = io_label_pos, y = io_label_pos, label = "1:1 line",
           angle = 35, vjust = -1, size = 3, colour = "grey45") +
  scale_colour_manual(values = TYPE_COLOURS, drop = FALSE) +
  guides(colour = guide_legend(override.aes = list(size = 4, alpha = 1))) +
  scale_size_continuous(range = c(1.5, 6), guide = "none") +
  coord_cartesian(xlim = c(0,7), ylim = c(0,7)) +
  labs(title = NULL,
       x = "Indegree",
       y = "Outdegree",
       colour = "Concept type") +
  theme_fcm() +
  theme(
    plot.title    = element_text(size = 18, face = "bold"),
    axis.title    = element_text(size = 13),
    strip.text    = element_text(size = 12, face = "bold"),
    legend.title  = element_text(size = 12),
    legend.text   = element_text(size = 12)
  )


ggsave("Galveston_indegree_vs_outdegree.png", p_gv_io,
       width = 18, height = 6, units = "in", dpi = 1200)
ggsave("Galveston_indegree_vs_outdegree.pdf", p_gv_io,
       width = 20, height = 7, units = "in", device = cairo_pdf)
ggsave("Galveston_indegree_vs_outdegree.tiff", p_gv_io,
       width = 20, height = 7, units = "in", dpi = 600, compression = "lzw")


# ---- 7.2 CLOSENESS vs BETWEENNESS (only aggregated model) ----

cat("Plot: Closeness vs Betweenness (Galveston, all units)...\n")

p_keep_n <- 10   # label the top-10 by closeness OR betweenness, per unit


plot_df_cb <- agg_df_gv %>%
  filter(!is.na(Closeness), !is.na(Betweenness)) %>%
  group_by(Unit) %>%
  mutate(
    keep = rank(-Closeness,   ties.method = "min") <= p_keep_n |
      rank(-Betweenness, ties.method = "min") <= p_keep_n,
    lab  = ifelse(keep, str_wrap(Concept, 16), "")
  ) %>%
  ungroup()

quad_lines_gv <- plot_df_cb %>%
  group_by(Unit) %>%
  summarise(med_closeness   = median(Closeness,   na.rm = TRUE),
            med_betweenness = median(Betweenness, na.rm = TRUE),
            .groups = "drop")

p_gv_cb <- ggplot(plot_df_cb, aes(x = Closeness, y = Betweenness)) +
  geom_vline(data = quad_lines_gv, aes(xintercept = med_closeness),
             linetype = "dashed", colour = "grey70", linewidth = 0.4) +
  geom_hline(data = quad_lines_gv, aes(yintercept = med_betweenness),
             linetype = "dashed", colour = "grey70", linewidth = 0.4) +
  geom_point(aes(colour = Type, size = Degree), alpha = 0.6, stroke = 0.2) +
  geom_label_repel(
    data = plot_df_cb, aes(label = lab),
    size = 3.2, label.padding = unit(0.12, "lines"),
    box.padding = unit(0.5, "lines"), point.padding = unit(0.4, "lines"),
    min.segment.length = 0, segment.colour = "grey55", segment.size = 0.3,
    max.overlaps = Inf, seed = 42, show.legend = FALSE,
    colour = "grey15", fill = alpha("white", 0.85)
  ) +
  scale_colour_manual(values = TYPE_COLOURS, drop = FALSE) +
  guides(colour = guide_legend(override.aes = list(size = 4, alpha = 1))) +
  scale_size_continuous(range = c(1.5, 6), guide = "none") +
  coord_cartesian(xlim = c(0.25,1.25), ylim = c(0,0.05)) +
  scale_x_sqrt() +
  scale_y_sqrt() +
  labs(title = NULL,
       x = "Closeness",
       y = "Betweenness",
       colour = "Concept type") +
  theme_fcm() +
  theme(
    plot.title    = element_text(size = 18, face = "bold"),
    axis.title    = element_text(size = 13),
    strip.text    = element_text(size = 12, face = "bold"),
    legend.title  = element_text(size = 12),
    legend.text   = element_text(size = 12)
  )

ggsave("Galveston_closeness_vs_betweenness.png", p_gv_cb,
       width = 18, height = 6, units = "in", dpi = 1200)
ggsave("Galveston_closeness_vs_betweenness.pdf", p_gv_cb,
       width = 20, height = 7, units = "in", device = cairo_pdf)
ggsave("Galveston_closeness_vs_betweenness.tiff", p_gv_cb,
       width = 20, height = 7, units = "in", dpi = 600, compression = "lzw")


# ---- 7.1b COMBINED PANEL: Indegree/Outdegree (left) + Closeness/Betweenness
#      (right), AGGREGATE MODEL ONLY ----


cat("Plot: Combined Indegree/Outdegree + Closeness/Betweenness panel Galveston aggregated model only)...\n")

p_gv_io_combo <- p_gv_io + labs(title = NULL, subtitle = NULL)
p_gv_cb_combo <- p_gv_cb + labs(title = NULL) +
  guides(colour = "none")   # <-- panel B contributes no legend of its own

p_gv_combo <- (p_gv_io_combo | p_gv_cb_combo) +
  patchwork::plot_layout(guides = "collect") +
  patchwork::plot_annotation(
    title = NULL,
    tag_levels = "A"
  ) &
  theme(
    legend.position = "bottom",
    plot.title      = element_text(size = 20, face = "bold", hjust = 0.5)
  )


ggsave("Galveston_indegree_closeness_combined.png", p_gv_combo,
       width = 16, height = 8, units = "in", dpi = 1200)
ggsave("Galveston_GC_indegree_closeness_combined.pdf", p_gv_combo,
       width = 17, height = 9, units = "in", device = cairo_pdf)
ggsave("Galveston_GC_indegree_closeness_combined.tiff", p_gv_combo,
       width = 17, height = 9, units = "in", dpi = 600, compression = "lzw")



# ---- 7.3 PROMINENCE RANKING (one column per unit, aggregate included) ----

cat("Plot: Prominence ranking (Galveston, all units)...\n")

plot_prominence_rank <- function(metrics_list, label, groups_df = NULL,
                                 top_n = 10, col_order = names(metrics_list)) {
  
  group_var <- if (!is.null(groups_df)) "Group" else "concept_type"
  
  plot_data <- bind_rows(lapply(names(metrics_list), function(r) {
    m <- metrics_list[[r]]; if (is.null(m)) return(NULL)
    df <- m$node_df %>% arrange(desc(prominence)) %>% head(top_n) %>%
      mutate(subregion = r, prom_rank = row_number())
    if (!is.null(groups_df))
      df <- df %>% left_join(groups_df, by = "name") %>% mutate(Group = replace_na(Group, "Other"))
    else
      df <- df %>% mutate(Group = concept_type)
    df
  }))
  
  plot_data <- plot_data %>% mutate(subregion = factor(subregion, levels = col_order))
  
  group_colours_used <- if (!is.null(groups_df)) GROUP_COLOURS else TYPE_COLOURS
  
  group_levels_present <- GROUP_ORDER[GROUP_ORDER %in% unique(plot_data$Group)]
  
  concept_groups <- plot_data %>%
    group_by(name) %>%
    summarise(Group = first(.data[[group_var]]),
              mean_rank = mean(prom_rank, na.rm = TRUE), .groups = "drop") %>%
    mutate(Group = factor(Group, levels = group_levels_present)) %>%
    arrange(Group, mean_rank)
  
  concept_order <- concept_groups$name
  plot_data <- plot_data %>% mutate(name = factor(name, levels = rev(concept_order)))
  
  label_colours <- group_colours_used[
    concept_groups$Group[match(rev(concept_order), concept_groups$name)]]
  label_colours[is.na(label_colours)] <- "grey20"
  
  group_lookup <- setNames(as.character(concept_groups$Group), concept_groups$name)
  plot_data <- plot_data %>% mutate(name = factor(name, levels = rev(concept_order)))
  
  ggplot(plot_data, aes(x = subregion, y = name)) +
    geom_tile(aes(fill = prom_rank), colour = "white", linewidth = 0.8) +
    geom_text(aes(label = prom_rank), size = 4, colour = "black") +
    scale_fill_gradientn(
      colours  = c("#2171b5", "#6baed6", "#c6dbef", "#f7fbff"),
      na.value = "grey92",
      name     = "Prominence\nranking",
      limits   = c(1, top_n),
      breaks   = c(1, top_n),
      labels   = c("1 (highest)", paste0(top_n, " (lowest)")),
      guide    = guide_colourbar(position = "right", reverse = TRUE)) +
    
    ggnewscale::new_scale_fill() +
    geom_point(aes(fill = Group), shape = 22, size = 0, alpha = 0, stroke = 0) +
    scale_fill_manual(values = group_colours_used, name = NULL, drop = FALSE,
                      guide = guide_legend(position = "bottom", override.aes = list(size = 5, alpha = 1))) +
    
    scale_x_discrete(expand = expansion(add = c(0.6, 0.3))) +
    labs(title = label, x = NULL, y = NULL) +
    theme_fcm() +
    theme(
      plot.title    = element_text(size = 14, face = "bold"),
      axis.title    = element_text(size = 11),
      axis.text     = element_text(size = 12),
      axis.text.y   = element_text(size = 12, face = "bold", colour = label_colours),
      axis.text.x   = element_text(size = 12, face = "bold"),
      strip.text    = element_text(size = 11, face = "bold"),
      legend.title  = element_text(size = 12, hjust = 0.5, margin = margin(b = 14)),
      legend.text   = element_text(size = 12),
      panel.grid.major.x = element_blank(),
      panel.grid.minor   = element_blank(),
      plot.caption  = element_text(size = 8, hjust = 0, colour = "grey30")
    )
}

prom_Gv_rank <- plot_prominence_rank(
  gv_metrics, "", groups_df = gv_groups,
  col_order = GV_UNITS)

ggsave("Galveston_prominence_rank.png", prom_Gv_rank,
       width = 14, height = 11, dpi = 1200)
ggsave("Galveston_prominence_rank.pdf", prom_Gv_rank,
       width = 16, height = 12, device = cairo_pdf)

cat("\n\nGalveston analysis complete. All outputs saved.\n")