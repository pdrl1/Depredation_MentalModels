# ============================================================
#  MENTAL MODEL ANALYSIS — COMBINED FIGURES
#  Australia (national) + US Gulf Coast (regional) + Galveston (local)
#
#  Rebuilds the three community models from their Kumu exports (same
#  data loading / graph building / metric computation as
#  National_Australia_Goal1.R, Regional_USGC_Goal2.R and
#  Localized_Galveston_Goal3_Metrics.R -- reproduced once here, since
#  the logic in those three scripts is identical) and assembles:
#
#  FIGURE 1 -- 3 columns x 2 rows, AGGREGATE (full) models only
#    Col 1 = Australia | Col 2 = US Gulf Coast | Col 3 = Galveston
#    Row 1 = Indegree vs Outdegree
#    Row 2 = Closeness vs Betweenness
#    Tags A-F (row-major). One shared "Concept type" legend, bottom,
#    centred.
#
#  FIGURE 2 -- 1 row x 3 columns, prominence ranking
#    Col 1 = Australia   (aggregate + 4 regions)
#    Col 2 = US Gulf Coast (aggregate + 5 states)
#    Col 3 = Galveston   (aggregate + 3 fishery sectors)
#    Tags A-C. One shared "Group" legend, bottom, centred.
#
#  Set OUT_DIR below, then source the whole script.
# ============================================================


# ============================================================
# SECTION 0 — PACKAGES
# ============================================================

required_pkgs <- c("readxl", "igraph", "dplyr", "tidyr", "ggplot2",
                   "stringr", "scales", "ggrepel", "purrr",
                   "tidytext", "ggnewscale", "patchwork")
for (pkg in required_pkgs) {
  if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
}
library(readxl); library(igraph); library(dplyr); library(tidyr)
library(ggplot2); library(stringr); library(scales)
library(ggrepel); library(purrr); library(tidytext); library(ggnewscale)
library(patchwork)

# Folder the combined figures are saved into. Defaults to the working
# directory (mirrors the bare-filename ggsave() calls in the three
# source scripts, which assume you've already setwd() into R_Analysis).
OUT_DIR <- getwd()


# ============================================================
# SECTION 1 — SHARED CONSTANTS (identical across the 3 source scripts)
# ============================================================

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


# ============================================================
# SECTION 2 — SHARED HELPER FUNCTIONS
#   (graph building + metric computation; identical logic in all
#    three source scripts, defined once here)
# ============================================================

# 2.1 Filter any data frame by a Tags column entry.
#     tag_name = NULL means "no filtering" -> the aggregate model.
filter_by_tag <- function(df, tag_name) {
  if (is.null(tag_name)) return(df)
  df %>% filter(!is.na(Tags), str_detect(Tags, fixed(tag_name)))
}

# 2.2 Build a deduplicated edge list for one unit (aggregate or subunit).
#     Duplicate (From, To) pairs (different workshops) are averaged
#     into one edge weight per pair.
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

# 2.3 Build an igraph directed weighted graph for one unit.
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

# 2.4 Weighted indegree / outdegree / degree
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

# 2.5 Concept type classification
classify_concepts <- function(node_df) {
  node_df %>%
    mutate(concept_type = case_when(
      indegree == 0 & outdegree  > 0 ~ "Transmitter",
      indegree  > 0 & outdegree == 0 ~ "Receiver",
      indegree  > 0 & outdegree  > 0 ~ "Ordinary",
      TRUE                           ~ "Isolated"))
}

# 2.6 Simple-cycle finder (depth-first search, Johnson 1975)
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

# 2.7 Classify a feedback loop as Reinforcing (+1) or Balancing (-1)
classify_loop_polarity <- function(g, cycle) {
  pol <- 1
  for (k in seq_len(length(cycle) - 1)) {
    eid <- get.edge.ids(g, c(cycle[k], cycle[k + 1]))
    if (eid == 0) return(NA)
    pol <- pol * sign(E(g)$weight[eid])
  }
  pol
}

# 2.8 Prominence (Hoffman et al., 2014) -- eigenvector centrality on the
#     UNDIRECTED graph so it stays comparable and non-zero across
#     aggregate models and small acyclic subunit models.
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

# 2.9 Full metric computation for one unit graph
compute_all_metrics <- function(g, model_name, verbose = TRUE) {
  if (is.null(g)) return(NULL)
  
  N       <- vcount(g)
  E_count <- ecount(g)
  density <- if (N > 1) E_count / (N * (N - 1)) else 0
  
  node_df <- compute_degrees(g) %>% classify_concepts()
  n_tx  <- sum(node_df$concept_type == "Transmitter")
  n_rx  <- sum(node_df$concept_type == "Receiver")
  n_ord <- sum(node_df$concept_type == "Ordinary")
  n_iso <- sum(node_df$concept_type == "Isolated")
  RT    <- if (n_tx > 0) n_rx / n_tx else NA_real_
  
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
  
  edge_df <- igraph::as_data_frame(g, what = "edges")
  n_pos   <- sum(edge_df$polarity > 0, na.rm = TRUE)
  n_neg   <- sum(edge_df$polarity < 0, na.rm = TRUE)
  
  dist_mat  <- distances(g, weights = NA)
  fin_d     <- dist_mat[is.finite(dist_mat) & dist_mat > 0]
  apl  <- if (length(fin_d) > 0) mean(fin_d)                     else NA_real_
  diam <- if (length(fin_d) > 0) max(dist_mat[is.finite(dist_mat)]) else NA_real_
  
  clust_avg <- transitivity(g, type = "average", isolates = "zero")
  clust_glo <- transitivity(g, type = "global")
  
  n_comp    <- components(g, mode = "weak")$no
  circ_rank <- E_count - N + n_comp
  
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
    top5 <- node_df %>% arrange(desc(prominence)) %>% head(5) %>% pull(name)
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

# 2.10 Build the aggregate-unit node table for one region, with Type
#      locked to the full TYPE_COLOURS factor level set (so
#      scale_colour_manual(drop = FALSE) keeps a consistent legend
#      across regions and panels).
build_agg_df <- function(metrics_list, aggregate_label) {
  m <- metrics_list[[aggregate_label]]
  m$node_df %>%
    rename(Concept = name, Type = concept_type, Degree = degree,
           Indegree = indegree, Outdegree = outdegree,
           Closeness = closeness, Betweenness = betweenness) %>%
    mutate(Type = factor(Type, levels = names(TYPE_COLOURS)))
}


# ============================================================
# SECTION 3 — LOAD DATA, BUILD GRAPHS, COMPUTE METRICS, PER REGION
# ============================================================

## ---- 3.1 Australia (national scale) ----

AU_FILE <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Australia/Kumu/Exported_Kumu_Australia_7May.xlsx"

au_elements    <- read_excel(AU_FILE, sheet = "Elements")
au_connections <- read_excel(AU_FILE, sheet = "Connections") %>%
  mutate(strength_num = suppressWarnings(as.numeric(Strength))) %>%
  filter(!is.na(From), !is.na(To), !is.na(strength_num))

au_groups <- au_elements %>%
  select(name = Label, Group = Type) %>%
  distinct() %>%
  filter(!is.na(Group))

AU_REGIONS       <- c("Western Australia", "North Australia",
                      "Queensland",        "New South Wales")
AU_AGGREGATE     <- "Australia"
AU_UNITS         <- c(AU_AGGREGATE, AU_REGIONS)

cat("\n=== Building Australia unit graphs (aggregate + regions) ===\n")
au_graphs <- setNames(
  c(list(build_unit_graph(au_elements, au_connections, tag_name = NULL)),
    lapply(AU_REGIONS, build_unit_graph,
           all_elements = au_elements, all_connections = au_connections)),
  AU_UNITS)

cat("\n========== AUSTRALIA -- METRICS ==========\n")
au_metrics <- setNames(
  lapply(AU_UNITS, function(u) compute_all_metrics(au_graphs[[u]], u)),
  AU_UNITS)

au_agg_df <- build_agg_df(au_metrics, AU_AGGREGATE)


## ---- 3.2 US Gulf Coast (regional scale) ----

GC_FILE <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Alabama/Kumu/Exported_Kumu_Alabama_30April.xlsx"

gc_elements    <- read_excel(GC_FILE, sheet = "Elements")
gc_connections <- read_excel(GC_FILE, sheet = "Connections") %>%
  mutate(strength_num = suppressWarnings(as.numeric(Strength))) %>%
  filter(!is.na(From), !is.na(To), !is.na(strength_num))

gc_groups <- gc_elements %>%
  select(name = Label, Group = Category) %>%
  distinct() %>%
  filter(!is.na(Group))

GC_STATES     <- c("Alabama", "Florida", "Louisiana", "Mississippi", "Texas")
GC_AGGREGATE  <- "US Gulf Coast"
GC_UNITS      <- c(GC_AGGREGATE, GC_STATES)

cat("\n=== Building US Gulf Coast unit graphs (aggregate + states) ===\n")
gc_graphs <- setNames(
  c(list(build_unit_graph(gc_elements, gc_connections, tag_name = NULL)),
    lapply(GC_STATES, build_unit_graph,
           all_elements = gc_elements, all_connections = gc_connections)),
  GC_UNITS)

cat("\n========== US GULF COAST -- METRICS ==========\n")
gc_metrics <- setNames(
  lapply(GC_UNITS, function(u) compute_all_metrics(gc_graphs[[u]], u)),
  GC_UNITS)

gc_agg_df <- build_agg_df(gc_metrics, GC_AGGREGATE)


## ---- 3.3 Galveston (local scale) ----

GV_FILE <- "~/Library/CloudStorage/GoogleDrive-paula.dominguez@arratiakomusikaeskola.eu/My Drive/ACTUAL/PhD/Projects/Depredation/MentalModels_Analysis/Galveston/Kumu/Exported_Kumu_Galveston_30April.xlsx"

gv_elements    <- read_excel(GV_FILE, sheet = "Elements")
gv_connections <- read_excel(GV_FILE, sheet = "Connections") %>%
  mutate(strength_num = suppressWarnings(as.numeric(Strength))) %>%
  filter(!is.na(From), !is.na(To), !is.na(strength_num))

gv_strength_vals <- sort(unique(gv_connections$strength_num))
if (any(abs(gv_strength_vals) > 0 & abs(gv_strength_vals) < 1))
  warning("Galveston Strength contains |value| < 1 (looks like the +-0.5/+-1.0 ",
          "Mental Modeler scale). Multiply by 2 to match the integer +-1/+-2 ",
          "scale used for Australia and Gulf Coast before trusting weighted metrics.")

gv_group_col <- if ("Category" %in% names(gv_elements)) "Category" else "Type"
gv_groups <- gv_elements %>%
  select(name = Label, Group = all_of(gv_group_col)) %>%
  distinct() %>%
  filter(!is.na(Group))

GV_SECTORS    <- c("Recreational", "Charter", "Commercial")
GV_AGGREGATE  <- "Galveston"
GV_UNITS      <- c(GV_AGGREGATE, GV_SECTORS)

cat("\n=== Building Galveston unit graphs (aggregate + fishery sectors) ===\n")
gv_graphs <- setNames(
  c(list(build_unit_graph(gv_elements, gv_connections, tag_name = NULL)),
    lapply(GV_SECTORS, build_unit_graph,
           all_elements = gv_elements, all_connections = gv_connections)),
  GV_UNITS)

cat("\n========== GALVESTON -- METRICS ==========\n")
gv_metrics <- setNames(
  lapply(GV_UNITS, function(u) compute_all_metrics(gv_graphs[[u]], u)),
  GV_UNITS)

gv_agg_df <- build_agg_df(gv_metrics, GV_AGGREGATE)


# ============================================================
# SECTION 4 — FIGURE 1 BUILDING BLOCKS
#   Indegree/Outdegree and Closeness/Betweenness plots, aggregate
#   (full) model only, one function reused across the 3 regions.
#   Per-region axis limits / label placement match the tuning in the
#   three source scripts.
# ============================================================

# 4.1 Indegree vs Outdegree
#     label_nudges: optional named list of c(dx, dy), e.g.
#     list("Shark Habituation" = c(8, 4)), giving ggrepel a starting push
#     toward empty space for a specific point whose label keeps landing on
#     top of a neighbour -- the leader line still points back to the real
#     data point, only the box's resting position is nudged.
build_io_plot <- function(df, xlim, ylim, label_frac, label_angle,
                          title, show_legend, show_y_label = TRUE, p_keep = 0.80,
                          label_size = 4, wrap_width = 20,
                          box_padding = 0.4, repel_force = 1,
                          repel_max_iter = 10000, label_nudges = NULL) {
  
  label_df <- df %>% filter(Degree >= quantile(Degree, p_keep, na.rm = TRUE))
  
  nudge_x_vec <- rep(0, nrow(label_df))
  nudge_y_vec <- rep(0, nrow(label_df))
  for (nm in names(label_nudges)) {
    idx <- which(label_df$Concept == nm)
    if (length(idx) > 0) {
      nudge_x_vec[idx] <- label_nudges[[nm]][1]
      nudge_y_vec[idx] <- label_nudges[[nm]][2]
    }
  }
  
  io_diag_max  <- min(xlim[2], ylim[2])
  io_label_pos <- label_frac * io_diag_max
  
  # Dummy off-panel points guaranteeing every non-Isolated Type has a real
  # data row behind its legend key -- Australia's aggregate model has zero
  # Receiver-type nodes, and relying on drop=FALSE/limits/breaks alone for
  # a totally-absent factor level was silently dropping its colour swatch.
  legend_dummy <- data.frame(
    Indegree = -Inf, Outdegree = -Inf, Degree = 1,
    Type = factor(setdiff(names(TYPE_COLOURS), "Isolated"),
                  levels = names(TYPE_COLOURS)))
  
  p <- ggplot(df, aes(x = Indegree, y = Outdegree,
                      colour = Type, size = Degree)) +
    geom_abline(slope = 1, intercept = 0,
                linetype = "dashed", colour = "grey60", linewidth = 0.5) +
    geom_point(alpha = 0.6, stroke = 0.2) +
    geom_point(data = legend_dummy, alpha = 0, size = 4, show.legend = show_legend) +
    geom_label_repel(
      data          = label_df,
      aes(label     = str_wrap(Concept, wrap_width)),
      size          = label_size,
      label.padding = unit(0.10, "lines"),
      box.padding   = unit(box_padding, "lines"),
      point.padding = unit(0.2, "lines"),
      nudge_x = nudge_x_vec, nudge_y = nudge_y_vec,
      min.segment.length = 0,
      segment.colour = "grey55", segment.size = 0.3,
      max.overlaps  = Inf, seed = 42, show.legend = FALSE,
      force = repel_force, force_pull = 0.3, max.iter = repel_max_iter,
      max.time = 5, direction = "both",
      colour        = "grey15", fill = alpha("white", 0.85)
    ) +
    annotate("text", x = io_label_pos, y = io_label_pos, label = "1:1 line",
             angle = label_angle, vjust = -1, size = 3, colour = "grey45") +
    scale_colour_manual(values = TYPE_COLOURS, drop = FALSE,
                        limits = names(TYPE_COLOURS),
                        breaks = setdiff(names(TYPE_COLOURS), "Isolated")) +
    guides(colour = if (show_legend)
      guide_legend(override.aes = list(size = 4, alpha = 1, shape = 16)) else "none") +
    scale_size_continuous(range = c(1.5, 6), guide = "none") +
    coord_cartesian(xlim = xlim, ylim = ylim) +
    labs(title = title, x = "Indegree",
         y = if (show_y_label) "Outdegree" else NULL,
         colour = "Concept type") +
    theme_fcm() +
    theme(
      plot.title   = element_text(size = 16, face = "bold", hjust = 0.5),
      axis.title   = element_text(size = 13),
      legend.title = element_text(size = 12),
      legend.text  = element_text(size = 12)
    )
  p
}

# 4.2 Closeness vs Betweenness
#     Labelled points are chosen either by a fixed top-N rank (p_keep_n,
#     the default) or, if p_keep is supplied, by the same top-percentile-
#     of-Degree logic build_io_plot uses -- so a panel can be tuned with
#     the identical "top 15%" language as the indegree/outdegree panels.
build_cb_plot <- function(df, xlim = NULL, ylim = NULL, title = NULL,
                          show_legend, show_y_label = TRUE, p_keep_n = 10,
                          p_keep = NULL,
                          label_size = 4, wrap_width = 16,
                          box_padding = 0.5, repel_force = 1,
                          repel_max_iter = 10000) {
  
  plot_df <- df %>%
    filter(!is.na(Closeness), !is.na(Betweenness)) %>%
    mutate(
      keep = if (!is.null(p_keep)) {
        Degree >= quantile(Degree, p_keep, na.rm = TRUE)
      } else {
        rank(-Closeness,   ties.method = "min") <= p_keep_n |
          rank(-Betweenness, ties.method = "min") <= p_keep_n
      },
      lab  = ifelse(keep, str_wrap(Concept, wrap_width), "")
    )
  
  quad_lines <- plot_df %>%
    summarise(med_closeness   = median(Closeness,   na.rm = TRUE),
              med_betweenness = median(Betweenness, na.rm = TRUE))
  
  p <- ggplot(plot_df, aes(x = Closeness, y = Betweenness)) +
    geom_vline(data = quad_lines, aes(xintercept = med_closeness),
               linetype = "dashed", colour = "grey70", linewidth = 0.4) +
    geom_hline(data = quad_lines, aes(yintercept = med_betweenness),
               linetype = "dashed", colour = "grey70", linewidth = 0.4) +
    geom_point(aes(colour = Type, size = Degree), alpha = 0.6, stroke = 0.2) +
    geom_label_repel(
      data = plot_df, aes(label = lab),
      size = label_size, label.padding = unit(0.10, "lines"),
      box.padding = unit(box_padding, "lines"), point.padding = unit(0.4, "lines"),
      min.segment.length = 0, segment.colour = "grey55", segment.size = 0.3,
      max.overlaps = Inf, seed = 42, show.legend = FALSE,
      force = repel_force, force_pull = 0.3, max.iter = repel_max_iter,
      max.time = 5, direction = "both",
      colour = "grey15", fill = alpha("white", 0.85)
    ) +
    scale_colour_manual(values = TYPE_COLOURS, drop = FALSE,
                        limits = names(TYPE_COLOURS),
                        breaks = setdiff(names(TYPE_COLOURS), "Isolated")) +
    guides(colour = if (show_legend)
      guide_legend(override.aes = list(size = 4, alpha = 1, shape = 16)) else "none") +
    scale_size_continuous(range = c(1.5, 6), guide = "none") +
    scale_x_sqrt() +
    scale_y_sqrt() +
    labs(title = title, x = "Closeness",
         y = if (show_y_label) "Betweenness" else NULL,
         colour = "Concept type") +
    theme_fcm() +
    theme(
      plot.title   = element_text(size = 16, face = "bold", hjust = 0.5),
      axis.title   = element_text(size = 13),
      legend.title = element_text(size = 12),
      legend.text  = element_text(size = 12)
    )
  
  if (!is.null(xlim) || !is.null(ylim))
    p <- p + coord_cartesian(xlim = xlim, ylim = ylim)
  
  p
}


# ============================================================
# SECTION 5 — FIGURE 1: ASSEMBLE 3 x 2 PANEL
#   Only the Australia Indegree/Outdegree panel (top-left, tag A)
#   keeps its "Concept type" legend -- patchwork collects it into a
#   single shared legend at the bottom of the whole figure.
# ============================================================

# A and D reverted to the same top-15%-of-Degree cutoff as B (p_keep = 0.85),
# keeping the stronger repel push to help the higher label count stay legible.
p_au_io <- build_io_plot(au_agg_df, xlim = c(0, 20),  ylim = c(0, 15),
                         label_frac = 0.65, label_angle = 45,
                         title = "Australia", show_legend = TRUE,
                         p_keep = 0.85, box_padding = 0.9, repel_force = 5,
                         repel_max_iter = 50000,
                         label_nudges = list("Shark Habituation" = c(8, 4)))
p_gc_io <- build_io_plot(gc_agg_df, xlim = c(0, 22),  ylim = c(0, 20),
                         label_frac = 0.90, label_angle = 35,
                         title = "US Gulf Coast", show_legend = FALSE,
                         show_y_label = FALSE,
                         p_keep = 0.85, box_padding = 0.6, repel_force = 2,
                         repel_max_iter = 30000)
# Galveston has far fewer nodes than AU/GC, so the default 80th-percentile
# cutoff was labelling most points and swamping the panel -- keep only the
# top 15% by Degree (same cutoff as A/B), with smaller/tighter boxes and a
# stronger repel push.
p_gv_io <- build_io_plot(gv_agg_df, xlim = c(0, 7),   ylim = c(0, 7),
                         label_frac = 0.25, label_angle = 35,
                         title = "Galveston", show_legend = FALSE,
                         show_y_label = FALSE,
                         p_keep = 0.85, label_size = 4, wrap_width = 13,
                         box_padding = 0.7, repel_force = 3,
                         repel_max_iter = 30000)

p_au_cb <- build_cb_plot(au_agg_df, xlim = c(0.40, 0.75), ylim = c(0, 0.12),
                         show_legend = FALSE,
                         p_keep_n = 6, box_padding = 0.9, repel_force = 5,
                         repel_max_iter = 50000)
p_gc_cb <- build_cb_plot(gc_agg_df, xlim = NULL, ylim = NULL,
                         show_legend = FALSE, show_y_label = FALSE,
                         p_keep_n = 8, box_padding = 0.6, repel_force = 2,
                         repel_max_iter = 30000)
# Galveston has far fewer nodes than AU/GC, so top-10 was labelling almost
# every point and swamping the panel -- fewer labels, smaller/tighter boxes,
# and a stronger repel push so what remains stays legible.
p_gv_cb <- build_cb_plot(gv_agg_df, xlim = c(0.25, 1.25), ylim = c(0, 0.05),
                         show_legend = FALSE, show_y_label = FALSE,
                         p_keep_n = 7, label_size = 4, wrap_width = 13,
                         box_padding = 0.7, repel_force = 3,
                         repel_max_iter = 30000)

fig1 <- (p_au_io | p_gc_io | p_gv_io) / (p_au_cb | p_gc_cb | p_gv_cb) +
  plot_layout(guides = "collect") +
  plot_annotation(tag_levels = "A") &
  theme(
    legend.position    = "bottom",
    legend.justification = "center",
    legend.box.just    = "center",
    plot.tag           = element_text(size = 14, face = "bold")
  )

ggsave(file.path(OUT_DIR, "Figure1_Indegree_Closeness_AU_GC_GV.png"), fig1,
       width = 22, height = 12, units = "in", dpi = 1200)
ggsave(file.path(OUT_DIR, "Figure1_Indegree_Closeness_AU_GC_GV.pdf"), fig1,
       width = 31, height = 21, units = "in", device = cairo_pdf)
ggsave(file.path(OUT_DIR, "Figure1_Indegree_Closeness_AU_GC_GV.tiff"), fig1,
       width = 31, height = 21, units = "in", dpi = 600, compression = "lzw")


# ============================================================
# SECTION 6 — FIGURE 2: PROMINENCE PANEL (1 row x 3 columns)
#   Each column reuses the region's own full prominence-ranking
#   heatmap (aggregate model + all its subunits, one column per
#   subunit within the panel -- already "complete" per source
#   scripts). Only the Australia panel keeps its "Group" legend;
#   the numeric rank colour bar is suppressed everywhere since the
#   rank is already printed as text in each tile -- this leaves
#   exactly one shared legend for the whole figure.
# ============================================================

plot_prominence_rank <- function(metrics_list, label, groups_df = NULL,
                                 top_n = 10, col_order = names(metrics_list),
                                 show_rank_legend = TRUE,
                                 show_group_legend = TRUE,
                                 label_map = NULL) {
  
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
      guide    = if (show_rank_legend)
        guide_colourbar(position = "right", reverse = TRUE) else "none") +
    
    ggnewscale::new_scale_fill() +
    geom_point(aes(fill = Group), shape = 22, size = 0, alpha = 0, stroke = 0) +
    scale_fill_manual(values = group_colours_used, name = NULL, drop = FALSE,
                      guide = if (show_group_legend)
                        guide_legend(position = "bottom", override.aes = list(size = 5, alpha = 1))
                      else "none") +
    
    scale_x_discrete(
      expand = expansion(add = c(0.6, 0.3)),
      labels = function(x) {
        renamed <- if (is.null(label_map)) x else
          unname(ifelse(x %in% names(label_map), label_map[x], x))
        # Wrap onto 2 lines instead of angling the text, so long names
        # (e.g. "New South Wales") fit within their column width while
        # staying horizontal and easy to read.
        str_wrap(renamed, width = 10)
      }
    ) +
    labs(title = label, x = NULL, y = NULL) +
    theme_fcm() +
    theme(
      plot.title    = element_text(size = 15, face = "bold", hjust = 0.5),
      axis.text     = element_text(size = 11),
      axis.text.y   = element_text(size = 11, face = "bold", colour = label_colours,
                                   hjust = 1, margin = margin(r = 6)),
      axis.text.x   = element_text(size = 11, face = "bold",
                                   angle = 0, hjust = 0.5, vjust = 1,
                                   lineheight = 0.9),
      legend.title  = element_text(size = 12, hjust = 0.5, margin = margin(b = 14)),
      legend.text   = element_text(size = 12),
      panel.grid.major.x = element_blank(),
      panel.grid.minor   = element_blank()
    )
}

prom_AU <- plot_prominence_rank(
  au_metrics, "Australia", groups_df = au_groups, col_order = AU_UNITS,
  show_rank_legend = FALSE, show_group_legend = TRUE,
  label_map = c("Western Australia" = "Western Aus.",
                "North Australia"   = "North Aus."))

prom_GC <- plot_prominence_rank(
  gc_metrics, "US Gulf Coast", groups_df = gc_groups, col_order = GC_UNITS,
  show_rank_legend = FALSE, show_group_legend = FALSE)

prom_GV <- plot_prominence_rank(
  gv_metrics, "Galveston", groups_df = gv_groups, col_order = GV_UNITS,
  show_rank_legend = FALSE, show_group_legend = FALSE)

fig2 <- (prom_AU | prom_GC | prom_GV) +
  plot_layout(guides = "collect") +
  plot_annotation(tag_levels = "A") &
  theme(
    legend.position     = "bottom",
    legend.justification = "center",
    legend.box.just     = "center",
    plot.tag            = element_text(size = 14, face = "bold")
  )

ggsave(file.path(OUT_DIR, "Figure2_Prominence_AU_GC_GV.png"), fig2,
       width = 26, height = 17, units = "in", dpi = 1200)
ggsave(file.path(OUT_DIR, "Figure2_Prominence_AU_GC_GV.pdf"), fig2,
       width = 28, height = 18, units = "in", device = cairo_pdf)

cat("\n\nCombined figures complete. Saved to:", OUT_DIR, "\n")
