# ============================================================================
# DIAGNOSTIC TOOL: BATCH PLOT TRAVERSAL AUDIT (Single & Multi-Tree Compatible)
# ============================================================================
library(tidyverse)
library(fs)

# ============================================================================
#### 1. USER CONFIGURATION ####
# ============================================================================
RAW_WIDE_DATA_CSV <- "C:/Users/james.baker/Forest Research/TW CBC-TBA-NextGenBritishConifers - Share/Sitka/High GCA Fullsib P85-P87 experiments/Craigellachie 49/Craigellachie_49_Full_Data_With_Flags.csv"

PLOT_TYPE      <- "MULTI"  # Options: "SINGLE" (evaluates Plot sequence within Blocks) or "MULTI" (Trees within Plots)

# FALLBACK GRID DIMENSIONS (Used only if spatial Prow/Ppos are missing from the data)
GRID_ROWS      <- 8
GRID_COLS      <- 1

BASELINE_TRAIT <- "Ht_10" 
TEST_TRAITS    <- c("Ht_03","Ht_06","Cr_07","Pil_15") 

USE_CORRECTED_DATA <- FALSE

# --- Auto-Path Logic & Subdirectory Management ---
if(USE_CORRECTED_DATA) {
  WIDE_DATA_CSV <- stringr::str_replace(RAW_WIDE_DATA_CSV, "(?i)\\.csv$", "_Corrected.csv")
  out_suffix <- "_Chained"
} else {
  WIDE_DATA_CSV <- RAW_WIDE_DATA_CSV
  out_suffix <- ""
}

OUTPUT_DIR <- dirname(WIDE_DATA_CSV)
DIAG_DIR   <- file.path(OUTPUT_DIR, "Traversal_Diagnostics")
if(!dir.exists(DIAG_DIR)) dir.create(DIAG_DIR)

helper_path <- file.path(OUTPUT_DIR, "TRAVERSAL_HELPER_MASTER.csv")
if (USE_CORRECTED_DATA && file.exists(helper_path)) {
  backup_name <- paste0("TRAVERSAL_HELPER_MASTER_PreChainBackup_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".csv")
  file.copy(helper_path, file.path(DIAG_DIR, backup_name))
  message(paste("Backed up original Master Helper to:", backup_name))
}
if (USE_CORRECTED_DATA && file.exists(helper_path)) {
  existing_helper <- suppressMessages(read_csv(helper_path, show_col_types = FALSE))
  conflict_traits <- intersect(TEST_TRAITS, existing_helper$Trait_ID)
  if (length(conflict_traits) > 0) stop("CRITICAL ERROR: CHAINING SAFEGUARD TRIGGERED.")
}

# ============================================================================
#### 2. VIRTUAL ASSESSOR MATH ENGINE ####
# ============================================================================
get_grid_layout <- function(n_rows, n_cols, interior_only = FALSE) {
  if (n_rows == 1 && n_cols == 8) return(matrix(1:8, nrow=1))
  mat <- matrix(1:(n_rows*n_cols), nrow=n_rows, byrow=TRUE)
  if(interior_only && n_rows >= 3 && n_cols >= 3) mat <- mat[2:(n_rows-1), 2:(n_cols-1)]
  return(mat)
}

get_traversal_path <- function(layout_mat, start_corner="top_left", direction="horizontal", snake=FALSE) {
  if (length(layout_mat) == 8) {
    if (start_corner %in% c("reversed", "bottom_right")) return(rev(as.vector(layout_mat)))
    return(as.vector(layout_mat))
  }
  n_rows <- nrow(layout_mat)
  n_cols <- ncol(layout_mat)
  row_indices <- if (start_corner %in% c("bottom_left", "bottom_right")) rev(seq_len(n_rows)) else seq_len(n_rows)
  col_indices <- if (start_corner %in% c("top_right", "bottom_right")) rev(seq_len(n_cols)) else seq_len(n_cols)
  
  path <- c()
  if (direction == "horizontal") {
    for (i_idx in seq_along(row_indices)) {
      i <- row_indices[i_idx]; cols <- col_indices
      if (snake && i_idx %% 2 == 0) cols <- rev(cols)
      for (j in cols) path <- c(path, layout_mat[i, j])
    }
  } else {
    for (j_idx in seq_along(col_indices)) {
      j <- col_indices[j_idx]; rows <- row_indices
      if (snake && j_idx %% 2 == 0) rows <- rev(rows)
      for (i in rows) path <- c(path, layout_mat[i, j])
    }
  }
  return(path)
}

# ============================================================================
#### 3. BATCH PROCESSING LOOP ####
# ============================================================================
message(paste("Loading master data from:", basename(WIDE_DATA_CSV)))
df_raw <- read_csv(WIDE_DATA_CSV, show_col_types = FALSE)

# ALIASING ENGINE
if (PLOT_TYPE == "SINGLE") {
  if (!"Block" %in% names(df_raw)) stop("CRITICAL ERROR: 'Block' column required.")
  message(">> RUNNING IN SINGLE-TREE MODE: Evaluating Plot sequences within Blocks.")
  df_raw <- df_raw %>%
    rename(Real_Plot = Plot) %>%
    mutate(Plot = as.character(Block), Tree = as.numeric(Real_Plot)) %>%
    group_by(Plot) %>%
    mutate(Tree = Tree - min(Tree) + 1) %>%
    ungroup()
}

# EXTRACT GLOBAL SPATIAL COORDINATES (Draws ALL trees, even if blocks drop out)
has_standard <- "Row" %in% names(df_raw) && "Position" %in% names(df_raw)
has_p_coords <- "Prow" %in% names(df_raw) && "Ppos" %in% names(df_raw)

if (has_standard || has_p_coords) {
  plot_spatial_trees <- df_raw %>%
    mutate(
      Plot = as.character(Plot), 
      Map_Row = if(has_standard) as.numeric(Row) else as.numeric(Prow),
      Map_Pos = if(has_standard) as.numeric(Position) else as.numeric(Ppos)
    ) %>%
    filter(!is.na(Map_Row) & !is.na(Map_Pos))
}

for (TEST_TRAIT in TEST_TRAITS) {
  cat("\n======================================================================\n")
  cat(">>> INITIATING DIAGNOSTIC FOR TRAIT:", TEST_TRAIT, "<<<\n")
  cat("======================================================================\n")
  
  # --- Auto-Detect Survival ---
  base_age <- suppressWarnings(as.numeric(str_extract(BASELINE_TRAIT, "\\d+")))
  target_surv_base <- paste0("Sur_", str_pad(base_age, 2, pad = "0"))
  if (target_surv_base %in% names(df_raw)) BASELINE_SURV <- target_surv_base else {
    surv_cols <- grep("(?i)^Sur_", names(df_raw), value = TRUE)
    if (length(surv_cols) > 0) BASELINE_SURV <- surv_cols[which.min(abs(suppressWarnings(as.numeric(str_extract(surv_cols, "\\d+"))) - base_age))] else stop("No survival traits found.")
  }
  
  test_age <- suppressWarnings(as.numeric(str_extract(TEST_TRAIT, "\\d+")))
  target_surv_test <- paste0("Sur_", str_pad(test_age, 2, pad = "0"))
  if (target_surv_test %in% names(df_raw)) TEST_SURV <- target_surv_test else {
    surv_cols <- grep("(?i)^Sur_", names(df_raw), value = TRUE)
    if (length(surv_cols) > 0) TEST_SURV <- surv_cols[which.min(abs(suppressWarnings(as.numeric(str_extract(surv_cols, "\\d+"))) - test_age))] else stop("No survival traits found.")
  }
  
  working_data <- df_raw %>%
    select(Plot, Tree, Base_Val = !!sym(BASELINE_TRAIT), Base_Surv = !!sym(BASELINE_SURV), Test_Val = !!sym(TEST_TRAIT), Test_Surv = !!sym(TEST_SURV)) %>%
    filter(!str_detect(Plot, "(?i)Filler")) %>%
    mutate(Tree = as.numeric(Tree), Plot = as.character(Plot), Base_Val = na_if(as.numeric(Base_Val), 0), Test_Val = na_if(as.numeric(Test_Val), 0))
  
  combos <- expand_grid(start_corner = c("top_left", "bottom_left", "top_right", "bottom_right"), direction = c("horizontal", "vertical"), snake = c(FALSE, TRUE)) %>% mutate(Perm_ID = row_number())
  
  plots <- unique(working_data$Plot)
  results_list <- list()
  
  message(paste("Testing", length(plots), ifelse(PLOT_TYPE=="SINGLE", "blocks", "plots"), "against permutations..."))
  
  for (p in plots) {
    plot_data <- working_data %>% filter(Plot == p)
    if(sum(!is.na(plot_data$Base_Val)) < 3 || sum(!is.na(plot_data$Test_Val)) < 3) next
    
    # --- DYNAMIC GRID RESIZING ---
    p_rows <- GRID_ROWS; p_cols <- GRID_COLS
    if (exists("plot_spatial_trees")) {
      sp_data <- plot_spatial_trees %>% filter(Plot == p)
      if(nrow(sp_data) > 0) {
        p_rows <- max(sp_data$Map_Row) - min(sp_data$Map_Row) + 1
        p_cols <- max(sp_data$Map_Pos) - min(sp_data$Map_Pos) + 1
      }
    }
    
    layout_mat <- get_grid_layout(p_rows, p_cols)
    canonical_path <- get_traversal_path(layout_mat, "top_left", "horizontal", FALSE)
    
    for (i in 1:nrow(combos)) {
      sc <- combos$start_corner[i]; dir <- combos$direction[i]; snk <- combos$snake[i]
      tested_path <- get_traversal_path(layout_mat, sc, dir, snk)
      translation_df <- tibble(Original_Tree = canonical_path, Mapped_Tree = tested_path)
      
      mapped_test <- plot_data %>% select(Tree, Test_Val, Test_Surv) %>% inner_join(translation_df, by = c("Tree" = "Original_Tree"))
      joined <- plot_data %>% select(Tree, Base_Val, Base_Surv) %>% left_join(mapped_test, by = c("Tree" = "Mapped_Tree"))
      
      zombie_count <- if (!is.na(base_age) && !is.na(test_age) && base_age > test_age) sum(joined$Test_Surv == 0 & (!is.na(joined$Base_Val) | joined$Base_Surv == 1), na.rm = TRUE) else sum(joined$Base_Surv == 0 & (!is.na(joined$Test_Val) | joined$Test_Surv == 1), na.rm = TRUE)
      
      is_valid <- (zombie_count == 0)
      cor_data <- joined %>% filter(!is.na(Base_Val) & !is.na(Test_Val))
      spearman_cor <- if(nrow(cor_data) >= 3) cor(cor_data$Base_Val, cor_data$Test_Val, method = "spearman") else NA_real_
      
      results_list[[length(results_list) + 1]] <- tibble(Plot = p, Perm_ID = combos$Perm_ID[i], start_corner = sc, direction = dir, snake = snk, Zombies_Created = zombie_count, Is_Biologically_Valid = is_valid, Sample_Size = nrow(cor_data), Spearman_Cor = spearman_cor)
    }
  }
  all_results <- bind_rows(results_list)
  
  # --- Synthesis & Recommendations ---
  message("Synthesizing recommendations...")
  original_recs <- all_results %>%
    group_by(Plot) %>%
    mutate(Normal_Cor = Spearman_Cor[start_corner %in% c("top_left", "normal") & direction == "horizontal" & snake == FALSE]) %>%
    filter(Zombies_Created == min(Zombies_Created)) %>% 
    arrange(desc(Spearman_Cor)) %>% 
    slice(1) %>% 
    ungroup() %>%
    mutate(Action_Required = case_when(
      start_corner %in% c("top_left", "normal") & direction == "horizontal" & snake == FALSE ~ "None (Normal is Best)",
      Zombies_Created > 0 ~ paste("FIX w/ WARNING:", Zombies_Created, "Zombies remain"),
      (Spearman_Cor - Normal_Cor) > 0.15 ~ "SUGGESTED FIX: Massive correlation improvement",
      TRUE ~ "Review Manually (Marginal Improvement)"
    ))
  
  # --- THE GLOBAL CONSENSUS OVERRIDE ---
  # 1. Tally up the suggested fixes
  fix_tally <- original_recs %>% 
    filter(str_detect(Action_Required, "(?i)SUGGESTED FIX|FIX w/ WARNING")) %>%
    mutate(Path_Name = paste(start_corner, direction, snake)) %>%
    count(Path_Name, start_corner, direction, snake, name = "Votes") %>%
    arrange(desc(Votes))
  
  total_plots <- length(unique(working_data$Plot))
  global_override_triggered <- FALSE
  holdout_plots <- c()
  
  # 2. If a single path accounts for > 60% of the entire trial, trigger!
  if (nrow(fix_tally) > 0 && (fix_tally$Votes[1] / total_plots) >= 0.60) {
    winning_path <- fix_tally[1, ]
    global_override_triggered <- TRUE
    message(sprintf("\n*** GLOBAL CONSENSUS DETECTED ***\n%s plots (%.1f%%) suggest '%s'. Overriding all plot-level filters to apply globally!\n", 
                    winning_path$Votes, (winning_path$Votes / total_plots)*100, winning_path$Path_Name))
    
    # Identify the holdouts (plots that did NOT organically suggest this fix)
    organic_winners <- original_recs %>% filter(paste(start_corner, direction, snake) == winning_path$Path_Name, str_detect(Action_Required, "(?i)FIX")) %>% pull(Plot)
    holdout_plots <- setdiff(unique(working_data$Plot), organic_winners)
    
    # Force every single plot in the trial to use the winning path
    recommendations <- tibble(Plot = unique(working_data$Plot)) %>%
      mutate(
        Action_Required = "GLOBAL CONSENSUS OVERRIDE",
        Best_Start = winning_path$start_corner,
        Best_Dir = winning_path$direction,
        Best_Snake = winning_path$snake,
        Best_Cor = NA_real_, Normal_Cor = NA_real_, Zombies_Created = NA_real_, Sample_Size = NA_real_
      )
  } else {
    recommendations <- original_recs %>%
      select(Plot, Action_Required, Best_Start = start_corner, Best_Dir = direction, Best_Snake = snake, Best_Cor = Spearman_Cor, Normal_Cor, Zombies_Created, Sample_Size) %>% 
      arrange(desc(Action_Required))
  }
  
  write_csv(all_results, file.path(DIAG_DIR, paste0(BASELINE_TRAIT, "_vs_", TEST_TRAIT, out_suffix, "_All_Permutations.csv")))
  write_csv(recommendations, file.path(DIAG_DIR, paste0("Suggested_Fixes_", TEST_TRAIT, "_anchored_to_", BASELINE_TRAIT, out_suffix, ".csv")))
  
  # --- VISUALIZATION BLOCK ---
  message("Generating visualization panels...")
  before_df <- working_data %>% select(Plot, Tree, Base_Val, Test_Val) %>% mutate(State = "1. Original Raw Data")
  after_list <- list()
  for (p in unique(working_data$Plot)) {
    p_data <- working_data %>% filter(Plot == p)
    rec <- recommendations %>% filter(Plot == p)
    if (nrow(rec) == 1 && !str_detect(rec$Action_Required, "(?i)None")) {
      
      # Use dynamic grid sizing
      p_rows <- GRID_ROWS; p_cols <- GRID_COLS
      if (exists("plot_spatial_trees")) {
        sp_data <- plot_spatial_trees %>% filter(Plot == p)
        if(nrow(sp_data) > 0) { p_rows <- max(sp_data$Map_Row) - min(sp_data$Map_Row) + 1; p_cols <- max(sp_data$Map_Pos) - min(sp_data$Map_Pos) + 1 }
      }
      layout_mat <- get_grid_layout(p_rows, p_cols)
      canonical_path <- get_traversal_path(layout_mat, "top_left", "horizontal", FALSE)
      
      tested_path <- get_traversal_path(layout_mat, rec$Best_Start, rec$Best_Dir, as.logical(rec$Best_Snake))
      translation_df <- tibble(Original_Tree = as.numeric(canonical_path), Mapped_Tree = as.numeric(tested_path))
      
      fixed_test <- p_data %>% select(Tree, Test_Val) %>% inner_join(translation_df, by = c("Tree" = "Original_Tree"))
      after_list[[length(after_list) + 1]] <- p_data %>% select(Plot, Tree, Base_Val) %>% left_join(fixed_test, by = c("Tree" = "Mapped_Tree"))
    } else {
      after_list[[length(after_list) + 1]] <- p_data %>% select(Plot, Tree, Base_Val, Test_Val)
    }
  }
  after_df <- bind_rows(after_list) %>% mutate(State = "2. Targeted Fixes")
  
  plot_df <- bind_rows(before_df, after_df) %>% filter(!is.na(Base_Val) & !is.na(Test_Val))
  calc_cor <- function(df) round(cor(df$Base_Val, df$Test_Val, use = "pairwise.complete.obs", method = "spearman"), 3)
  
  cor_before <- calc_cor(before_df)
  cor_after <- calc_cor(after_df)
  
  subtitle_text <- sprintf("Trial Correlations  ->  Raw: %s  |  Targeted Fix: %s", cor_before, cor_after)
  
  # Inject the Proof for the Holdout Cohort!
  if (global_override_triggered && length(holdout_plots) > 0) {
    holdout_before <- before_df %>% filter(Plot %in% holdout_plots)
    holdout_after <- after_df %>% filter(Plot %in% holdout_plots)
    
    # Only calculate if there are enough trees left to run a correlation
    if (sum(!is.na(holdout_before$Base_Val) & !is.na(holdout_before$Test_Val)) >= 3) {
      cor_h_before <- calc_cor(holdout_before)
      cor_h_after <- calc_cor(holdout_after)
      subtitle_text <- sprintf("%s\n[GLOBAL PROOF] Holdout Cohort (N=%s plots)  ->  Raw: %s  |  Forced Fix: %s", 
                               subtitle_text, length(holdout_plots), cor_h_before, cor_h_after)
    }
  }
  
  p_compare <- ggplot(plot_df, aes(x = Base_Val, y = Test_Val)) +
    geom_point(alpha = 0.5, size=0.25, color = "#2c3e50") +
    geom_smooth(method = "lm", formula = y ~ x, color = "#e74c3c", linetype = "dashed", se = FALSE) +
    facet_wrap(~State, ncol = 2) + theme_bw() +
    labs(title = paste("Plot Traversal Correction:", BASELINE_TRAIT, "vs", TEST_TRAIT), 
         subtitle = subtitle_text, 
         x = paste("Trusted Baseline:", BASELINE_TRAIT), y = paste("Suspected Trait:", TEST_TRAIT))
  
  ggsave(file.path(DIAG_DIR, paste0(BASELINE_TRAIT, "_vs_", TEST_TRAIT, out_suffix, "_Correction_Plot.png")), plot = p_compare, width = 12, height = 6, dpi = 300)
  
  # --- Helper Export ---
  fixes_to_apply <- recommendations %>% filter(str_detect(Action_Required, "FIX"))
  if (nrow(fixes_to_apply) > 0) {
    main_pipe_helper <- fixes_to_apply %>%
      mutate(Trial_ID = basename(OUTPUT_DIR), Trait_ID = TEST_TRAIT, Anchor_Used = BASELINE_TRAIT, Manual_Verification = "Pending", Best_Snake = as.character(Best_Snake), Plot = if (PLOT_TYPE == "SINGLE") paste0("Block_", Plot) else as.character(Plot)) %>%
      select(Trial_ID, Trait_ID, Anchor_Used, Plot, Best_Start, Best_Dir, Best_Snake, Best_Cor, Normal_Cor)
    helper_path <- file.path(OUTPUT_DIR, "TRAVERSAL_HELPER_MASTER.csv")
    if (file.exists(helper_path)) {
      existing_helper <- read_csv(helper_path, col_types = cols(.default = "c")) 
      combined_helper <- existing_helper %>% filter(Trait_ID != TEST_TRAIT) %>% bind_rows(main_pipe_helper %>% mutate(across(everything(), as.character)))
    } else { combined_helper <- main_pipe_helper }
    write_csv(combined_helper, helper_path)
    message(paste("Master Integration Helper updated at:", helper_path, "\n"))
  } else { message("No fixes required. Master Integration Helper unchanged.\n") }
}
cat(">>> BATCH DIAGNOSTIC COMPLETE <<<\n")