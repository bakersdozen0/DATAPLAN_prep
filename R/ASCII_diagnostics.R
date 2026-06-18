# ============================================================================
# ENGINE: UTILITIES & DIAGNOSTICS
# ============================================================================

# ----------------------------------------------------------------------------
# 1. Summarize ASCII Inventory
# ----------------------------------------------------------------------------
summarize_ascii_inventory <- function(base_dir) {
  file_list <- dir_ls(base_dir, recurse = TRUE, regexp = "(?i)ASCII\\.(xlsx|csv)$")
  message("Found ", length(file_list), " ASCII files to summarize.")
  
  trial_analysis <- file_list %>%
    map_df(function(file) {
      if (str_detect(file, "(?i)\\.csv$")) {
        df <- suppressMessages(read_csv(file, col_types = cols(.default = "c"))) %>% clean_names()
      } else {
        df <- suppressMessages(read_xlsx(file, col_types = "text")) %>% clean_names()
      }
      
      assess_col <- intersect(c("assessment", "assessment_type_long", "assessment_type"), names(df))[1]
      if (!is.na(assess_col)) df <- df %>% rename(assessment = !!sym(assess_col))
      
      year_col <- intersect(c("assessment_year", "assessment_year_long", "age"), names(df))[1]
      if (!is.na(year_col)) df <- df %>% rename(assessment_year = !!sym(year_col))
      
      if("plot" %in% names(df)) df <- df %>% rename(plot = plot)
      
      required_base <- c("assessment", "assessment_year", "plot")
      if (!all(required_base %in% names(df))) return(NULL)
      
      if ("inferred_tree_position" %in% names(df)) {
        df <- df %>% rename(tree_pos = inferred_tree_position)
      } else {
        df <- df %>% group_by(assessment, assessment_year, plot) %>% mutate(tree_pos = row_number()) %>% ungroup()
      }
      
      df %>%
        filter(!is.na(tree_pos)) %>%
        group_by(assessment, assessment_year, plot) %>%
        summarise(stems_measured = n_distinct(tree_pos), .groups = "drop") %>%
        group_by(assessment, assessment_year) %>%
        summarise(
          avg_stems_per_plot = mean(stems_measured, na.rm = TRUE),
          min_stems_in_a_plot = min(stems_measured),
          max_stems_in_a_plot = max(stems_measured),
          total_plots_count = n(),
          .groups = "drop"
        ) %>%
        mutate(experiment_file = basename(file))
    })
  
  final_report <- trial_analysis %>% select(experiment_file, assessment, assessment_year, avg_stems_per_plot, min_stems_in_a_plot, max_stems_in_a_plot, total_plots_count) %>% arrange(experiment_file, assessment_year)
  
  write_csv(final_report, file.path(base_dir, "MASTER_ASCII_Inventory_Report.csv"))
  message("ASCII Report saved successfully to ", base_dir)
}

# ----------------------------------------------------------------------------
# 2. Scan AV Duplicates
# ----------------------------------------------------------------------------
scan_av_duplicates <- function(base_dir) {
  file_list <- dir_ls(base_dir, recurse = TRUE, regexp = "(?i)ASCII\\.(xlsx|csv)$")
  message("Scanning ", length(file_list), " ASCII files for duplicate Av measurements...")
  
  duplicate_av_list <- list()
  
  for (file_path in file_list) {
    exp_name <- basename(dirname(file_path))
    tryCatch({
      if (str_detect(file_path, "(?i)\\.csv$")) {
        raw_data <- suppressMessages(read_csv(file_path, col_types = cols(.default = "c"))) %>% clean_names()
      } else {
        raw_data <- suppressMessages(read_excel(file_path, col_types = "text")) %>% clean_names()
      }
      
      assess_col <- intersect(c("assessment", "assessment_type_long", "assessment_type"), names(raw_data))[1]
      if (!is.na(assess_col)) raw_data <- raw_data %>% rename(assessment = !!sym(assess_col))
      
      year_col <- intersect(c("assessment_year", "assessment_year_long", "age"), names(raw_data))[1]
      if (!is.na(year_col)) raw_data <- raw_data %>% rename(assessment_year = !!sym(year_col))
      
      if("plot" %in% names(raw_data)) raw_data <- raw_data %>% rename(plot = plot)
      
      if ("inferred_tree_position" %in% names(raw_data)) {
        raw_data <- raw_data %>% rename(tree_pos = inferred_tree_position)
      } else {
        raw_data <- raw_data %>% group_by(assessment, assessment_year, plot) %>% mutate(tree_pos = row_number()) %>% ungroup()
      }
      
      req_cols <- c("plot", "tree_pos", "assessment", "assessment_year")
      if (all(req_cols %in% names(raw_data))) {
        duplicates <- raw_data %>%
          filter(str_detect(assessment, "(?i)^AV")) %>%
          group_by(plot, tree_pos, assessment, assessment_year) %>%
          summarise(measurement_count = n(), .groups = "drop") %>%
          filter(measurement_count > 1) 
        
        if (nrow(duplicates) > 0) duplicate_av_list[[exp_name]] <- duplicates %>% mutate(experiment = exp_name)
      }
    }, error = function(e) {})
  }
  
  if (length(duplicate_av_list) > 0) {
    all_duplicates <- bind_rows(duplicate_av_list) %>% select(Experiment = experiment, Plot = plot, Tree = tree_pos, Assessment = assessment, Age = assessment_year, Measurement_Count = measurement_count) %>% arrange(Experiment, as.numeric(Plot), as.numeric(Tree))
    out_path <- file.path(base_dir, "Repeat_Av_Scan_Results.csv")
    write_csv(all_duplicates, out_path)
    message("Found duplicates! Diagnostic report saved to: ", out_path)
  } else {
    message("Scan Complete! No repeat Av measurements found.")
  }
}

# ----------------------------------------------------------------------------
# 3. Extract Open Pollinated (OP) Instances
# ----------------------------------------------------------------------------
extract_op_instances <- function(pending_dir, base_dir) {
  cat("\nScanning Full Data files for OP instances...\n")
  full_data_files <- dir_ls(pending_dir, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")
  
  op_trial_data <- full_data_files %>%
    map_df(function(file) {
      df <- tryCatch(read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())), error = function(e) return(NULL))
      if (is.null(df)) return(NULL)
      fam_col <- grep("(?i)^family_name$", names(df), value = TRUE)
      if (length(fam_col) == 0) return(NULL)
      
      df %>% rename(Family_name = all_of(fam_col[1])) %>% select(Family_name) %>% distinct() %>% filter(str_detect(Family_name, "(?i)OP")) %>% mutate(Experiment_Name = str_replace_all(str_extract(basename(file), "^[^_]+"), " ", "_"), Source_File = basename(file), Data_Type = "Trial Data")
    })
  
  if (!is.null(op_trial_data) && nrow(op_trial_data) > 0) write_csv(op_trial_data, file.path(base_dir,"Diagnostic_TrialData_OP.csv"))
  message("OP instances extracted to ", base_dir)
}

# ----------------------------------------------------------------------------
# 4. Format Matrix Files
# ----------------------------------------------------------------------------
format_matrix_files <- function(root_dir) {
  all_dirs <- list.dirs(path = root_dir, recursive = FALSE, full.names = FALSE)
  experiments_to_check <- setdiff(all_dirs, c("00_Scripts", "Archive", ".git", ".Rproj.user"))
  
  for (curr_exp in experiments_to_check) {
    exp_path <- file.path(root_dir, curr_exp)
    source_file <- dir_ls(exp_path, regexp = "(?i)_matrix\\.xlsx$")
    if (length(source_file) == 0) next
    
    matrix_data <- tryCatch({ read_excel(source_file, sheet = "matrix", col_names = FALSE) }, 
                            error = function(e) { return(read_excel(source_file, col_names = FALSE)) })
    
    clean_matrix <- as.matrix(matrix_data)
    clean_matrix[is.na(clean_matrix)] <- 0
    dest_path <- file.path(exp_path, paste0(str_replace_all(curr_exp, " ", "_"), "_Matrix.csv"))
    
    write.table(clean_matrix, dest_path, row.names = FALSE, col.names = FALSE, sep = ",")
    message(paste("Saved clean matrix for:", curr_exp))
  }
}