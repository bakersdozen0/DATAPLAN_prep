# ============================================================================
# ENGINE: UTILITIES & DIAGNOSTICS
# ============================================================================

# ----------------------------------------------------------------------------
# 1. Summarize ASCII Inventory
# ----------------------------------------------------------------------------
summarize_ascii_inventory <- function(target_dir) {
  file_list <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)ASCII\\.(xlsx|csv)$")
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
  
  write_csv(final_report, file.path(target_dir, "MASTER_ASCII_Inventory_Report.csv"))
  message("ASCII Report saved successfully to ", target_dir)
}

# ----------------------------------------------------------------------------
# 2. Scan AV Duplicates
# ----------------------------------------------------------------------------
scan_duplicates <- function(target_dir) {
  file_list <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)ASCII\\.(xlsx|csv)$")
  message("Scanning ", length(file_list), " ASCII files for any duplicate trait measurements...")
  
  duplicate_master_list <- list()
  
  for (file_path in file_list) {
    exp_name <- basename(dirname(file_path))
    tryCatch({
      if (str_detect(file_path, "(?i)\\.csv$")) {
        raw_data <- suppressMessages(read_csv(file_path, col_types = cols(.default = "c"))) %>% clean_names()
      } else {
        raw_data <- suppressMessages(read_excel(file_path, col_types = "text")) %>% clean_names()
      }
      
      avail_cols <- names(raw_data)
      
      col_trait <- intersect(c("assessment", "assessment_type", "assessment_type_long", "trait"), avail_cols)[1]
      col_age   <- intersect(c("assessment_year", "age", "year_of_assessment", "assessment_yr"), avail_cols)[1]
      col_plot  <- intersect(c("plot", "plot_no", "treatment"), avail_cols)[1]
      col_tree  <- intersect(c("inferred_tree_position", "tree", "tree_no", "tree_pos"), avail_cols)[1]
      col_unit  <- intersect(c("unit", "units"), avail_cols)[1]
      col_year  <- intersect(c("year", "calendar_year"), avail_cols)[1]
      
      col_meas  <- intersect(c("measurement", "value", "score"), avail_cols)[1]
      col_rem   <- intersect(c("remarks", "remark", "notes", "comments"), avail_cols)[1]
      
      if (!is.na(col_trait)) raw_data <- raw_data %>% rename(assessment = !!sym(col_trait))
      if (!is.na(col_age))   raw_data <- raw_data %>% rename(assessment_year = !!sym(col_age))
      if (!is.na(col_plot))  raw_data <- raw_data %>% rename(plot = !!sym(col_plot))
      
      if (!is.na(col_unit)) { raw_data <- raw_data %>% rename(unit = !!sym(col_unit)) } else { raw_data <- raw_data %>% mutate(unit = NA_character_) }
      if (!is.na(col_year)) { raw_data <- raw_data %>% rename(year = !!sym(col_year)) } else { raw_data <- raw_data %>% mutate(year = NA_character_) }
      if (!is.na(col_meas)) { raw_data <- raw_data %>% rename(measurement = !!sym(col_meas)) } else { raw_data <- raw_data %>% mutate(measurement = NA_character_) }
      if (!is.na(col_rem))  { raw_data <- raw_data %>% rename(remarks = !!sym(col_rem)) } else { raw_data <- raw_data %>% mutate(remarks = NA_character_) }
      
      if (!is.na(col_tree)) {
        raw_data <- raw_data %>% rename(tree_pos = !!sym(col_tree))
      } else if (all(c("assessment", "assessment_year", "plot") %in% names(raw_data))) {
        raw_data <- raw_data %>% group_by(assessment, assessment_year, plot) %>% mutate(tree_pos = row_number()) %>% ungroup()
      }
      
      req_cols <- c("plot", "tree_pos", "assessment")
      if (all(req_cols %in% names(raw_data))) {
        
        # 1. Filter out NAs and identify the granular, true duplicates
        duplicates <- raw_data %>%
          filter(!is.na(assessment)) %>% # <-- NEW: Vaporizes the ghost rows!
          group_by(plot, tree_pos, assessment, unit, year, assessment_year) %>%
          summarise(
            measurement_count = n(), 
            Conflicting_Remarks = paste(na.omit(remarks), collapse = " vs "),
            .groups = "drop"
          ) %>%
          filter(measurement_count > 1) 
        
        if (nrow(duplicates) > 0) duplicate_master_list[[exp_name]] <- duplicates %>% mutate(experiment = exp_name)
      }
    }, error = function(e) {})
  }
  
  if (length(duplicate_master_list) > 0) {
    # 2. Roll up the granular duplicates into a high-level summary
    all_duplicates <- bind_rows(duplicate_master_list) %>% 
      group_by(experiment, assessment, unit, year, assessment_year) %>%
      summarise(
        Affected_Trees = n(),
        Max_Repeats_Per_Tree = max(measurement_count),
        Sample_Remark = first(Conflicting_Remarks),
        .groups = "drop"
      ) %>%
      select(
        Experiment = experiment, 
        Assessment = assessment, 
        Unit = unit, 
        Calendar_Year = year, 
        Age = assessment_year, 
        Affected_Trees, 
        Max_Repeats = Max_Repeats_Per_Tree,
        Sample_Remark
      ) %>% 
      arrange(Experiment, Assessment, suppressWarnings(as.numeric(Age)))
    
    out_path <- file.path(target_dir, "Global_Duplicate_Scan_Results.csv")
    write_csv(all_duplicates, out_path)
    message("Found duplicates! Summarized report saved to: ", out_path)
  } else {
    message("Scan Complete! No repeat measurements found for any traits.")
  }
}


# ----------------------------------------------------------------------------
# 3. Extract Open Pollinated (OP) Instances
# ----------------------------------------------------------------------------
extract_op_instances <- function(target_dir, species_code) {
  cat("\nScanning for OP instances in Design and Trial data...\n")
  
  # --- PART 1: EXTRACT FROM TRIAL DATA ---
  full_data_files <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")
  
  op_trial_data <- full_data_files %>%
    map_df(function(file) {
      df <- tryCatch(read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())), error = function(e) return(NULL))
      if (is.null(df)) return(NULL)
      
      fam_col <- grep("(?i)^family_name$", names(df), value = TRUE)
      if (length(fam_col) == 0) return(NULL)
      
      df %>% 
        rename(Family_name = all_of(fam_col[1])) %>% 
        select(Family_name) %>% 
        distinct() %>% 
        filter(str_detect(Family_name, "(?i)OP")) %>% 
        mutate(Experiment_Name = str_replace_all(str_extract(basename(file), "^[^_]+"), " ", "_"))
    })
  
  # --- PART 2: EXTRACT FROM DESIGN FILES ---
  # Searches for both .txt and .xlsx design files
  design_files <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)_DF(\\.txt|\\.xlsx|\\.)?$")
  
  op_design_data <- design_files %>%
    map_df(function(file) {
      exp_prefix <- str_replace_all(str_extract(basename(file), "^[^_]+"), " ", "_")
      
      # Leverage the globally available parsers from DP_batch_process_Master.R!
      if (str_detect(file, "(?i)\\.xlsx$")) {
        df <- tryCatch(parse_xlsx_design_file(file, exp_prefix, species_code), error = function(e) return(NULL))
      } else {
        df <- tryCatch(parse_long_design_file(file, exp_prefix, species_code), error = function(e) return(NULL))
      }
      
      if (is.null(df) || !"Family_name" %in% names(df)) return(NULL)
      
      df %>% 
        select(Family_name) %>% 
        distinct() %>% 
        filter(str_detect(Family_name, "(?i)OP")) %>% 
        mutate(Experiment_Name = exp_prefix)
    })
  
  # --- PART 3: THE COMPARISON REPORT ---
  trial_summary <- if(!is.null(op_trial_data) && nrow(op_trial_data) > 0) op_trial_data %>% mutate(In_Trial_Data = "Yes") else tibble(Experiment_Name=character(), Family_name=character(), In_Trial_Data=character())
  design_summary <- if(!is.null(op_design_data) && nrow(op_design_data) > 0) op_design_data %>% mutate(In_Design_Data = "Yes") else tibble(Experiment_Name=character(), Family_name=character(), In_Design_Data=character())
  
  if (nrow(trial_summary) == 0 && nrow(design_summary) == 0) {
    message("No OP instances found in either Design or Trial files.")
    return(invisible(NULL))
  }
  
  # Full Join to catch matches, missing trials, and unexpected additions
  comparison_df <- full_join(design_summary, trial_summary, by = c("Experiment_Name", "Family_name")) %>%
    mutate(
      In_Design_Data = replace_na(In_Design_Data, "No"),
      In_Trial_Data = replace_na(In_Trial_Data, "No"),
      Status = case_when(
        In_Design_Data == "Yes" & In_Trial_Data == "Yes" ~ "Match (Processed Properly)",
        In_Design_Data == "Yes" & In_Trial_Data == "No"  ~ "WARNING: In Design, Missing in Trial",
        In_Design_Data == "No"  & In_Trial_Data == "Yes" ~ "FLAG: Found in Trial, Missing in Design"
      )
    ) %>%
    arrange(Experiment_Name, Family_name)
  
  out_path <- file.path(target_dir, "Diagnostic_OP_Design_vs_Trial_Comparison.csv")
  write_csv(comparison_df, out_path)
  message("OP Comparison complete! Exported to: ", basename(out_path))
  
  # Print a quick summary to the console so the user sees anomalies immediately
  print(comparison_df %>% count(Status, name = "Total_Instances"))
}

# ----------------------------------------------------------------------------
# 4. Format Matrix Files
# ----------------------------------------------------------------------------
format_matrix_files <- function(root_dir) {
  all_dirs <- list.dirs(path = root_dir, recursive = FALSE, full.names = FALSE)
  experiments_to_check <- setdiff(all_dirs, c("00_Scripts", "Archive", ".git", ".Rproj.user"))
  
  for (curr_exp in experiments_to_check) {
    exp_path <- file.path(root_dir, curr_exp)
    source_file <- dir_ls(exp_path, regexp = "(?i)_matrix\\.xlsx|layout\\.xlsx)$")
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