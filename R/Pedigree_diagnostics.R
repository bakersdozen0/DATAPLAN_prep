# ============================================================================
# ENGINE: PEDIGREE DIAGNOSTICS & PRE-FLIGHT CHECKS
# ============================================================================

# ----------------------------------------------------------------------------
# HELPER FUNCTIONS
# ----------------------------------------------------------------------------
count_parents_from_family <- function(dir_path, batch_name) {
  file_list <- dir_ls(dir_path, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")
  
  all_parents <- file_list %>%
    map_df(function(file) {
      df <- tryCatch(read_csv(file, show_col_types = FALSE) %>% clean_names(), error = function(e) return(NULL))
      if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
      
      df %>% select(family_name) %>% distinct() %>% 
        separate(col = family_name, into = c("dam", "sire"), sep = "_", fill = "right", extra = "merge") %>%
        mutate(across(everything(), as.character)) 
    })
  
  tibble(
    batch = batch_name,
    unique_dams = n_distinct(all_parents$dam, na.rm = TRUE),
    unique_sires = n_distinct(all_parents$sire, na.rm = TRUE),
    total_unique_parent_trees = n_distinct(c(all_parents$dam, all_parents$sire), na.rm = TRUE)
  )
}

get_unique_families <- function(dir_path) {
  file_list <- dir_ls(dir_path, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")
  if (length(file_list) == 0) return(character(0)) 
  
  extracted_data <- file_list %>%
    map_df(function(file) {
      df <- tryCatch(read_csv(file, show_col_types = FALSE) %>% clean_names(), error = function(e) return(NULL))
      if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
      df %>% select(family_name) %>% distinct()
    })
  
  if (nrow(extracted_data) == 0) return(character(0))
  extracted_data %>% pull(family_name) %>% unique() %>% na.omit()
}

get_parsed_families <- function(dir_path) {
  file_list <- dir_ls(dir_path, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")
  if (length(file_list) == 0) return(tibble(family_name = character(), family_type = character(), dam = character(), sire = character()))
  
  file_list %>%
    map_df(function(file) {
      df <- tryCatch(read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())) %>% clean_names(), error = function(e) return(NULL))
      if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
      df %>% select(family_name) %>% distinct()
    }) %>%
    distinct() %>% drop_na(family_name) %>%
    mutate(
      dam = str_extract(family_name, "^[^_]+"),
      sire = str_extract(family_name, "(?<=_).*"),
      family_type = case_when(
        str_detect(family_name, "(?i)OP") ~ "Open Pollinated (OP)",
        str_detect(family_name, "(?i)iller") ~ "Filler",
        is.na(sire) ~ "Control / Provenance", 
        TRUE ~ "Control Pollinated (CP)"
      )
    ) %>% select(family_name, family_type, dam, sire)
}

get_genotype_origins <- function(target_dir, founders_df) {
  families <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$") %>%
    map_df(function(file) {
      df <- tryCatch(read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())) %>% clean_names(), error = function(e) return(NULL))
      if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
      df %>% select(family_name) %>% distinct()
    }) %>% drop_na() %>% pull(family_name) %>% unique()
  
  parents <- tibble(family_name = families) %>%
    filter(str_detect(family_name, "_"), !str_detect(family_name, "(?i)iller")) %>%
    mutate(Mum = str_extract(family_name, "^[^_]+"), Raw_Dad = str_extract(family_name, "(?<=_).*"), Dad = if_else(str_detect(Raw_Dad, "(?i)OP"), NA_character_, Raw_Dad))
  
  unique_genotypes <- unique(na.omit(c(parents$Mum, parents$Dad)))
  tibble(Genotype_name = unique_genotypes) %>% left_join(founders_df %>% select(Genotype_name, GEN), by = "Genotype_name")
}


# ----------------------------------------------------------------------------
# MAIN EXECUTION WRAPPER
# ----------------------------------------------------------------------------
run_pedigree_diagnostics <- function(base_dir, pending_dir, existing_dir, founders_file_path, species_code, has_existing_db) {
  
  cat("\n==========================================")
  cat("\nINITIATING PEDIGREE PRE-FLIGHT CHECKS")
  cat("\n==========================================\n")
  
  # --- 1. COUNT PARENTS ---
  pending_summary <- count_parents_from_family(pending_dir, "Pending Upload")
  if (has_existing_db) {
    existing_summary <- count_parents_from_family(existing_dir, "Existing Database")
  } else {
    existing_summary <- tibble(batch = "Existing Database", unique_dams = 0, unique_sires = 0, total_unique_parent_trees = 0)
  }
  
  cat("\n--- Parent Tree Counts ---\n")
  print(bind_rows(pending_summary, existing_summary))
  
  # --- 2. COUNT FAMILIES ---
  families_pending <- get_unique_families(pending_dir)
  families_existing <- if(has_existing_db) get_unique_families(existing_dir) else character(0)
  
  family_comparison <- tibble(
    metric = c("Total Families in Pending Upload", "Total Families in Existing DB", "Shared (In Both)", "Unique to Pending Upload Only", "Unique to Existing DB Only"),
    count = c(length(families_pending), length(families_existing), length(intersect(families_pending, families_existing)), length(setdiff(families_pending, families_existing)), length(setdiff(families_existing, families_pending)))
  )
  cat("\n--- Family Name Overlap ---\n")
  print(family_comparison)
  
  # --- 3. CATEGORIZE & COMPARE ---
  pending_data <- get_parsed_families(pending_dir)
  existing_data <- if(has_existing_db) get_parsed_families(existing_dir) else tibble(family_name = character(), family_type = character(), dam = character(), sire = character())
  
  cat("\n--- Exclusive Entity Summary ---\n")
  cat("Exclusive Families in Pending Upload:", nrow(anti_join(pending_data, existing_data, by = "family_name")), "\n")
  cat("Exclusive Families in Existing DB:", nrow(anti_join(existing_data, pending_data, by = "family_name")), "\n")
  
  # --- 4. CHECK 'GEN' ORIGINS ---
  # Load founders and dynamically apply the species prefix
  founders <- read_csv(founders_file_path, show_col_types = FALSE) %>%
    mutate(Genotype_name = paste0(tolower(species_code), number))
  
  pending_genotypes <- get_genotype_origins(pending_dir, founders) %>% count(GEN, name = "Pending_Count")
  existing_genotypes <- if(has_existing_db) get_genotype_origins(existing_dir, founders) %>% count(GEN, name = "Existing_Count") else tibble(GEN = character(), Existing_Count = numeric())
  
  comparison_summary <- full_join(pending_genotypes, existing_genotypes, by = "GEN") %>% mutate(across(c(Pending_Count, Existing_Count), ~replace_na(.x, 0))) %>% arrange(GEN)
  
  cat("\n--- Side-by-Side Comparison of 'GEN' values ---\n")
  print(comparison_summary)
  
  # --- 5. EXTRACT OP INSTANCES ---
  cat("\nScanning for OP Families in Trial Data...\n")
  full_data_files <- dir_ls(pending_dir, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")
  
  op_trial_data <- full_data_files %>%
    map_df(function(file) {
      df <- tryCatch(read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())), error = function(e) return(NULL))
      if (is.null(df)) return(NULL)
      fam_col <- grep("(?i)^family_name$", names(df), value = TRUE)
      if (length(fam_col) == 0) return(NULL)
      df %>% rename(Family_name = all_of(fam_col[1])) %>% select(Family_name) %>% distinct() %>% filter(str_detect(Family_name, "(?i)OP")) %>% mutate(Experiment_Name = str_replace_all(str_extract(basename(file), "^[^_]+"), " ", "_"), Source_File = basename(file), Data_Type = "Trial Data")
    })
  
  if (!is.null(op_trial_data) && nrow(op_trial_data) > 0) {
    out_path <- file.path(base_dir, "Diagnostic_TrialData_OP.csv")
    write_csv(op_trial_data, out_path)
    cat("Found", nrow(op_trial_data), "OP instances. Exported to:", basename(out_path), "\n")
  } else {
    cat("No 'OP' families found in Full Data files.\n")
  }
  
  cat("\n>>> PRE-FLIGHT CHECKS COMPLETE <<<\n")
}