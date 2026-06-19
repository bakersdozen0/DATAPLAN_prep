# ============================================================================
# ENGINE: PEDIGREE DIAGNOSTICS & PRE-FLIGHT CHECKS
# ============================================================================
# ----------------------------------------------------------------------------
# 1. DATA EXTRACTORS (Returns standard tibble: family_name, dam, sire)
# ----------------------------------------------------------------------------

get_trial_parents <- function(dir_path) {
  file_list <- dir_ls(dir_path, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")
  if (length(file_list) == 0) return(tibble(family_name = character(), dam = character(), sire = character()))
  
  extracted <- file_list %>%
    map_df(function(file) {
      df <- tryCatch(read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())) %>% clean_names(), error = function(e) return(NULL))
      if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
      df %>% select(family_name) %>% distinct()
    })
  
  if (is.null(extracted) || nrow(extracted) == 0) return(tibble(family_name = character(), dam = character(), sire = character()))
  
  extracted %>% 
    drop_na(family_name) %>% 
    distinct(family_name) %>%
    mutate(
      dam = if_else(str_detect(family_name, "_"), str_extract(family_name, "^[^_]+"), family_name),
      sire = if_else(str_detect(family_name, "_"), str_extract(family_name, "(?<=_).*"), family_name)
    )
}

get_dms_parents <- function(existing_dir) {
  dms_fam_path <- file.path(existing_dir, "DMS_fams.xlsx")
  
  # Try reading as XLSX first, fallback to CSV if the user saved it that way
  if (!file.exists(dms_fam_path)) {
    dms_fam_path <- file.path(existing_dir, "DMS_fams.csv")
    if (!file.exists(dms_fam_path)) {
      message("  -> WARNING: 'DMS_fams' (xlsx or csv) not found in Existing directory.")
      return(tibble(family_name = character(), dam = character(), sire = character()))
    }
    df <- tryCatch(read_csv(dms_fam_path, show_col_types = FALSE) %>% clean_names(), error = function(e) return(NULL))
  } else {
    df <- tryCatch(read_excel(dms_fam_path, col_types = "text") %>% clean_names(), error = function(e) return(NULL))
  }
  
  if (is.null(df) || !"family_name" %in% names(df) || !"mum_name" %in% names(df) || !"dad_name" %in% names(df)) {
    message("  -> WARNING: DMS_fams file is missing required columns (family_name, mum_name, dad_name).")
    return(tibble(family_name = character(), dam = character(), sire = character()))
  }
  
  df %>% 
    select(family_name, dam = mum_name, sire = dad_name) %>% 
    drop_na(family_name) %>% 
    distinct()
}

# ----------------------------------------------------------------------------
# 2. ANALYSIS MATH
# ----------------------------------------------------------------------------

summarize_parents <- function(df, batch_name) {
  if (nrow(df) == 0) return(tibble(batch = batch_name, unique_dams = 0, unique_sires = 0, total_unique_parent_trees = 0))
  
  # Filter out "OP" from sires so we don't count "OP" as a physical tree
  clean_sires <- df$sire[!str_detect(df$sire, "(?i)OP") & !is.na(df$sire)]
  clean_dams <- df$dam[!is.na(df$dam)]
  
  tibble(
    batch = batch_name,
    unique_dams = n_distinct(clean_dams),
    unique_sires = n_distinct(clean_sires),
    total_unique_parent_trees = n_distinct(c(clean_dams, clean_sires))
  )
}

parse_family_types <- function(df) {
  if(nrow(df) == 0) return(tibble(family_name = character(), family_type = character(), dam = character(), sire = character()))
  
  df %>%
    mutate(
      family_type = case_when(
        str_detect(family_name, "(?i)OP[A-Z]{0,2}") ~ "Open Pollinated (OP)",
        str_detect(family_name, "(?i)iller") ~ "Filler",
        dam == sire ~ "Control / Provenance / Founder", 
        TRUE ~ "Control Pollinated (CP)"
      )
    )
}

get_gen_from_families <- function(df, founders_df) {
  if(nrow(df) == 0) return(tibble(Genotype_name = character(), GEN = character()))
  
  unique_genotypes <- unique(na.omit(c(df$dam, df$sire)))
  
  # Remove "OP" so it doesn't look for an "OP" tree in the founders list
  unique_genotypes <- unique_genotypes[!str_detect(unique_genotypes, "(?i)OP")]
  
  tibble(Genotype_name = unique_genotypes) %>% 
    left_join(founders_df %>% select(Genotype_name, GEN), by = "Genotype_name")
}

# ----------------------------------------------------------------------------
# 3. MAIN EXECUTION WRAPPER
# ----------------------------------------------------------------------------
run_pedigree_diagnostics <- function(base_dir, pending_dir, existing_dir, founders_file_path, species_code, has_existing_db) {
  
  cat("\n==========================================")
  cat("\nINITIATING PEDIGREE PRE-FLIGHT CHECKS")
  cat("\n==========================================\n")
  
  # --- 1. EXTRACT DATA TABLES ---
  pending_data <- get_trial_parents(pending_dir)
  existing_data <- if(has_existing_db) get_dms_parents(existing_dir) else tibble(family_name = character(), dam = character(), sire = character())
  
  # --- 2. COUNT PARENTS ---
  pending_summary <- summarize_parents(pending_data, "Pending Upload")
  existing_summary <- summarize_parents(existing_data, "Existing Database")
  
  cat("\n--- Parent Tree Counts ---\n")
  print(bind_rows(pending_summary, existing_summary))
  
  # --- 3. COUNT FAMILIES OVERLAP ---
  families_pending <- unique(pending_data$family_name)
  families_existing <- unique(existing_data$family_name)
  
  family_comparison <- tibble(
    metric = c("Total Families in Pending Upload", "Total Families in Existing DB", "Shared (In Both)", "Unique to Pending Upload Only", "Unique to Existing DB Only"),
    count = c(length(families_pending), length(families_existing), length(intersect(families_pending, families_existing)), length(setdiff(families_pending, families_existing)), length(setdiff(families_existing, families_pending)))
  )
  cat("\n--- Family Name Overlap ---\n")
  print(family_comparison)
  
  # --- 4. CATEGORIZE & COMPARE ---
  pending_categorized <- parse_family_types(pending_data)
  existing_categorized <- parse_family_types(existing_data)
  
  cat("\n--- Exclusive Entity Summary ---\n")
  cat("Exclusive Families in Pending Upload:", nrow(anti_join(pending_categorized, existing_categorized, by = "family_name")), "\n")
  cat("Exclusive Families in Existing DB:", nrow(anti_join(existing_categorized, pending_categorized, by = "family_name")), "\n")
  
  # --- 5. CHECK 'GEN' ORIGINS ---
  founders <- tryCatch(
    read_csv(founders_file_path, show_col_types = FALSE) %>% 
      clean_names() %>% 
      rename(GEN = gen), # <--- FIX: Forces 'gen' back to uppercase immediately
    error = function(e) return(NULL)
  )
  
  if (!is.null(founders) && "number" %in% names(founders)) {
    founders <- founders %>% mutate(Genotype_name = paste0(tolower(species_code), number))
    
    pending_genotypes <- get_gen_from_families(pending_data, founders) %>% count(GEN, name = "Pending_Count")
    existing_genotypes <- get_gen_from_families(existing_data, founders) %>% count(GEN, name = "Existing_Count")
    
    comparison_summary <- full_join(pending_genotypes, existing_genotypes, by = "GEN") %>% 
      mutate(across(c(Pending_Count, Existing_Count), ~replace_na(.x, 0))) %>% 
      arrange(GEN)
    
    cat("\n--- Side-by-Side Comparison of 'GEN' values ---\n")
    print(comparison_summary)
  } else {
    cat("\n--- GEN Comparison Skipped (Founders file missing or invalid) ---\n")
  }
  
  cat("\n>>> PRE-FLIGHT CHECKS COMPLETE <<<\n")
}