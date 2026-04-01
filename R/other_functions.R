library(tidyverse)
library(readxl)
library(fs)
library(janitor)
library(here)
library(purrr)
library(stringr)


## Summarize ASCII files: ####
data_dir <- "//forestresearch.gov.uk/shares/CSFCC/Forest Resource and Product Assessment and Improvement/NRS-Tree Improvement/CONIFERS/SITKA SPRUCE/psi_DATAPLAN_prep/second powerbi download/"

file_list <- dir_ls(data_dir, recurse = TRUE, regexp = "ASCII\\.xlsx$")

trial_analysis <- file_list %>%
  map_df(function(file) {
    
    # 1. Load data
    df <- read_xlsx(file) %>% clean_names()
    
    # 2. Identify required columns
    # We assume 'value' or similar holds the measurement. 
    # If the measurement is in a column named after the assessment, 
    # we need to pivot the data first.
    
    required <- c("assessment", "assessment_year", "plot", "inferred_tree_position")
    
    if (!all(required %in% names(df))) return(NULL)
    
    df %>%
      # 3. Filter for valid entries (remove rows where the tree position is missing)
      filter(!is.na(inferred_tree_position)) %>%
      
      # 4. Count unique stems per Plot for every Trait + Year combo
      group_by(assessment, assessment_year, plot) %>%
      summarise(
        stems_measured = n_distinct(inferred_tree_position), 
        .groups = "drop"
      ) %>%
      
      # 5. Calculate the Average Stems per Plot across the whole Experiment/File
      group_by(assessment, assessment_year) %>%
      summarise(
        avg_stems_per_plot = mean(stems_measured, na.rm = TRUE),
        min_stems_in_a_plot = min(stems_measured),
        max_stems_in_a_plot = max(stems_measured),
        total_plots_count = n(),
        .groups = "drop"
      ) %>%
      
      # 6. Label with the filename (Experiment ID)
      mutate(experiment_file = basename(file))
  })

# 7. Final Output Organization
final_report <- trial_analysis %>%
  select(experiment_file, assessment, assessment_year, avg_stems_per_plot, total_plots_count) %>%
  arrange(experiment_file, assessment_year)

print(final_report, n = 100)
write.csv(final_report, here("ASCII_initial_report.csv"))



## Counting families/dams & sires: ####

## count unique and shared parents
# 1. Define your two directories
cycle_1_dir <- here("High GCA Fullsib P85-P87 experiments") 
cycle_2_dir <- here("Backwards Selected Fullsib P96-P99 experiments")

# 2. Create the function
count_parents_from_family <- function(dir_path, cycle_name) {
  
  file_list <- dir_ls(dir_path, recurse = TRUE, regexp = "Full_Data_With_Flags\\.csv$")
  
  # Extract and split the family names
  all_parents <- file_list %>%
    map_df(function(file) {
      
      df <- tryCatch(
        read_csv(file, show_col_types = FALSE) %>% clean_names(),
        error = function(e) return(NULL) 
      )
      
      # Check if the file loaded and has the 'family_name' column
      if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
      
      df %>%
        # Grab just the family name column
        select(family_name) %>%
        # Filter down to unique families early on to save processing time
        distinct() %>% 
        # Split the string at the underscore into two new columns
        separate(
          col = family_name, 
          into = c("dam", "sire"), 
          sep = "_", 
          fill = "right",   # If there's no underscore, it assumes the tree is the dam and makes the sire NA
          extra = "merge"   # If there are two underscores, it keeps the extra bits attached to the sire
        ) %>%
        mutate(across(everything(), as.character)) 
    })
  
  # 3. Calculate the unique statistics
  unique_dams <- n_distinct(all_parents$dam, na.rm = TRUE)
  unique_sires <- n_distinct(all_parents$sire, na.rm = TRUE)
  
  # Total unique trees used as a parent (union of both Dam and Sire)
  total_unique_trees <- n_distinct(c(all_parents$dam, all_parents$sire), na.rm = TRUE)
  
  # 4. Return as a single row summary
  tibble(
    cycle = cycle_name,
    unique_dams = unique_dams,
    unique_sires = unique_sires,
    total_unique_parent_trees = total_unique_trees
  )
}

# 5. Run the function and ASSIGN the results
cycle_1_summary <- count_parents_from_family(cycle_1_dir, "P80 experiments")
cycle_2_summary <- count_parents_from_family(cycle_2_dir, "P90 experiments")

# 6. Now combine them
final_comparison <- bind_rows(cycle_1_summary, cycle_2_summary)

##count unique and shared families ###

# 2. Function to just grab a clean list of unique family names from a directory
get_unique_families <- function(dir_path) {
  
  # 1. Use case-insensitive regex to be safe
  file_list <- dir_ls(dir_path, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")
  
  # Check if we even found files before proceeding
  if (length(file_list) == 0) {
    warning(paste("No files found in:", dir_path))
    return(character(0)) 
  }
  
  extracted_data <- file_list %>%
    map_df(function(file) {
      df <- tryCatch(
        read_csv(file, show_col_types = FALSE) %>% janitor::clean_names(),
        error = function(e) return(NULL) 
      )
      
      # Check if the column exists (clean_names makes it lowercase)
      if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
      
      df %>% select(family_name) %>% distinct()
    })
  
  # 2. Check if the resulting dataframe is empty before pulling
  if (nrow(extracted_data) == 0) {
    return(character(0))
  }
  
  extracted_data %>%
    pull(family_name) %>%
    unique() %>%
    na.omit()
}

# 3. Get the lists of families for both cycles
families_c1 <- get_unique_families(cycle_1_dir)
families_c2 <- get_unique_families(cycle_2_dir)

# 4. Perform set operations to count overlap and uniqueness
shared_families   <- length(intersect(families_c1, families_c2))
unique_to_cycle_1 <- length(setdiff(families_c1, families_c2))
unique_to_cycle_2 <- length(setdiff(families_c2, families_c1))

# 5. Build a neat summary table
family_comparison <- tibble(
  metric = c(
    "Total Families in Cycle 1", 
    "Total Families in Cycle 2",
    "Shared (In Both Cycles)", 
    "Unique to Cycle 1 Only", 
    "Unique to Cycle 2 Only"
  ),
  count = c(
    length(families_c1),
    length(families_c2),
    shared_families,
    unique_to_cycle_1,
    unique_to_cycle_2
  )
)

print(family_comparison)


### get all unique family names ####
cycle_1_unique_families <- dir_ls(cycle_1_dir, recurse = TRUE, regexp = "Full_Data_With_Flags\\.csv$") %>%
  map_df(function(file) {
    
    # Safely read the CSV and clean column names
    df <- tryCatch(
      read_csv(file, show_col_types = FALSE) %>% clean_names(),
      error = function(e) return(NULL) 
    )
    
    # Check if the dataframe loaded and actually contains the family_name column
    if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
    
    # Grab just the family names
    df %>% 
      select(family_name) %>% 
      distinct()
  }) %>%
  # Pull the column into a vector, find unique values, and remove NAs
  pull(family_name) %>%
  unique() %>%
  na.omit()

# 3. View the final list
print(cycle_1_unique_families)

write_csv(all_unique_parents, here("All_Unique_Parents_List.csv"))



# 1. Define your directories
cycle_1_dir <- here("High GCA Fullsib P85-P87 experiments") 
cycle_2_dir <- here("Backwards Selected Fullsib P96-P99 experiments")

# 2. Create the robust extraction function
get_parsed_families <- function(dir_path) {
  
  # Extract raw family names
  raw_families <- dir_ls(dir_path, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$") %>%
    map_df(function(file) {
      df <- tryCatch(
        read_csv(file, show_col_types = FALSE) %>% clean_names(),
        error = function(e) return(NULL) 
      )
      if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
      
      df %>% select(family_name) %>% distinct()
    }) %>%
    pull(family_name) %>%
    unique() %>%
    na.omit() %>%
    as_tibble() %>%
    rename(family_name = value)
  
  # Parse and categorize
  parsed_df <- raw_families %>%
    mutate(
      family_type = case_when(
        str_detect(family_name, "(?i)Filler") ~ "Control/Filler",
        str_detect(family_name, "(?i)OP") ~ "Open Pollinated",
        !str_detect(family_name, "_") ~ "Other/Single ID", 
        TRUE ~ "Controlled Cross"
      ),
      dam = case_when(
        family_type %in% c("Control/Filler", "Other/Single ID") ~ family_name, 
        TRUE ~ str_extract(family_name, "^[^_]+") 
      ),
      sire = case_when(
        family_type %in% c("Control/Filler", "Other/Single ID") ~ NA_character_, 
        TRUE ~ str_extract(family_name, "(?<=_).*") 
      )
    )
  
  return(parsed_df)
}

# 3. Run the function on both directories
c1_data <- get_parsed_families(cycle_1_dir)
c2_data <- get_parsed_families(cycle_2_dir)


# 4. FIND EXCLUSIVE FAMILIES


# Families in Cycle 1 but NOT in Cycle 2
families_unique_to_c1 <- anti_join(c1_data, c2_data, by = "family_name")

# Families in Cycle 2 but NOT in Cycle 1
families_unique_to_c2 <- anti_join(c2_data, c1_data, by = "family_name")


# 5. FIND EXCLUSIVE INDIVIDUAL PARENTS


# First, extract just the unique vectors of parents for both
c1_parents <- na.omit(unique(c(c1_data$dam, c1_data$sire)))
c2_parents <- na.omit(unique(c(c2_data$dam, c2_data$sire)))

# Parents in Cycle 1 but NOT in Cycle 2
parents_unique_to_c1 <- setdiff(c1_parents, c2_parents) %>% sort()

# Parents in Cycle 2 but NOT in Cycle 1
parents_unique_to_c2 <- setdiff(c2_parents, c1_parents) %>% sort()


# Print some summaries to the console
cat("\n--- SUMMARY ---\n")
cat("Exclusive Families in Cycle 1:", nrow(families_unique_to_c1), "\n")
cat("Exclusive Families in Cycle 2:", nrow(families_unique_to_c2), "\n")
cat("Exclusive Parents in Cycle 1:", length(parents_unique_to_c1), "\n")
cat("Exclusive Parents in Cycle 2:", length(parents_unique_to_c2), "\n")



### Pull "GEN" info for Cycle 1/2 and compare: ####
## NB: "GEN" is renamed "Origin" in the output files, and is NOT referring to the ORIGIN column in clones_tibdb
# 1. Load Founders once
founders <- read_csv(here("Pedigree", "SS_tibdb_clones.csv"), show_col_types = FALSE) %>%
  mutate(Genotype_name = paste0("ss", number))

# 2. Create the generalized function
get_genotype_origins <- function(target_dir, founders_df) {
  
  # Extract unique families from the target directory
  families <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$") %>%
    map_df(function(file) {
      df <- tryCatch(
        read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())) %>% clean_names(), 
        error = function(e) return(NULL)
      )
      if (is.null(df) || !"Family_name" %in% names(df)) return(NULL)
      
      df %>% select(Family_name) %>% distinct()
    }) %>%
    drop_na() %>%
    pull(Family_name) %>%
    unique()
  
  # Parse out the unique parent IDs (Mums and Dads)
  parents <- tibble(Family_name = families) %>%
    filter(str_detect(Family_name, "_"), !str_detect(Family_name, "(?i)iller")) %>%
    mutate(
      Mum = str_extract(Family_name, "^[^_]+"),
      Raw_Dad = str_extract(Family_name, "(?<=_).*"),
      Dad = if_else(str_detect(Raw_Dad, "(?i)OP"), NA_character_, Raw_Dad)
    )
  
  # Create a single, distinct list of parent genotypes
  unique_genotypes <- unique(na.omit(c(parents$Mum, parents$Dad)))
  
  # Join with founders to check the GEN column
  genotype_gen_check <- tibble(Genotype_name = unique_genotypes) %>%
    left_join(founders_df %>% select(Genotype_name, GEN), by = "Genotype_name")
  
  return(genotype_gen_check)
}

# 3. Run the function for both directories
c1_genotypes <- get_genotype_origins(here("High GCA Fullsib P85-P87 experiments"), founders)
c2_genotypes <- get_genotype_origins(here("Backwards Selected Fullsib P96-P99 experiments"), founders)

# 4. Create summaries for both
c1_summary <- c1_genotypes %>% count(GEN, name = "Cycle_1_Count")
c2_summary <- c2_genotypes %>% count(GEN, name = "Cycle_2_Count")

# 5. Join them together for a side-by-side comparison
comparison_summary <- full_join(c1_summary, c2_summary, by = "GEN") %>%
  # Replace NA counts with 0 for cleaner reading
  mutate(across(c(Cycle_1_Count, Cycle_2_Count), ~replace_na(.x, 0))) %>%
  arrange(GEN)

# Print the final comparison
cat("\n--- Side-by-Side Comparison of 'GEN' values ---\n")
print(comparison_summary)



#### Pull all instances of Open-pollination in Wide_data_With_Flags and the Design files: ####
library(tidyverse)
library(fs)
library(here)
library(readxl)

# 1. Define your target directory 
target_dir <- here("High GCA Fullsib P85-P87 experiments")

# PART 1: EXTRACT FROM FULL DATA FILES
cat("\nScanning Full Data files...\n")
# FIX: Now scanning Full_Data_With_Flags.csv!
full_data_files <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")

if (length(full_data_files) == 0) cat("WARNING: No Full Data files found!\n")

op_trial_data <- full_data_files %>%
  map_df(function(file) {
    df <- tryCatch(
      read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())), 
      error = function(e) return(NULL)
    )
    if (is.null(df)) return(NULL)
    
    # Safely find the family_name column
    fam_col <- grep("(?i)^family_name$", names(df), value = TRUE)
    if (length(fam_col) == 0) return(NULL)
    
    # Extract the experiment name
    exp_name <- str_replace_all(str_extract(basename(file), "^[^_]+"), " ", "_")
    
    df %>%
      rename(Family_name = all_of(fam_col[1])) %>%
      select(Family_name) %>%
      distinct() %>%
      # The Hunt: Find anything with "OP" (this will catch OPST, OPCB, ssOP, etc.)
      filter(str_detect(Family_name, "(?i)OP")) %>%
      mutate(
        Experiment_Name = exp_name,
        Source_File = basename(file),
        Data_Type = "Trial Data"
      )
  })

# PART 2: EXTRACT FROM DESIGN FILES

cat("Scanning Design files...\n")
design_files <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)design.*\\.(csv|xlsx)$")

if (length(design_files) == 0) cat("WARNING: No Design files found!\n")

op_design_data <- design_files %>%
  map_df(function(file) {
    ext <- str_to_lower(path_ext(file))
    df <- tryCatch({
      if (ext == "csv") {
        read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character()))
      } else {
        read_excel(file, col_types = "text")
      }
    }, error = function(e) return(NULL))
    
    if (is.null(df)) return(NULL)
    
    # Safely find the seedlot column
    seed_col <- grep("(?i)seedlot", names(df), value = TRUE)
    if (length(seed_col) == 0) return(NULL)
    
    exp_name <- str_replace_all(str_extract(basename(file), "^[^_]+"), " ", "_")
    
    df %>%
      rename(Seedlot_Name = all_of(seed_col[1])) %>%
      select(Seedlot_Name) %>%
      distinct() %>%
      filter(str_detect(Seedlot_Name, "(?i)OP")) %>%
      mutate(
        Experiment_Name = exp_name,
        Source_File = basename(file),
        Data_Type = "Design Data"
      )
  })


# PART 3: SAFELY EXPORT FOR COMPARISON


cat("\n--- Full Data OP Families (Trial Data) ---\n")
if (!is.null(op_trial_data) && nrow(op_trial_data) > 0 && "Experiment_Name" %in% names(op_trial_data)) {
  print(op_trial_data %>% arrange(Experiment_Name, Family_name), n = 100)
  write_csv(op_trial_data, here("Diagnostic_TrialData_OP.csv"))
} else {
  cat("No 'OP' families found in Full Data files (or files missing).\n")
}

cat("\n--- Design Data OP Seedlots ---\n")
if (!is.null(op_design_data) && nrow(op_design_data) > 0 && "Experiment_Name" %in% names(op_design_data)) {
  print(op_design_data %>% arrange(Experiment_Name, Seedlot_Name), n = 100)
  write_csv(op_design_data, here("Diagnostic_DesignData_OP.csv"))
} else {
  cat("No 'OP' seedlots found in Design Data files (or files missing).\n")
}


#### create pedigree import files ####
library(dplyr)
library(readr)
library(stringr)
library(purrr)
library(fs)
library(tidyr)



# 1. LOAD FOUNDERS AND CONTROLS 

founders <- read_csv(here("Pedigree","SS_tibdb_clones.csv"), show_col_types = FALSE) %>%
  mutate(Genotype_name = paste0("ss", number))

controls <- read_csv(here("Pedigree","dataplan_family_control_import.csv"), show_col_types = FALSE)

#Load the OP families
op_families <- read_excel(here("Pedigree", "Cycle 1 OP.xlsx")) 

# 2. CREATE A PEDIGREE GENERATOR FUNCTION
build_pedigree <- function(target_dir, founders, controls, op_families) {
  is_high_gca <- str_detect(target_dir, "(?i)High GCA")
  
  # Clean the Excel input to remove hidden trailing spaces
  op_families <- op_families %>% mutate(across(where(is.character), str_trim))
  
  # 1. EXTRACT CURATED TRIAL DATA (Changed to _DP_ready.csv)
  trial_files <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)_DP_ready\\.csv$")
  
  trial_data <- trial_files %>%
    map_df(function(file) {
      df <- tryCatch(read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())), error = function(e) NULL)
      if (is.null(df)) return(NULL)
      fam_col <- grep("(?i)^family_name$", names(df), value = TRUE)
      if (length(fam_col) == 0) return(NULL)
      
      trial_name <- str_replace_all(str_extract(basename(file), "^[^_]+"), " ", "_")
      df %>% rename(Family_name = all_of(fam_col[1])) %>% select(Family_name) %>% distinct() %>% 
        mutate(
          Family_name = str_trim(Family_name),
          # Ensure any old OPST labels match the new OPCB notation
          Family_name = str_replace(Family_name, "(?i)_OPST$", "_OPCB"),
          Trial_id = trial_name
        )
    }) %>% na.omit()
  
  # 2. MASTER PARENT METADATA
  raw_mums <- str_extract(trial_data$Family_name, "^[^_]+")
  raw_dads <- str_extract(trial_data$Family_name, "(?<=_).*")
  all_parents <- unique(c(raw_mums, raw_dads, op_families$Mum_name))
  
  parent_meta <- tibble(Genotype_name = all_parents) %>%
    drop_na() %>%
    left_join(founders, by = "Genotype_name") %>%
    mutate(Origin = case_when(GEN == "QC" ~ "HG", GEN == "WC" ~ "WC", TRUE ~ "Unk"))
  
  # 3. BUILD GROUPS 
  grp_locat <- parent_meta %>% filter(!is.na(LOCAT)) %>% select(LOCAT, Origin) %>% distinct() %>%
    mutate(Group_name = paste0(LOCAT, "_", Origin, "++"), Species = "CBCSitka", Type = "UKLR++", Description = paste("Selection group from", LOCAT, "Origin", Origin))
  
  # Clone Banks
  grp_cb <- op_families %>% filter(!is.na(LOCAT)) %>% select(LOCAT) %>% distinct() %>%
    mutate(Group_name = paste0(LOCAT, "+"), Species = "CBCSitka", Type = "Clone Bank+", Description = paste("Open pollinated clone bank at", LOCAT))
  
  groups_final <- bind_rows(
    grp_locat, grp_cb, 
    controls %>% select(Group_name = Family_name, Type, Description = Fam_description),
    tibble(Group_name = "Unknown", Species = "CBCSitka", Type = "Unknown", Description = "Dummy")
  ) %>% distinct(Group_name, .keep_all = TRUE) %>% mutate(Group_id = row_number())
  
  # 4. PROCESS OP FAMILIES (Directly from Excel!)
  fam_op <- op_families %>%
    mutate(
      Stage = 4,
      Dad_id = NA_character_ # Blank, so Step 8 will auto-assign temporary IDs
    ) %>%
    select(Family_name, Mum_name, Mum_type, Dad_name, Dad_type, Fam_description, Stage, Dad_id)
  
  # 5. PROCESS CP FAMILIES
  fam_cp <- trial_data %>%
    filter(str_detect(Family_name, "_"), !str_detect(Family_name, "(?i)iller")) %>%
    filter(!tolower(Family_name) %in% tolower(fam_op$Family_name)) %>%
    mutate(
      Mum_name = str_extract(Family_name, "^[^_]+"), Mum_type = "I",
      Raw_Dad = str_extract(Family_name, "(?<=_).*"),
      Is_OPCB = str_detect(Raw_Dad, "(?i)OPCB"),
      Dad_type = if_else(Is_OPCB, "G", "I"),
      Dad_name = case_when(Is_OPCB ~ "Unknown", TRUE ~ Raw_Dad),
      Dad_id = NA_character_,
      Stage = 4,
      Fam_description = case_when(
        Is_OPCB ~ paste("Open pollinated family from", Mum_name, "in unknown clone bank"),
        is_high_gca ~ "High GCA parents control pollinated",
        TRUE ~ paste("Control pollinated family", Mum_name, "x", Dad_name)
      )
    ) %>%
    select(Family_name, Mum_name, Mum_type, Dad_name, Dad_type, Fam_description, Stage, Dad_id) %>%
    distinct()
  
  # 6. ASSEMBLE ALL FAMILIES
  fam_founders <- grp_locat %>% mutate(Family_name = paste0(Group_name, "_Founders"), Mum_name = Group_name, Mum_type = "G", Dad_name = Group_name, Dad_type = "G", Fam_description = paste("Dummy family for founders in", Group_name), Stage = 2, Dad_id = NA_character_) %>% select(Family_name, Mum_name, Mum_type, Dad_name, Dad_type, Fam_description, Stage, Dad_id)
  fam_fillers <- tibble(Trial_id = unique(trial_data$Trial_id)) %>% mutate(Family_name = paste0(Trial_id, "_Filler"), Mum_name = "Unknown", Mum_type = "G", Dad_name = "Unknown", Dad_type = "G", Fam_description = paste("Fillers for trial", Trial_id), Stage = 3, Dad_id = NA_character_) %>% select(names(fam_founders))
  
  families_combined <- bind_rows(fam_fillers, fam_founders, fam_op, fam_cp) %>% distinct(Family_name, .keep_all = TRUE) %>% arrange(Stage, Family_name) %>% mutate(Family_id = row_number())
  
  # 7. BUILD GENOTYPES 
  genotypes_final <- parent_meta %>%
    filter(!is.na(LOCAT)) %>% 
    mutate(Family_name = paste0(LOCAT, "_", Origin, "++_Founders"), Geno_description = paste("Backward selected founder", Genotype_name, "in", LOCAT), Ortet_locat = LOCAT, Ortet_pyr = PYR, Ortet_origin = Origin, Ortet_lat = if("lat" %in% names(.)) lat else NA, Ortet_long = if("long" %in% names(.)) long else NA, Ortet_ngr_status = if("Status" %in% names(.)) Status else NA, Ortet_prec = if("prec" %in% names(.)) prec else NA, Ortet_tavg = if("tavg" %in% names(.)) tavg else NA, Ortet_elev = if("Elevation" %in% names(.)) Elevation else NA) %>%
    left_join(families_combined %>% select(Family_name, Family_id, Mum_name, Mum_type, Dad_name, Dad_type), by = "Family_name") %>%
    mutate(Genotype_id = row_number()) %>% 
    select(Genotype_id, Genotype_name, Family_name, Family_id, Mum_name, Mum_type, Dad_name, Dad_type, Geno_description, Ortet_locat, Ortet_pyr, Ortet_origin, Ortet_lat, Ortet_long, Ortet_ngr_status, Ortet_prec, Ortet_tavg, Ortet_elev)
  
  # 8. FINAL ID LINKAGE
  families_final <- families_combined %>%
    mutate(
      Mum_id = case_when(
        Mum_type == "G" ~ as.character(groups_final$Group_id[match(Mum_name, groups_final$Group_name)]),
        Mum_type == "I" ~ as.character(genotypes_final$Genotype_id[match(Mum_name, genotypes_final$Genotype_name)]),
        TRUE ~ NA_character_ # Fallback to prevent closure errors if neither condition is met
      ),
      Dad_id = case_when(
        !is.na(Dad_id) ~ as.character(Dad_id), # Keep existing temporary IDs if they exist
        Dad_type == "G" ~ as.character(groups_final$Group_id[match(Dad_name, groups_final$Group_name)]), 
        Dad_type == "I" ~ as.character(genotypes_final$Genotype_id[match(Dad_name, genotypes_final$Genotype_name)]),
        TRUE ~ NA_character_ # Fallback
      )
    ) %>%
    mutate(
      Mum_id = replace_na(Mum_id, "Unknown"), 
      Dad_id = replace_na(Dad_id, "Unknown")
    ) %>%
    select(Family_id, Family_name, Mum_name, Mum_id, Mum_type, Dad_name, Dad_id, Dad_type, Fam_description, Stage)
  
  return(list(groups = groups_final, genotypes = genotypes_final, families = families_final))
}

# 3. RUN FOR BOTH CYCLES & SUBTRACT CYCLE 2
# Generate complete tables for both
c1_tables <- build_pedigree("High GCA Fullsib P85-P87 experiments", founders, controls, op_families)
c2_tables <- build_pedigree("Backwards Selected Fullsib P96-P99 experiments", founders, controls, op_families)

# Filter Cycle 1 to strictly keep things NOT found in Cycle 2
groups_export <- anti_join(c1_tables$groups, c2_tables$groups, by = "Group_name")
genotypes_export <- anti_join(c1_tables$genotypes, c2_tables$genotypes, by = "Genotype_name")
families_export <- anti_join(c1_tables$families, c2_tables$families, by = "Family_name")

# 4. EXPORT
write_csv(c1_tables$groups, here("Pedigree","Cycle1_All_Groups.csv"))
write_csv(c1_tables$genotypes, here("Pedigree","Cycle1_All_Genotypes.csv"))


write_csv(groups_export, here("Pedigree","Cycle1_Unique_Groups_Import.csv"))
write_csv(genotypes_export, here("Pedigree","Cycle1_Unique_Genotypes_Import.csv"))
write_csv(families_export, here("Pedigree","Cycle1_Unique_Families_Import.csv"))

#### debug ;
# 1. Grab all the raw parent strings directly from the Cycle 1 trial names
c1_trial_files <- dir_ls("High GCA Fullsib P85-P87 experiments", recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")
c1_raw_families <- c1_trial_files %>% 
  map_df(~ tryCatch(read_csv(.x, show_col_types = FALSE) %>% select(matches("(?i)^family_name$")), error = function(e) NULL)) %>% 
  pull(1) %>% unique()

# 2. Extract every single Mum and Dad mentioned in those raw strings
c1_mums <- str_extract(c1_raw_families, "^[^_]+")
c1_dads <- str_extract(c1_raw_families, "(?<=_).*")
# Clean out the OP tags so we only have real genotype names
c1_dads_clean <- c1_dads[!str_detect(c1_dads, "(?i)OP") & !is.na(c1_dads)]

c1_all_raw_parents <- unique(c(c1_mums, c1_dads_clean, op_families$Mum_name)) %>% na.omit()

# 3. Who is in the raw data, but didn't make it to the final c1_tables?
dropped_from_c1 <- setdiff(c1_all_raw_parents, c1_tables$genotypes$Genotype_name)

# 4. Who made it to c1_tables, but was deleted by the anti_join?
deleted_by_c2 <- intersect(c1_tables$genotypes$Genotype_name, c2_tables$genotypes$Genotype_name)

cat("\n--- GENOTYPES DROPPED BEFORE JOINING (Likely missing LOCAT or missing from Founders) ---\n")
# Let's see what the founders database actually says about these dropped trees
founders %>% 
  filter(Genotype_name %in% dropped_from_c1) %>% 
  select(Genotype_name, LOCAT, GEN) %>% 
  print(n = 50)

cat("\n--- GENOTYPES DELETED BY ANTI-JOIN (They exist in Cycle 2) ---\n")
print(deleted_by_c2)


##### test antijoin 2 ;; 

library(tidyverse)
library(here)

# 1. Load your actual database exports
# 1. Load your actual database exports correctly
db_fams <- read_excel(here("Backwards Selected Fullsib P96-P99 experiments", "DMS_all_fams.xlsx"))
db_genos <- read_excel(here("Backwards Selected Fullsib P96-P99 experiments", "DMS_all_genotypes.xlsx"))

str(db_fams)
str(db_genos)

db_fam_col <- grep("(?i)family.*name", names(db_fams), value = TRUE)[1]
# Find the genotype name column
db_geno_col <- grep("(?i)name", names(db_genos), value = TRUE)[1]

db_geno_list <- db_genos[[db_geno_col]]
db_fam_list <- db_fams[[db_fam_col]]


# TEST 1: THE ss1551 THEORY
cat("\n--- THE ss1551 MYSTERY ---\n")
cat("Is ss1551 in the online Database?        ", "ss1551" %in% db_geno_list, "\n")
cat("Is ss1551 in the local Cycle 2 files?    ", "ss1551" %in% c2_tables$genotypes$Genotype_name, "\n")
cat("Is ss1551 in the local Cycle 1 files?    ", "ss1551" %in% c1_tables$genotypes$Genotype_name, "\n")

# THE TRUE ANTI-JOIN & UPLOAD BUILDER
cat("\n--- BUILDING TRUE UPLOAD FILES ---\n")

# 1. Establish the TRUE Families to upload (Cycle 1 unique, not in DB)
true_families_export <- families_export %>%
  filter(!Family_name %in% db_fam_list)

# 2. Extract all Individual ('I') parents absolutely needed by these specific families
needed_mums <- true_families_export %>% filter(Mum_type == "I") %>% pull(Mum_name)
needed_dads <- true_families_export %>% filter(Dad_type == "I") %>% pull(Dad_name)
required_parents <- unique(c(needed_mums, needed_dads))

# 3. Build the TRUE Genotypes list based on what the families ACTUALLY need
true_genotypes_export <- c1_tables$genotypes %>%
  filter(Genotype_name %in% required_parents | Genotype_name %in% genotypes_export$Genotype_name) %>%
  filter(!Genotype_name %in% db_geno_list) %>%
  distinct(Genotype_name, .keep_all = TRUE) # Safety catch for duplicates

cat("Verified Cycle 1 Genotypes to upload:", nrow(true_genotypes_export), "\n")
cat("Verified Cycle 1 Families to upload: ", nrow(true_families_export), "\n")

# THE ORPHAN CHECK
# Re-build the clean list strictly from the freshly generated true_families_export
upload_parents_clean <- unique(c(needed_mums, needed_dads))

orphans <- upload_parents_clean[!(upload_parents_clean %in% db_geno_list | upload_parents_clean %in% true_genotypes_export$Genotype_name)]

cat("\n--- ORPHAN CHECK ---\n")
if(length(orphans) == 0) {
  cat("SUCCESS! All Individual parents are accounted for. Safe to upload.\n")
} else {
  cat("WARNING: The following Individual parents are missing from BOTH the database and the genotype upload list:\n")
  print(orphans)
}

# Export the true, database-verified files
write_csv(true_genotypes_export, here("Pedigree","Verified_Cycle1_Genotypes_Import.csv"))
write_csv(true_families_export, here("Pedigree","Verified_Cycle1_Families_Import.csv"))



#### plot pedigree ####
library(dplyr)
library(readr)
library(igraph)
library(ggraph)
library(here)

# 1. LOAD ALL PEDIGREE FILES
families  <- read_csv(here("Pedigree", "Verified_Cycle1_Families_Import.csv"), show_col_types = FALSE)
genotypes <- read_csv(here("Pedigree", "Verified_Cycle1_Genotypes_Import.csv"), show_col_types = FALSE)
groups    <- read_csv(here("Pedigree", "Cycle1_Unique_Groups_Import.csv"), show_col_types = FALSE)


# 2. BUILD THE EDGES (The connecting lines)

# A. Level 3 to 4: Genotype -> CP Family
crosses <- families %>%
  filter(Stage == 4, Mum_name != "Unknown", Dad_name != "Unknown")

edges_mum <- crosses %>% select(from = Mum_name, to = Family_name)
edges_dad <- crosses %>% select(from = Dad_name, to = Family_name)

# B. Level 2 to 3: Group -> Genotype
# We rebuild the Group_name (e.g., "LOCAT_Origin++") from the genotype's attributes
edges_group_geno <- genotypes %>%
  mutate(Group_name = paste0(Ortet_locat, "_", Ortet_origin, "++")) %>%
  select(from = Group_name, to = Genotype_name) %>%
  distinct()

# C. Level 1 to 2: Origin -> Group
# Filter out "Unknown" or blank origins to keep the very top of the plot clean
edges_origin_group <- groups %>%
  filter(Type == "UKLR++", Origin != "Unknown", Origin != "") %>%
  select(from = Origin, to = Group_name) %>%
  distinct()

# Combine all edges into one master list
pedigree_edges <- bind_rows(
  edges_origin_group,
  edges_group_geno,
  edges_mum,
  edges_dad
) %>% 
  filter(!is.na(from), !is.na(to)) # Drop any stray NAs


# 3. BUILD THE NODES (For color-coding)


# Extract a unique list of every single dot that will be on the plot
all_nodes <- data.frame(name = unique(c(pedigree_edges$from, pedigree_edges$to)))

# Assign a category to each dot based on which list it came from
all_nodes <- all_nodes %>%
  mutate(
    Node_Type = case_when(
      name %in% edges_origin_group$from ~ "1. Origin",
      name %in% edges_group_geno$from ~ "2. Group",
      name %in% edges_mum$from | name %in% edges_dad$from ~ "3. Genotype (Parent)",
      TRUE ~ "4. Family (Offspring)"
    )
  )


# 4. CREATE THE GRAPH AND PLOT

# Build the graph object using both the edges and our categorized nodes
pedigree_graph <- graph_from_data_frame(d = pedigree_edges, vertices = all_nodes, directed = TRUE)

# Plot it out
ggraph(pedigree_graph, layout = 'sugiyama') + 
  geom_edge_diagonal(arrow = arrow(length = unit(1.5, 'mm')), 
                     end_cap = circle(3, 'mm'), 
                     alpha = 0.3, 
                     color = "gray50") +
  # Color the dots based on our Node_Type category
  geom_node_point(aes(color = Node_Type), size = 3) +
  geom_node_text(aes(label = name), 
                 vjust = 1.5, 
                 hjust = 0.5, 
                 size = 2.5, 
                 repel = TRUE) +
  theme_void() +
  theme(legend.position = "bottom") +
  scale_color_manual(
    values = c("1. Origin" = "#E41A1C", 
               "2. Group" = "#377EB8", 
               "3. Genotype (Parent)" = "#4DAF4A", 
               "4. Family (Offspring)" = "#984EA3"),
    name = "Pedigree Level"
  ) +
  labs(
    title = "Complete Cycle 1 Pedigree Network",
    subtitle = "Tracking lineages from Regional Origin down to CP Families"
  )



#### Plot ENTIRE Peidgree: ####

# 1. BIND ALL TABLES FROM BOTH CYCLES
# Instead of loading the verified CSVs, we bind the raw, complete 
# tables from both cycles directly from your R environment.

all_families <- bind_rows(c1_tables$families, c2_tables$families) %>% 
  distinct(Family_name, .keep_all = TRUE)

all_genotypes <- bind_rows(c1_tables$genotypes, c2_tables$genotypes) %>% 
  distinct(Genotype_name, .keep_all = TRUE)

all_groups <- bind_rows(c1_tables$groups, c2_tables$groups) %>% 
  distinct(Group_name, .keep_all = TRUE)

# 2. BUILD THE EDGES (The connecting lines)

# A. Level 3 to 4: Genotype -> CP Family
# (Updated to use all_families)
crosses <- all_families %>%
  filter(Stage == 4, Mum_name != "Unknown", Dad_name != "Unknown")

edges_mum <- crosses %>% select(from = Mum_name, to = Family_name)
edges_dad <- crosses %>% select(from = Dad_name, to = Family_name)

# B. Level 2 to 3: Group -> Genotype
# (Updated to use all_genotypes)
edges_group_geno <- all_genotypes %>%
  mutate(Group_name = paste0(Ortet_locat, "_", Ortet_origin, "++")) %>%
  select(from = Group_name, to = Genotype_name) %>%
  distinct()

# C. Level 1 to 2: Origin -> Group
# (Updated to use all_groups)
edges_origin_group <- all_groups %>%
  filter(Type == "UKLR++", !is.na(Origin), Origin != "Unknown", Origin != "") %>%
  select(from = Origin, to = Group_name) %>%
  distinct()

# Combine all edges into one master list
pedigree_edges <- bind_rows(
  edges_origin_group,
  edges_group_geno,
  edges_mum,
  edges_dad
) %>% 
  filter(!is.na(from), !is.na(to)) # Drop any stray NAs


# Extract a unique list of every single dot that will be on the plot
all_nodes <- data.frame(name = unique(c(pedigree_edges$from, pedigree_edges$to)))

# Assign a category to each dot based on which list it came from
all_nodes <- all_nodes %>%
  mutate(
    Node_Type = case_when(
      name %in% edges_origin_group$from ~ "1. Origin",
      name %in% edges_group_geno$from ~ "2. Group",
      name %in% edges_mum$from | name %in% edges_dad$from ~ "3. Genotype (Parent)",
      TRUE ~ "4. Family (Offspring)"
    )
  )


# 4. CREATE THE GRAPH AND PLOT

# Build the graph object using both the edges and our categorized nodes
pedigree_graph <- graph_from_data_frame(d = pedigree_edges, vertices = all_nodes, directed = TRUE)

# Plot it out
ggraph(pedigree_graph, layout = 'sugiyama') + 
  geom_edge_diagonal(arrow = arrow(length = unit(1.5, 'mm')), 
                     end_cap = circle(3, 'mm'), 
                     alpha = 0.3, 
                     color = "gray50") +
  # Color the dots based on our Node_Type category
  geom_node_point(aes(color = Node_Type), size = 3) +
  geom_node_text(aes(label = name), 
                 vjust = 1.5, 
                 hjust = 0.5, 
                 size = 2.5, 
                 repel = TRUE) +
  theme_void() +
  theme(legend.position = "bottom") +
  scale_color_manual(
    values = c("1. Origin" = "#E41A1C", 
               "2. Group" = "#377EB8", 
               "3. Genotype (Parent)" = "#4DAF4A", 
               "4. Family (Offspring)" = "#984EA3"),
    name = "Pedigree Level"
  ) +
  labs(
    title = "Complete SS Pedigree Network",
    subtitle = "Tracking lineages from Regional Origin down to CP Families"
  )



#### What's going on with Radnor 55 Cr_07? ####
## There's an overabundance of values ~200: 
## It's 199 at n=149
rd55<-read.csv(here("High GCA Fullsib P85-P87 experiments","Radnor 55","Radnor 55_Full_Data_With_Flags.csv"))

# Find the most frequent exact values in Cr_07
rd55 %>%
  count(Cr_07) %>%
  arrange(desc(n)) %>%
  head(10)

## test data and residuals of linear model for normalacy: 

# Load necessary libraries
library(ggplot2)
# install.packages("lme4") # Uncomment if you need to install it
library(lme4) 

# PART A: Test normality of raw Cr_07 data

# 1. Visual Check: Q-Q Plot
# If the data is normal, the points will hug the diagonal line tightly.
# The 199 spike will look like a horizontal flat line on this plot.
qqnorm(rd55$Cr_07, main = "Q-Q Plot: Raw Cr_07 Data")
qqline(rd55$Cr_07, col = "red", lwd = 2)

# 2. Statistical Check: Shapiro-Wilk Test
# Note: shapiro.test() fails if you have > 5000 rows. 
shapiro_raw <- shapiro.test(rd55$Cr_07)
print("Shapiro-Wilk Test for Raw Cr_07:")
print(shapiro_raw)

# PART B: Test normality of Model Residuals

# 1. Build the Linear Mixed-Effects Model
# Family_name is fixed; Prow and Ppos are random intercepts.
mod1 <- lmer(Cr_07 ~  + Family_name + (1 | Prow) + (1 | Ppos), data = rd55)

# 2. Extract the residuals
model_resids <- resid(mod1)

# 3. Visual Check: Q-Q Plot of Residuals
qqnorm(model_resids, main = "Q-Q Plot: Model Residuals")
qqline(model_resids, col = "blue", lwd = 2)

# 4. Statistical Check: Shapiro-Wilk on Residuals
shapiro_resids <- shapiro.test(model_resids)

print("Shapiro-Wilk Test for Model Residuals:")
print(shapiro_resids)
