

data_dir<-"C:/Users/james.baker/Forest Research/TW CBC-TBA-NextGenBritishConifers - Share/Demo1/Scots_Pine"
  
# Set to FALSE if this is the first tranche and there is no DB to filter against.
HAS_EXISTING_DB <- FALSE

#"C:/Users/james.baker/Forest Research/TW CBC-TBA-NextGenBritishConifers - Share/Demo1/Sitka/High GCA Fullsib P85-P87 experiments"

#High GCA Fullsib P85-P87 experiments
#Backwards Selected Fullsib P96-P99 experiments
#"C:/Users/james.baker/Forest Research/TW CBC-TBA-NextGenBritishConifers - Share/Demo1/Scots_Pine/Trials"

library(tidyverse)
library(readxl)
library(fs)
library(janitor)
library(here)
library(purrr)
library(stringr)


## Summarize ASCII files: ####
# 1. Broaden the regex to capture both .xlsx and .csv ASCII files
file_list <- dir_ls(data_dir, recurse = TRUE, regexp = "(?i)ASCII\\.(xlsx|csv)$")
message("Found ", length(file_list), " ASCII files to summarize.")

trial_analysis <- file_list %>%
  map_df(function(file) {
    
    # 2. Dynamically load the data based on file extension
    if (str_detect(file, "(?i)\\.csv$")) {
      df <- read_csv(file, col_types = cols(.default = "c")) %>% clean_names()
    } else {
      df <- read_xlsx(file, col_types = "text") %>% clean_names()
    }
    
    # --- 3. DYNAMIC COLUMN RENAMING (Copied from Master Pipeline) ---
    
    # Assessment Trait
    assess_col <- intersect(c("assessment", "assessment_type_long", "assessment_type"), names(df))[1]
    if (!is.na(assess_col)) df <- df %>% rename(assessment = !!sym(assess_col))
    
    # Assessment Year
    year_col <- intersect(c("assessment_year", "assessment_year_long", "age"), names(df))[1]
    if (!is.na(year_col)) df <- df %>% rename(assessment_year = !!sym(year_col))
    
    # Plot
    if("plot" %in% names(df)) df <- df %>% rename(plot = plot)
    
    # ----------------------------------------------------------------
    
    # Check for the absolute bare minimum columns required for a summary
    required_base <- c("assessment", "assessment_year", "plot")
    if (!all(required_base %in% names(df))) {
      message("Skipping ", basename(file), " - missing core columns (assessment, year, plot)")
      return(NULL)
    }
    
    # --- 4. DYNAMIC TREE POSITION HANDLING ---
    if ("inferred_tree_position" %in% names(df)) {
      # Standard Modern Trial: Use existing column
      df <- df %>% rename(tree_pos = inferred_tree_position)
    } else {
      # Fallback for 90s Single-Tree OR Scots Pine Multi-Tree without IDs
      # We group by Plot + Trait + Year and just count the rows!
      df <- df %>% 
        group_by(assessment, assessment_year, plot) %>% 
        mutate(tree_pos = row_number()) %>% 
        ungroup()
    }
    
    df %>%
      # 5. Filter for valid entries
      filter(!is.na(tree_pos)) %>%
      
      # 6. Count unique stems per Plot for every Trait + Year combo
      group_by(assessment, assessment_year, plot) %>%
      summarise(
        stems_measured = n_distinct(tree_pos), 
        .groups = "drop"
      ) %>%
      
      # 7. Calculate the Average Stems per Plot across the whole Experiment/File
      group_by(assessment, assessment_year) %>%
      summarise(
        avg_stems_per_plot = mean(stems_measured, na.rm = TRUE),
        min_stems_in_a_plot = min(stems_measured),
        max_stems_in_a_plot = max(stems_measured),
        total_plots_count = n(),
        .groups = "drop"
      ) %>%
      
      # 8. Label with the filename (Experiment ID)
      mutate(experiment_file = basename(file))
  })

# Output Organization
final_report <- trial_analysis %>%
  select(experiment_file, assessment, assessment_year, avg_stems_per_plot, min_stems_in_a_plot, max_stems_in_a_plot, total_plots_count) %>%
  arrange(experiment_file, assessment_year)

print(final_report, n = 100)
write_csv(final_report, file.path(data_dir, "MASTER_ASCII_Inventory_Report.csv"))
message("Report saved successfully!")


## 2: Count instances where AV was measured more than once per stem #### 
## (see ASCII remarks; some instances where it was taken on both NE and SW side)

message(paste("Scanning", length(file_list), "ASCII files for duplicate Av measurements..."))

# Initialize an empty list to store results
duplicate_av_list <- list()

for (file_path in file_list) {
  # Extract experiment name from the folder path for our report
  exp_name <- basename(dirname(file_path))
  
  tryCatch({
    # 1. Dynamically load the data based on file extension (Mirroring Part 1)
    if (str_detect(file_path, "(?i)\\.csv$")) {
      raw_data <- read_csv(file_path, col_types = cols(.default = "c")) %>% clean_names()
    } else {
      raw_data <- read_excel(file_path, col_types = "text") %>% clean_names()
    }
    
    # 2. Dynamic Column Renaming (Mirroring Part 1)
    assess_col <- intersect(c("assessment", "assessment_type_long", "assessment_type"), names(raw_data))[1]
    if (!is.na(assess_col)) raw_data <- raw_data %>% rename(assessment = !!sym(assess_col))
    
    year_col <- intersect(c("assessment_year", "assessment_year_long", "age"), names(raw_data))[1]
    if (!is.na(year_col)) raw_data <- raw_data %>% rename(assessment_year = !!sym(year_col))
    
    if("plot" %in% names(raw_data)) raw_data <- raw_data %>% rename(plot = plot)
    
    if ("inferred_tree_position" %in% names(raw_data)) {
      raw_data <- raw_data %>% rename(tree_pos = inferred_tree_position)
    } else {
      raw_data <- raw_data %>% 
        group_by(assessment, assessment_year, plot) %>% 
        mutate(tree_pos = row_number()) %>% 
        ungroup()
    }
    
    # 3. Ensure the required columns exist after standardizing names
    req_cols <- c("plot", "tree_pos", "assessment", "assessment_year")
    if (all(req_cols %in% names(raw_data))) {
      
      # Group by tree and age, and count the occurrences
      duplicates <- raw_data %>%
        filter(str_detect(assessment, "(?i)^AV")) %>% # Isolate Av measurements
        group_by(plot, tree_pos, assessment, assessment_year) %>%
        summarise(measurement_count = n(), .groups = "drop") %>%
        filter(measurement_count > 1) # Keep only the ones with multiple readings
      
      if (nrow(duplicates) > 0) {
        duplicates <- duplicates %>% mutate(experiment = exp_name)
        duplicate_av_list[[exp_name]] <- duplicates
      }
    }
  }, error = function(e) {
    message(paste("  -> Skipped or Error reading", exp_name, ":", e$message))
  })
}

# Compile and print the results
if (length(duplicate_av_list) > 0) {
  all_duplicates <- bind_rows(duplicate_av_list) %>%
    select(Experiment = experiment, Plot = plot, Tree = tree_pos, Assessment = assessment, Age = assessment_year, Measurement_Count = measurement_count) %>%
    arrange(Experiment, as.numeric(Plot), as.numeric(Tree))
  
  message("\n==========================================")
  message("Scan Complete! Found repeat Av measurements in the following trials:")
  print(unique(all_duplicates$Experiment))
  
  message("\nHere is a preview of the duplicates:")
  print(head(all_duplicates, 15))
  
  # Export the full report to a CSV using the unified data_dir variable
  out_path <- file.path(data_dir, "Repeat_Av_Scan_Results.csv")
  write_csv(all_duplicates, out_path)
  message(paste("\nFull diagnostic report saved to:", out_path))
  
} else {
  message("\nScan Complete! No repeat Av measurements found in any of the checked files.")
}

## Counting families/dams & sires: ####
## Count unique and shared parents
# 1. Define your two directories using generalized terms (previously Cycle 1 and 2)
pending_dir <- file.path(data_dir, "Trials") 
existing_dir <- file.path(data_dir, "Backwards Selected Fullsib P96-P99 experiments")

# 2. Create the function
count_parents_from_family <- function(dir_path, batch_name) {
  
  file_list <- dir_ls(dir_path, recurse = TRUE, regexp = "Full_Data_With_Flags\\.csv$")
  
  # Extract and split the family names
  all_parents <- file_list %>%
    map_df(function(file) {
      
      df <- tryCatch(
        read_csv(file, show_col_types = FALSE) %>% clean_names(),
        error = function(e) return(NULL) 
      )
      
      if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
      
      df %>%
        select(family_name) %>%
        distinct() %>% 
        separate(
          col = family_name, 
          into = c("dam", "sire"), 
          sep = "_", 
          fill = "right",   
          extra = "merge"   
        ) %>%
        mutate(across(everything(), as.character)) 
    })
  
  # 3. Calculate the unique statistics
  unique_dams <- n_distinct(all_parents$dam, na.rm = TRUE)
  unique_sires <- n_distinct(all_parents$sire, na.rm = TRUE)
  total_unique_trees <- n_distinct(c(all_parents$dam, all_parents$sire), na.rm = TRUE)
  
  # 4. Return as a single row summary
  tibble(
    batch = batch_name,
    unique_dams = unique_dams,
    unique_sires = unique_sires,
    total_unique_parent_trees = total_unique_trees
  )
}

# 5. Run the function and ASSIGN the results
pending_summary <- count_parents_from_family(pending_dir, "Pending Upload")

if (HAS_EXISTING_DB) {
  existing_summary <- count_parents_from_family(existing_dir, "Existing Database")
} else {
  # Generate a zeroed-out placeholder to keep the output table tidy
  existing_summary <- tibble(
    batch = "Existing Database",
    unique_dams = 0,
    unique_sires = 0,
    total_unique_parent_trees = 0
  )
}

# 6. Now combine them
final_comparison <- bind_rows(pending_summary, existing_summary)

## Count unique and shared families ###

# 2. Function to just grab a clean list of unique family names from a directory
get_unique_families <- function(dir_path) {
  
  file_list <- dir_ls(dir_path, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")
  
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
      
      if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
      df %>% select(family_name) %>% distinct()
    })
  
  if (nrow(extracted_data) == 0) {
    return(character(0))
  }
  
  extracted_data %>% pull(family_name) %>% unique() %>% na.omit()
}

# 3. Get the lists of families for both batches
families_pending <- get_unique_families(pending_dir)

if (HAS_EXISTING_DB) {
  families_existing <- get_unique_families(existing_dir)
} else {
  families_existing <- character(0) # Empty vector for seamless setdiff math
}

# 4. Perform set operations to count overlap and uniqueness
shared_families   <- length(intersect(families_pending, families_existing))
unique_to_pending <- length(setdiff(families_pending, families_existing))
unique_to_existing <- length(setdiff(families_existing, families_pending))

# 5. Build a neat summary table
family_comparison <- tibble(
  metric = c(
    "Total Families in Pending Upload", 
    "Total Families in Existing DB",
    "Shared (In Both)", 
    "Unique to Pending Upload Only", 
    "Unique to Existing DB Only"
  ),
  count = c(
    length(families_pending),
    length(families_existing),
    shared_families,
    unique_to_pending,
    unique_to_existing
  )
)

print(family_comparison)


### Get all unique family names from Pending ###
pending_unique_families <- dir_ls(pending_dir, recurse = TRUE, regexp = "Full_Data_With_Flags\\.csv$") %>%
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
  na.omit()

print(pending_unique_families)

# write_csv(all_unique_parents, file.path(data_dir,"All_Unique_Parents_List.csv"))


### Categorize parsed families and compare ###
pending_data <- get_parsed_families(pending_dir)

if (HAS_EXISTING_DB) {
  existing_data <- get_parsed_families(existing_dir)
} else {
  # Empty dataframe to allow anti_join to pass all pending data through
  existing_data <- tibble(family_name = character(), family_type = character(), dam = character(), sire = character())
}

# FIND EXCLUSIVE FAMILIES
families_exclusive_to_pending <- anti_join(pending_data, existing_data, by = "family_name")
families_exclusive_to_existing <- anti_join(existing_data, pending_data, by = "family_name")

# FIND EXCLUSIVE INDIVIDUAL PARENTS
pending_parents <- na.omit(unique(c(pending_data$dam, pending_data$sire)))
existing_parents <- na.omit(unique(c(existing_data$dam, existing_data$sire)))

parents_exclusive_to_pending <- setdiff(pending_parents, existing_parents) %>% sort()
parents_exclusive_to_existing <- setdiff(existing_parents, pending_parents) %>% sort()


cat("\n--- SUMMARY ---\n")
cat("Exclusive Families in Pending Upload:", nrow(families_exclusive_to_pending), "\n")
cat("Exclusive Families in Existing DB:", nrow(families_exclusive_to_existing), "\n")
cat("Exclusive Parents in Pending Upload:", length(parents_exclusive_to_pending), "\n")
cat("Exclusive Parents in Existing DB:", length(parents_exclusive_to_existing), "\n")

### Pull "GEN" info for Pending/Existing and compare: ####

# Define species prefix for founder genotypes (e.g., "sp" or "ss")
species_prefix <- "sp" 

# 1. Load Founders once (Ensure file name matches your species)
founders <- read_csv(file.path(data_dir,"Pedigree", "SP_tibdb_clones.csv"), show_col_types = FALSE) %>%
  mutate(Genotype_name = paste0(species_prefix, number))

# 2. Create the generalized function
get_genotype_origins <- function(target_dir, founders_df) {
  families <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$") %>%
    map_df(function(file) {
      df <- tryCatch(
        # Force all columns to character to prevent read_csv parsing warnings
        read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())) %>% clean_names(), 
        error = function(e) return(NULL)
      )
      
      # Update: Check for lowercase "family_name" generated by clean_names()
      if (is.null(df) || !"family_name" %in% names(df)) return(NULL)
      
      df %>% select(family_name) %>% distinct()
    }) %>%
    drop_na() %>% pull(family_name) %>% unique()
  
  parents <- tibble(family_name = families) %>%
    filter(str_detect(family_name, "_"), !str_detect(family_name, "(?i)iller")) %>%
    mutate(
      Mum = str_extract(family_name, "^[^_]+"),
      Raw_Dad = str_extract(family_name, "(?<=_).*"),
      Dad = if_else(str_detect(Raw_Dad, "(?i)OP"), NA_character_, Raw_Dad)
    )
  
  unique_genotypes <- unique(na.omit(c(parents$Mum, parents$Dad)))
  
  genotype_gen_check <- tibble(Genotype_name = unique_genotypes) %>%
    left_join(founders_df %>% select(Genotype_name, GEN), by = "Genotype_name")
  
  return(genotype_gen_check)
}

pending_genotypes <- get_genotype_origins(pending_dir, founders)
pending_summary <- pending_genotypes %>% count(GEN, name = "Pending_Count")

if (HAS_EXISTING_DB) {
  existing_genotypes <- get_genotype_origins(existing_dir, founders)
  existing_summary <- existing_genotypes %>% count(GEN, name = "Existing_Count")
} else {
  existing_summary <- tibble(GEN = character(), Existing_Count = numeric())
}

comparison_summary <- full_join(pending_summary, existing_summary, by = "GEN") %>%
  mutate(across(c(Pending_Count, Existing_Count), ~replace_na(.x, 0))) %>%
  arrange(GEN)

cat("\n--- Side-by-Side Comparison of 'GEN' values ---\n")
print(comparison_summary)

#### Extract all instances of Open-pollination in Data and Design files: ####

# PART 1: EXTRACT FROM FULL DATA FILES
cat("\nScanning Full Data files...\n")
full_data_files <- dir_ls(pending_dir, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")

if (length(full_data_files) == 0) cat("WARNING: No Full Data files found!\n")

op_trial_data <- full_data_files %>%
  map_df(function(file) {
    df <- tryCatch(
      read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())), 
      error = function(e) return(NULL)
    )
    if (is.null(df)) return(NULL)
    
    fam_col <- grep("(?i)^family_name$", names(df), value = TRUE)
    if (length(fam_col) == 0) return(NULL)
    
    exp_name <- str_replace_all(str_extract(basename(file), "^[^_]+"), " ", "_")
    
    df %>%
      rename(Family_name = all_of(fam_col[1])) %>%
      select(Family_name) %>%
      distinct() %>%
      filter(str_detect(Family_name, "(?i)OP")) %>%
      mutate(
        Experiment_Name = exp_name,
        Source_File = basename(file),
        Data_Type = "Trial Data"
      )
  })

# PART 2: EXTRACT FROM DESIGN FILES
cat("Scanning Design files...\n")
design_files <- dir_ls(pending_dir, recurse = TRUE, regexp = "(?i)design.*\\.(csv|xlsx)$")

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
  write_csv(op_trial_data, file.path(data_dir,"Diagnostic_TrialData_OP.csv"))
} else {
  cat("No 'OP' families found in Full Data files (or files missing).\n")
}

cat("\n--- Design Data OP Seedlots ---\n")
if (!is.null(op_design_data) && nrow(op_design_data) > 0 && "Experiment_Name" %in% names(op_design_data)) {
  print(op_design_data %>% arrange(Experiment_Name, Seedlot_Name), n = 100)
  write_csv(op_design_data, file.path(data_dir,"Diagnostic_DesignData_OP.csv"))
} else {
  cat("No 'OP' seedlots found in Design Data files (or files missing).\n")
}

#### What's going on with Radnor 55 Cr_07? ####
## There's an overabundance of values ~200: 
## It's 199 at n=149
rd55<-read.csv(file.path(data_dir,"High GCA Fullsib P85-P87 experiments","Radnor 55","Radnor 55_Full_Data_With_Flags.csv"))

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


#### Loop through project directories and format formulaic cell contents into R-legible values.  ####
# CONFIGURATION
# Directories to skip
skip_dirs <- c("00_Scripts", "Archive", ".git", ".Rproj.user")

# Get list of trial folders
all_dirs <- list.dirs(path = here(), recursive = FALSE, full.names = FALSE)
experiments_to_check <- setdiff(all_dirs, skip_dirs)

message(paste("Found", length(experiments_to_check), "folders to check for matrix files."))

# THE EXTRACTION LOOP

for (curr_exp in experiments_to_check) {
  
  exp_path <- here(curr_exp)
  
  # 1. Look for the "no_matrix" file
  # Regex: Case insensitive, ends in _no_matrix.xlsx
  source_file <- dir_ls(exp_path, regexp = "(?i)_matrix\\.xlsx$")
  
  if (length(source_file) == 0) {
    # Optional: print message if missing, or just silent skip
    # message(paste("  [SKIP]", curr_exp, "- No '_no_matrix.xlsx' found."))
    next
  }
  
  message(paste("\nProcessing:", curr_exp))
  message(paste("  -> Found:", basename(source_file)))
  
  # 2. Read the "matrix" sheet
  # We use tryCatch because the sheet name might vary (Matrix, matrix, Sheet1?)
  matrix_data <- tryCatch({
    # Try reading "matrix" sheet specifically
    read_excel(source_file, sheet = "matrix", col_names = FALSE)
  }, error = function(e) {
    message("  -> WARNING: Could not find 'matrix' sheet. Trying first sheet...")
    return(read_excel(source_file, col_names = FALSE))
  })
  
  # 3. Clean the Data
  # Convert to matrix to drop any formulas/formatting weirdness
  # Convert NA to 0
  clean_matrix <- as.matrix(matrix_data)
  clean_matrix[is.na(clean_matrix)] <- 0
  
  # 4. Save as Clean CSV (Preferred for R processing)
  # Naming convention: [Trial]_Matrix.csv
  dest_filename <- paste0(str_replace_all(curr_exp, " ", "_"), "_Matrix.csv")
  dest_path <- file.path(exp_path, dest_filename)
  
  write.table(clean_matrix, dest_path, 
              row.names = FALSE, col.names = FALSE, sep = ",")
  
  message(paste("  -> SAVED:", dest_filename))
}

#### Correct Kielder 162 Ht_05 mirroring issue ####
## This trait is fully mirrored over the plot: compare patterns of dead tress between Ht_05 and Dm_10. 
library(tidyverse)

# 1. Point this directly at Kielder 162
target_csv <- "C:/Users/james.baker/Forest Research/TW CBC-TBA-NextGenBritishConifers - Share/Sitka/Backwards Selected Fullsib P96-P99 experiments/Kielder 162/Kielder_162_Full_Data_With_Flags.csv"

# Load the wide data
df <- read_csv(target_csv, show_col_types = FALSE)

# --- STORE "BEFORE" STATE FOR PLOTTING ---
df_before <- df %>% 
  select(Plot, Tree, Dm_09, Ht_05) %>% 
  mutate(State = "1. Before Fix (Mirrored Error)")

# 2. The critical fix: Use BOTH Plot and Tree as the unique identifiers!
spatial_map <- df %>%
  select(Plot, Tree, Prow, Ppos) %>%
  filter(!is.na(Prow) & !is.na(Ppos)) %>%
  distinct(Prow, Ppos, .keep_all = TRUE) 

# 3. Isolate the exact columns that need to be flipped
traits_to_flip <- intersect(names(df), c("Ht_05", "Sur_05", "Ht_05_reject"))

flip_data <- df %>%
  select(Plot, Tree, Prow, Ppos, all_of(traits_to_flip)) %>%
  filter(!is.na(Prow) & !is.na(Ppos)) %>%
  distinct(Plot, Tree, .keep_all = TRUE)

# 4. The Spatial Math: Find target coordinates
min_x <- min(flip_data$Ppos, na.rm = TRUE)
max_x <- max(flip_data$Ppos, na.rm = TRUE)

# Map the source data to the TARGET Plot/Tree at the mirrored destination
flip_mapping <- flip_data %>%
  mutate(Mirrored_Ppos = max_x - Ppos + min_x) %>%
  left_join(spatial_map %>% select(Target_Plot = Plot, Target_Tree = Tree, Prow, Ppos), 
            by = c("Prow" = "Prow", "Mirrored_Ppos" = "Ppos")) %>%
  filter(!is.na(Target_Plot))

# 5. Extract the data and assign it to the new Target Trees
fixed_traits <- flip_mapping %>%
  select(Plot = Target_Plot, Tree = Target_Tree, all_of(traits_to_flip)) %>%
  distinct(Plot, Tree, .keep_all = TRUE)

# 6. Merge the pristine data back together!
df_clean <- df %>% select(-all_of(traits_to_flip))

df_final <- df_clean %>%
  # Join securely using BOTH Plot and Tree
  left_join(fixed_traits, by = c("Plot", "Tree")) %>% 
  mutate(
    Validation_record = case_when(
      is.na(Validation_record) ~ "Ht_05 Globally Mirrored L-R",
      TRUE ~ paste(Validation_record, "| Ht_05 Globally Mirrored L-R")
    )
  )

# 7. Save the fixed dataset
out_path <- str_replace(target_csv, "(?i)\\.csv$", "_Corrected.csv")
write_csv(df_final, out_path, na = "")

message("Global Mirror applied flawlessly! Saved to: ", basename(out_path))

# 8. VERIFICATION PLOT
df_after <- df_final %>% 
  select(Plot, Tree, Dm_09, Ht_05) %>% 
  mutate(State = "2. After Fix (Corrected)")

plot_data <- bind_rows(df_before, df_after) %>%
  filter(!is.na(Dm_09) & !is.na(Ht_05))

p_verify <- ggplot(plot_data, aes(x = Dm_09, y = Ht_05)) +
  geom_point(alpha = 0.5, size = 1, color = "#2c3e50") +
  geom_smooth(method = "lm", formula = y ~ x, color = "#e74c3c", linetype = "dashed", se = FALSE) +
  facet_wrap(~State) +
  theme_bw() +
  labs(
    title = "Kielder 162: Ht_05 Global Spatial Correction",
    subtitle = "Verifying the Left-to-Right Field Mirror Fix",
    x = "Dm_09 (Trusted Baseline)",
    y = "Ht_05"
  ) +
  theme(
    strip.background = element_rect(fill = "#ecf0f1"),
    strip.text = element_text(face = "bold", size = 11)
  )

print(p_verify)

# 9. BEFORE & AFTER SPATIAL HEATMAP

# 1. Extract the BEFORE spatial data directly from the raw 'df'
spatial_before <- df %>% 
  select(Prow, Ppos, Value = Ht_05) %>% 
  mutate(State = "1. Before Fix (Mirrored Error)",
         Prow = as.numeric(Prow), 
         Ppos = as.numeric(Ppos)) %>%
  filter(!is.na(Value), !is.na(Prow), !is.na(Ppos))

# 2. Extract the AFTER spatial data directly from the corrected 'df_final'
spatial_after <- df_final %>% 
  select(Prow, Ppos, Value = Ht_05) %>% 
  mutate(State = "2. After Fix (Corrected)",
         Prow = as.numeric(Prow), 
         Ppos = as.numeric(Ppos)) %>%
  filter(!is.na(Value), !is.na(Prow), !is.na(Ppos))

# 3. Combine them into a single dataframe for faceting
heatmap_data <- bind_rows(spatial_before, spatial_after)

# 4. Generate the side-by-side Spatial Map using your custom styling
p_heatmap <- ggplot(heatmap_data, aes(x = Ppos, y = Prow, fill = Value)) + 
  geom_tile(color = "white", size = 0.1) + 
  scale_fill_distiller(palette = "Spectral", direction = -1, name = "Ht_05") + 
  scale_y_reverse() + 
  facet_wrap(~State) +
  labs(
    title = "Kielder 162: Ht_05 Spatial Heatmap Comparison",
    subtitle = "Visually verifying the left-to-right mirror correction",
    x = "Field Position (X)", 
    y = "Field Row (Y)"
  ) + 
  theme_minimal() + 
  coord_fixed() +
  theme(
    strip.background = element_rect(fill = "#ecf0f1"),
    strip.text = element_text(face = "bold", size = 11),
    legend.position = "right"
  )

# Display the heatmap in your RStudio Viewer
print(p_heatmap)