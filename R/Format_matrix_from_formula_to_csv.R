#### extract matrix values from formulatic xlsx files:
#### Library ####
library(here)
library(tidyverse)
library(readxl)
library(writexl)
library(fs)

b59<-read.csv(here("Brecon 59", "Brecon 59_Full_Data_With_Flags.csv"))
# Filter the dataframe and pull the column
b59 %>%
  filter(is.na(Av_20) | is.na(Dm_20)) %>%
  pull(Plot)

# ==============================================================================
# CONFIGURATION
# ==============================================================================

# Directories to skip
skip_dirs <- c("00_Scripts", "Archive", ".git", ".Rproj.user")

# Get list of trial folders
all_dirs <- list.dirs(path = here(), recursive = FALSE, full.names = FALSE)
experiments_to_check <- setdiff(all_dirs, skip_dirs)

message(paste("Found", length(experiments_to_check), "folders to check for matrix files."))


# ==============================================================================
# THE EXTRACTION LOOP
# ==============================================================================

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

message("\nMatrix extraction complete!")