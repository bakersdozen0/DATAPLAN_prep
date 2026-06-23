# ============================================================================
# DATAPLAN & DMS PIPELINE: MASTER WHEEL
# ============================================================================
# Instructions: Adjust the Configuration variables below, then run the 
#               execution blocks as needed. Do not edit the Engine files.
# ============================================================================

library(tidyverse)
library(usethis)
library(readxl)
library(fs)
library(here)
library(gridExtra)
library(grid)
library(janitor)
library(purrr)
library(stringr)
library(igraph)
library(ggraph)

# --- LOAD ENGINES ---
source(here("R", "DP_batch_process_Master.R"))
source(here("R", "Data_orientation_corrections.R"))
source(here("R", "Pedigree_diagnostics.R")) 
source(here("R", "Pedigree.R"))
source(here("R", "ASCII_diagnostics.R"))

# ============================================================================
#### USER CONFIGURATION ####
# ============================================================================

# GLOBAL SETTINGS
BASE_DIR      <- "Z:/CSFCC/Forest Resource and Product Assessment and Improvement/NRS-Tree Improvement/CONIFERS/SITKA SPRUCE/psi_DATAPLAN_prep/Latest_Series"
TRIAL_SERIES  <- "SS P91 5 Polycrosses" # "High GCA Fullsib P85-P87 experiments" / "Backwards selected Fullsib P96-P99 experiments"/ "Population_Studies" /"Trials" / "Diallel"
TARGET_DIR    <- file.path(BASE_DIR, TRIAL_SERIES) # for user legibility 
TARGET_TRIALS <- NULL # set to NULL to run on all experiments in trial_series directory, or to a list of trials that you want to test: e.g. c("Craigellachie 49", "Ae 58")

PLOT_TYPE     <- "MULTI" # Refering to number of trees per plot "SINGLE" or "MULTI"
SPECIES_CODE  <- "SS"    # this is FR notation ( "SS" or "SP" )
SPECIES_NAME  <- "CBCSitka" # This is DMS notation ( "CBC Sitka" or "CBCScots" )

# DATAPLAN SPECIFIC
TRAITS_FILE   <- here("Trait_trans.csv")

# TRAVERSAL AUDIT SPECIFIC
TRAVERSAL_FILE      <- "C:/Users/james.baker/Forest Research/TW CBC-TBA-NextGenBritishConifers - Share/Demo1/Sitka/High GCA Fullsib P85-P87 experiments/Craigellachie 49/Craigellachie_49_Full_Data_With_Flags.csv"
BASELINE_TRAIT      <- "Ht_06" 
TEST_TRAITS         <- c("Dm_10","Ht_10","Pil_15","Dm_15","Cr_07") 
GRID_ROWS           <- 8 ## these are only used if a matrix/layout file is missing
GRID_COLS           <- 1
EXPECT_NEGATIVE_COR <- FALSE
USE_CORRECTED_DATA  <- FALSE
FORCE_GLOBAL_OVERRIDE <- FALSE

# PEDIGREE SPECIFIC
HAS_EXISTING_DB <- TRUE 
EXISTING_DIR    <- file.path(BASE_DIR, "Backwards Selected Fullsib P96-P99 experiments")
FOUNDERS_FILE   <- paste0(SPECIES_CODE, "_tibdb_clones.csv") 
CONTROLS_FILE   <- "dataplan_family_control_import.csv" 
OP_FAM_FILE     <- paste0(SPECIES_CODE, "_OP_Families.xlsx")

# ============================================================================
#### EXECUTION BLOCKS ####
# ============================================================================

# UTILITIES & DIAGNOSTICS
# Standalone functions for specific data checks.
summarize_ascii_inventory(target_dir = TARGET_DIR)
scan_duplicates(target_dir = TARGET_DIR)
extract_op_instances(target_dir= TARGET_DIR, species_code = SPECIES_CODE)
format_matrix_files(root_dir = BASE_DIR)

# 1. MAIN DATAPLAN PIPELINE
# Converts raw ASCII/XLSX into formatted wide/long data with flags and XML.
run_dataplan_pipeline(
  base_dir     = BASE_DIR,
  trial_series = TRIAL_SERIES,
  traits_file  = TRAITS_FILE,
  plot_type    = PLOT_TYPE,
  target_trials = TARGET_TRIALS
)

# 2.PLOT TRAVERSAL AUDIT
# Diagnoses spatial orientation and mirroring issues.
run_traversal_audit(
  wide_data_csv       = TRAVERSAL_FILE,
  plot_type           = PLOT_TYPE,
  baseline_trait      = BASELINE_TRAIT,
  test_traits         = TEST_TRAITS,
  grid_rows           = GRID_ROWS,
  grid_cols           = GRID_COLS,
  expect_negative_cor = EXPECT_NEGATIVE_COR,
  use_corrected_data  = USE_CORRECTED_DATA,
  force_global_override = FORCE_GLOBAL_OVERRIDE
)

# 2.5 PEDIGREE PRE-FLIGHT DIAGNOSTICS
# Summarizes unique families/parents, overlap with DB, and OP instances.
run_pedigree_diagnostics(
  base_dir           = BASE_DIR,
  pending_dir        = file.path(BASE_DIR, TRIAL_SERIES),
  existing_dir       = EXISTING_DIR,
  founders_file_path = file.path(BASE_DIR, "Pedigree", FOUNDERS_FILE),
  species_code       = SPECIES_CODE,
  has_existing_db    = HAS_EXISTING_DB
)

# 3. DMS PEDIGREE BUILDER
# Generates cross, family, and group linkages for database import.
run_pedigree_builder(
  base_dir        = BASE_DIR,
  pending_dir     = file.path(BASE_DIR, TRIAL_SERIES),
  existing_dir    = EXISTING_DIR,
  species_code    = SPECIES_CODE,
  species_name    = SPECIES_NAME,
  founders_file   = FOUNDERS_FILE,
  controls_file   = CONTROLS_FILE,
  op_fam_file     = OP_FAM_FILE,
  has_existing_db = HAS_EXISTING_DB
)

