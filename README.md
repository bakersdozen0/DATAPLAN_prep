# DATAPLAN Prep: Unified Master Pipeline

This repository contains the R-based data processing pipeline for standardizing, validating, and formatting raw tree breeding data exported from the CEDD prior to upload to Dataplan. 

It takes raw ASCII exports, merges them with design files and spatial matrices, applies validation rules (outlier detection, temporal shrinkage checks, spatial coordinate mapping), and outputs clean, wide-format datasets, graphical summary and `.xml` files.

## ⚠️ Important Setup Instructions (Read First)

This project uses a **hybrid file architecture**:
1. **The Code:** Lives locally on your `C:` drive (managed via this GitHub repository).
2. **The Data:** Lives externally on the shared network drive (`Z:`).

**Do NOT copy raw data (`.csv`, `.xlsx`, `.txt`) into this local repository.** To run these scripts, you must define the `z_drive` path variable at the top of the scripts to point to your local mapping of the shared network drive. 
Example: `z_drive <- "Z:/Shares/CSFCC/.../psi_DATAPLAN_prep"`


## Core Scripts

## 1. The Master Script (`DP_batch_process_Master.R`)
This repository has been refactored into a **single unified pipeline** that automatically handles:
* **Any Species:** The script dynamically reads the `species` or `species_code` column from the raw ASCII file to build the correct genetic family names and control prefixes (e.g., `ss_` vs `sp_`).
* **Any File Format:** Seamlessly accepts both `.csv` and `.xlsx` inputs for raw data, and either `.txt` or general format `.`design files.
* **Single-Tree & Multi-Tree Plots:** Governed by a simple toggle at the top of the script.

## 2. Pedigree Generation
* **`Pedigree.R`**: Script for pulling family names from Wide format data and generating pedigre import files (`Groups`, `Genotypes`, `Families`) by cross-referencing field trial data against the founder database, and comparing these data to the pedigree data existing in DP.

## 3. Other Helpful Functions
* **`other_functions.R`** . Includes functions to:
  * Count unique and shared parents/families across different breeding cycles (e.g., P80 vs P90 experiments).
  * Isolate Open-Pollinated (OP) families from trial and design data.
  * Perform diagnostic checks (like mapping origin locations and generating network plots of the pedigree).

## 4. Reading field data sheets: 
* **`read_BrSt_sheets.R`**: A parser for extracting Branch and Straightness (Br/St) assessment data from  unformatted Kintyre field layout sheets into tidy formats.

# ⚙️ Configuration
At the very top of `DP_batch_process_Master.R`, you must set three global variables before running a batch:

    PLOT_TYPE     <- "MULTI" # Set to "SINGLE" or "MULTI" depending on the trial design
    BASE_DIR      <- "C:/Users/.../Sitka" # The root directory of the species
    TRIAL_SERIES  <- "High GCA Fullsib P85-P87 experiments" # The specific subfolder to process

* **`PLOT_TYPE = "SINGLE"`**: Forces all measurements to `Tree = 1`. Bypasses multi-tree specific operations like 8x1 mirrored plot trait reversals.
* **`PLOT_TYPE = "MULTI"`**: Expects or generates sequential `InferredTreePosition` IDs. Maps interior subset trees to full spatial plots and automatically detects/reverses mirrored data entry.

### 📂 Repository Requirements

* **`DP_batch_process_Master.R`**: The main execution script.
* **`Trait_trans.csv`**: The master trait translation map. **This file must remain in the same directory as the R Project.** It maps raw CEDD headers (e.g., `HEIGHT`) to standardized output variables and contains XML formatting rules.

### 📁 Expected Trial Folder Structure
The script dynamically crawls the `TRIAL_SERIES` folder. For each trial (e.g., `Kintyre_17`), it looks for the following files:
**NB:** This script will still run without throwing errors even if it only has the ASCII file! It will warn you if a DF or Matrix is missing, however. 

1. **`*_ASCII.csv`** or **`*_ASCII.xlsx`** *(Required)*: The raw Dataplan measurement export.
2. **`*_DF.txt`**, **`*_DF.`** or **`*_DF.xlsx`** *(Optional)*: The design file containing Crosses, Plot, and Block logic.
3. **`*_Matrix.csv`** *(Optional)*: The spatial layout mapping Plot IDs to physical Rows/Positions.
4. **`*_AD_<age>.csv/xlsx`** *(Optional)*: Additional Data files (e.g., Resi data) to be merged. The age must be in the filename (e.g., `Ardross_AD_10.xlsx`).

### 📊 Pipeline Outputs
For each processed trial folder, the script generates:

* `*_Full_Data_With_Flags.csv`: The master wide-format dataset containing all measurements, spatial coordinates, pedigree info, and binary `_reject` flags for invalid data.
* `*_graphs.pdf`: A comprehensive diagnostic report containing histograms, temporal shrinkage scatter plots, correlation matrices, and spatial heatmaps.
* `*_Stats.csv`: A summary table of N, Mean, Min, Max, and CV% for all valid traits.
* `*_Traits.xml`: The standardized XML trait definitions required for downstream evaluation.

### 🛑 Validation Flags (`Reject_Flag = 1`)
Data points are flagged for rejection under the following conditions:
* **Extreme Outliers:** Values exceeding Mean ± 4 Standard Deviations (for continuous traits).
* **Invalid Ordinal Scores:** Visual scores (e.g., 0) that fall outside expected 1-6 ordinal scales. 
* **Temporal Shrinkage:** A physical measurement (e.g., Height, DBH) that is strictly smaller than the measurement at the previous age (requires matching units).
