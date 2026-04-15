# DATAPLAN Prep: Unified Master Pipeline

This repository contains the R-based data processing pipeline for standardizing, validating, and formatting raw tree breeding data exported from Dataplan. 

It takes raw ASCII exports, merges them with design files and spatial matrices, applies rigorous validation rules (outlier detection, temporal shrinkage checks, spatial coordinate mapping), and outputs clean, wide-format datasets and `.xml` files ready for genetic evaluation.

## 🚀 The Master Script (`DP_batch_process_Master.R`)
This repository has been refactored into a **single unified pipeline** that automatically handles:
* **Any Species:** The script dynamically reads the `species` or `species_code` column from the raw ASCII file to build the correct genetic family names and control prefixes (e.g., `ss_` vs `sp_`).
* **Any File Format:** Seamlessly accepts both `.csv` and `.xlsx` inputs for raw data, and either `.txt` or general format `.`design files.
* **Single-Tree & Multi-Tree Plots:** Governed by a simple toggle at the top of the script.

### ⚙️ Configuration
At the very top of `DP_batch_process_Master.R`, you must set three global variables before running a batch:

    PLOT_TYPE     <- "MULTI" # Set to "SINGLE" or "MULTI" depending on the trial design
    BASE_DIR      <- "C:/Users/.../Sitka" # The root directory of the species
    TRIAL_SERIES  <- "High GCA Fullsib P85-P87 experiments" # The specific subfolder to process

* **`PLOT_TYPE = "SINGLE"`**: Forces all measurements to `Tree = 1`. Bypasses multi-tree specific operations like 8x1 mirrored plot trait reversals.
* **`PLOT_TYPE = "MULTI"`**: Expects or generates sequential `InferredTreePosition` IDs. Maps interior subset trees to full spatial plots and automatically detects/reverses mirrored data entry.

## 📂 Repository Requirements

* **`DP_batch_process_Master.R`**: The main execution script.
* **`Trait_trans.csv`**: The master trait translation map. **This file must remain in the same directory as the R Project.** It maps raw Dataplan headers (e.g., `HEIGHT`) to standardized output variables and contains XML formatting rules.

## 📁 Expected Trial Folder Structure
The script dynamically crawls the `TRIAL_SERIES` folder. For each trial (e.g., `Kintyre_17`), it looks for the following files:

1. **`*_ASCII.csv`** or **`*_ASCII.xlsx`** *(Required)*: The raw Dataplan measurement export.
2. **`*_DF.txt`**, **`*_DF.`** or **`*_DF.xlsx`** *(Optional)*: The design file containing Crosses, Plot, and Block logic.
3. **`*_Matrix.csv`** *(Optional)*: The spatial layout mapping Plot IDs to physical Rows/Positions.
4. **`*_AD_<age>.csv/xlsx`** *(Optional)*: Additional Data files (e.g., Resi data) to be merged. The age must be in the filename (e.g., `Ardross_AD_10.xlsx`).

## 📊 Pipeline Outputs
For each processed trial folder, the script generates:

* `*_Full_Data_With_Flags.csv`: The master wide-format dataset containing all measurements, spatial coordinates, pedigree info, and binary `_reject` flags for invalid data.
* `*_graphs.pdf`: A comprehensive diagnostic report containing histograms, temporal shrinkage scatter plots, correlation matrices, and spatial heatmaps.
* `*_Stats.csv`: A summary table of N, Mean, Min, Max, and CV% for all valid traits.
* `*_Traits.xml`: The standardized XML trait definitions required for downstream evaluation.

## 🛑 Validation Flags (`Reject_Flag = 1`)
Data points are flagged for rejection under the following conditions:
* **Extreme Outliers:** Values exceeding Mean ± 4 Standard Deviations (for continuous traits).
* **Invalid Ordinal Scores:** Visual scores (e.g., 0) that fall outside expected 1-6 ordinal scales. 
* **Temporal Shrinkage:** A physical measurement (e.g., Height, DBH) that is strictly smaller than the measurement at the previous age (requires matching units).
