# Dataplan Processing and Pedigree Pipeline

This repository contains the R scripts used for processing field trial data, standardizing traits, performing quality control, and generating import files for Dataplan and pedigree databases.

## ⚠️ Important Setup Instructions (Read First)

This project uses a **hybrid file architecture**:
1. **The Code:** Lives locally on your `C:` drive (managed via this GitHub repository).
2. **The Data:** Lives externally on the shared network drive (`Z:`).

**Do NOT copy raw data (`.csv`, `.xlsx`, `.txt`) into this local repository.** To run these scripts, you must define the `z_drive` path variable at the top of the scripts to point to your local mapping of the shared network drive. 
Example: `z_drive <- "Z:/Shares/CSFCC/.../psi_DATAPLAN_prep"`

## Core Scripts

### 1. Data Processing & Validation
* **`DP_batch_process.R` & `DP_batch_process Multi Tree plots.R`**: The core engines for trial data. These scripts loop through trial folders, parse ASCII measurements and design files, and apply standard naming conventions based on `PPGTraits_UK.xlsm`. They flag statistical outliers, identify tree shrinkage over time, calculate survival, and generate final `.csv` datasets alongside visual PDF reports (histograms, spatial maps, shrinkage plots). They also generate the Dataplan XML import files.
* **`read_BrSt_sheets.R`**: A specialized parser for extracting Branch and Straightness (Br/St) assessment data from highly formatted Kintyre Excel layout sheets into tidy formats.
* **`Format_matrix_from_formula_to_csv.R`**: A utility script to strip formulas from spatial matrix Excel files and save them as clean, machine-readable CSVs.

### 2. Pedigree & Database Generation
* **`other_functions.R`**: A comprehensive suite of tools for database management. Includes functions to:
  * Count unique and shared parents/families across different breeding cycles (e.g., P80 vs P90 experiments).
  * Isolate Open-Pollinated (OP) families from trial and design data.
  * Build complex, verified pedigree import files (`Groups`, `Genotypes`, `Families`) by cross-referencing field trial data against the founder database.
  * Perform diagnostic checks (like mapping origin locations and generating network plots of the pedigree).

## Quality Control (QC) Logic Applied
The batch processors automatically flag data based on the following rules:
* **Missing Data:** Formatted `0`s in continuous traits (like Pilodyn) are converted to `NA`.
* **Outliers:** Standard continuous values outside 4 standard deviations from the mean are flagged.
* **Shrinkage:** Trees that shrink in size compared to previous age assessments are flagged.
* **Resurrection:** Trees marked dead (0) in earlier years but alive (1) in later years are flagged for review.
