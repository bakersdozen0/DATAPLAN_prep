# DATAPLAN Prep: Unified Master Pipeline

This repository contains the R-based data processing pipeline for standardizing, validating, and formatting raw tree breeding data exported from the CEDD prior to upload to our Data Management System (DMS). 

It takes raw ASCII exports, merges them with design files and spatial matrices, applies validation rules (outlier detection, temporal shrinkage checks, spatial coordinate mapping), and outputs clean, wide-format datasets, graphical summaries, correlation tables, and `.xml` files.

## ⚠️ Important Setup Instructions (Read First)

This project uses a **hybrid file architecture**:
1. **The Code:** Lives locally on your `C:` drive (managed via this GitHub repository).
2. **The Data:** Lives externally on the shared network drive (`Z:`) or synced Teams/SharePoint folders.

**Do NOT copy raw data (`.csv`, `.xlsx`, `.txt`) into this local repository.** To run these scripts, you must define the `BASE_DIR` and `TRIAL_SERIES` path variables at the top of the scripts to point to your local data location.

---

## Core Scripts

### 1. The Master Script (`DP_batch_process_Master.R`)
The core execution engine, refactored to handle complex corrections and standardizations. 
* **Primary Inputs:** Designed primarily to ingest PowerBI export ASCII (long-format) files produced by RT.
* **Fallback Ingestion:** Automatically sweeps for Additional Data (`_AD_<age>.csv/xlsx`) or raw text files (e.g., `Ht_06.txt`) to capture data missing from the CEDD database or accidentally omitted from the ASCII export.
* **Surgical ID Re-alignment:** Automatically detects the presence of a `TRAVERSAL_HELPER_MASTER.csv`. It "un-scrambles" the Tree IDs for specific plots and traits in-memory, stamping the fix and the anchor used into the final validation record.
* **Global Zero-Padding:** Automatically enforces leading zeros on all assessment ages (e.g., `Ht_3` becomes `Ht_03`). This ensures trait names match flawlessly across all data sources.
* **Hardcoded Edge Cases:** Contains logic for massive trial-level intercepts (e.g., the Kielder 162 global left-to-right mirror hotfix) that bypass standard plot-level logic.

### 2. Traversal Diagnostic & Orientation Tool (`Data_orientation_corrections.R`)
This tool is used *after* an initial run of the Master Script. Discrepancies in expected inter-age or inter-trait correlations (typically visible as notable left-skewing in the `.pdf` correlation plots, such as `Ht_05` vs `Dm_15`) are investigated here. 

This code scans and corrects instances of human error in spatial data entry:
* **16-Way Brute Force:** Simulates all 16 possible physical traversal paths (Starting corners, Horizontal/Vertical, Typewriter/Snake) and finds the one that maximizes biological correlation against a trusted anchor.
* **Inverse Logic (Negative Correlations):** Includes an `EXPECT_NEGATIVE_COR` toggle to hunt for paths that maximize *negative* correlations, which is essential when anchoring density traits (like Pilodyn) against growth traits.
* **Global Consensus Override:** If >60% of plots in a trial suggest the exact same spatial fix, the script overrides plot-level filtering and applies the fix globally to the entire trial.
* **Temporal Chaining & Safeguards:** Supports iterative cleaning (using a corrected trait as a "stepping stone" to fix another). *Note: Chaining is speculative. It is highly effective for square/rectangular multi-tree plots where 16 distinct paths exist, but it relies heavily on visually identifying consistent assessor error patterns and should ideally be anchored to one or more manually verified traits.*
    * **Anti-Double-Scramble:** Includes a critical safety catch. If you attempt to chain a fix on a trait that already has a rule in the helper file, the script will abort to prevent applying a relative path on top of already scrambled data.
* **Output:** Generates and appends to a `TRAVERSAL_HELPER_MASTER.csv` in the parent directory—a "recipe book" of surgical fixes that the Master Script reads automatically when run again.

*(Note: Initially pushed to this repo by JB, much of this code was originally written by MC and adapted for an iterative pipeline).*

### 3. Pedigree Generation (`Pedigree.R`)
* Generates import files (`Groups`, `Genotypes`, `Families`) to facilitate batch uploads ("tranches" of ~15-20 trials at a time) to the Data Management System (DMS). 
* Cross-references local trial data against static downloaded DMS pedigree files (Groups, Families, Controls, Genotypes) representing the data already in the system, effectively preventing duplicate uploads for the new tranche.

### 4. Reading Field Data Sheets (`read_BrSt_sheets.R`)
* A parser for extracting Branch and Straightness (Br/St) assessment data from unformatted, raw field layout sheets (e.g., Kintyre) into tidy formats, complete with automated Chi-Square tests to validate assessor scoring distributions.

### 5. Utilities (`other_functions.R`)
* A collection of ad-hoc scripts used for ground-truthing data. Includes tools for generating master ASCII inventory reports, scanning for duplicate measurements, and tallying family/parent overlaps between breeding cycles.

---

# ⚙️ Workflow: Detect, Diagnose, & Chain

To ensure the highest data integrity in multi-tree trials, the following workflow is recommended:

1. **The Baseline Run (Detect):** Run the Master Script (`DP_batch_process_Master.R`) on your raw data. This produces an initial `_Full_Data_With_Flags.csv` and the diagnostic `_graphs.pdf`. *(Note: Output files will NOT have the `_Corrected` suffix during a standard baseline run).*
2. **Visual Inspection:** Review the correlation plots in the generated PDF. Look for traits with poor correlations or severe left-skewing against expected 1:1 baselines. 
3. **Establish an Anchor:** Identify a stable trait to use as your Absolute Anchor. **Ideally, this trait should be manually verified by comparing the physical paper records (from the drawers) against the CEDD database export.**
4. **The Fix (Diagnose):** Open the Diagnostic Tool (`Data_orientation_corrections.R`). Feed it the wide data CSV from Step 1. Test the suspect trait against your verified anchor. Save the recommendations to the Master Helper.
5. **The Chain (Speculative):** Toggle `USE_CORRECTED_DATA <- TRUE` in the Diagnostic Tool. You can now use your newly corrected trait as an anchor for further diagnostics. Use this primarily when you can observe a consistent pattern of assessor error.
6. **The Final Pass (Correct):** Run the Master Script one final time on the raw data folder. It will seamlessly ingest the `TRAVERSAL_HELPER_MASTER.csv`, apply the fixes simultaneously, and output files cleanly marked with a `_Corrected` suffix to indicate spatial logic was applied.

---

# 📂 Folder Structure & Requirements

### 📁 Expected Trial Folder Input
The script dynamically trawls the `TRIAL_SERIES` folder for:
* **`*_ASCII.csv/xlsx`** *(Required)*: The raw PowerBI measurement export.
* **`*_Matrix.csv`** *(Soft Requirement)*: The physical spatial layout of the trial. Technically optional for raw pipeline processing, but strictly required for downstream spatial analyses and final upload to DMS.
* **`*_DF.xlsx/txt`** *(Optional)*: The design file containing Crosses and Blocks.
* **`*_AD_<age>.csv/xlsx`** *(Optional)*: Additional Data files to be merged. Examples include harvest age Dm and Av from Moray, or resi data from Kintyre, Kielder, Moray and Brecon. 
* **`TRAVERSAL_HELPER_MASTER.csv`** *(Optional)*: The orientation "recipe book" generated by the Diagnostic Tool.

### 📊 Main Pipeline Outputs
*(The `_Corrected` suffix is automatically appended ONLY if the `TRAVERSAL_HELPER_MASTER.csv` is detected and applied, or a hardcoded edge-case is triggered).*

* `*_Full_Data_With_Flags[_Corrected].csv`: Master wide-format dataset with all spatial, pedigree, and corrected measurement data.
* `*_graphs[_Corrected].pdf`: Comprehensive diagnostic report featuring **Spearman Rho** coefficients on all correlation plots to verify the success of ID re-alignment.
* `*_Stats[_Corrected].csv`: Summary statistics (N, Mean, CV%, etc.) for all valid traits.
* `*_Trait_Correlations[_Corrected].csv`: A tidy table of all trait-to-trait Spearman rank correlations for the specific trial.
* `MASTER_Trait_Correlations.csv`: A synthesized global table combining correlation data across all processed trials in the series.
* `*_Traits.xml`: Standardized XML trait definitions for Dataplan upload.

### 🛑 Validation Flags
Any flags generated during processing are automatically consolidated into the `Validation_record` column. This includes:
* **Extreme Outliers:** Values exceeding Mean ± 4 SD.
* **Temporal Shrinkage:** Any repeated measurements (most often Ht or Dm) where latter-age stems are recorded as smaller than the previous age. The presence of these flags in a dataset will be indicated by the "Shrinkage" plots in the pdf report. If the data is included in the ASCII import, their units will be verified to be the same before flags are assigned. Repeated measurements of pilodyn and/or ordinal data are excluded from these flags.
* **ID Re-alignment:** If a tree was moved by the diagnostic logic, it will state: *"ID Realigned from [Path] (Anchored to [Trait])"*.
