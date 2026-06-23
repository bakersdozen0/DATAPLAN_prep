# DATAPLAN Prep: Unified Master Pipeline

This repository contains the R-based data processing pipeline for standardizing, validating, and formatting raw tree breeding data exported from the CEDD prior to upload to our Data Management System (DMS). 

It takes raw ASCII exports, merges them with design files and spatial matrices, applies validation rules (outlier detection, temporal shrinkage checks, spatial coordinate mapping), and outputs clean, wide-format datasets, graphical summaries, correlation tables, and `.xml` files.

## ⚠️ Important Setup Instructions (Read First)

This project uses a **hybrid file architecture**:
1. **The Code:** Lives locally on your `C:` drive (managed via this GitHub repository).
2. **The Data:** Lives externally on the shared network drive (`Z:`) or synced Teams/SharePoint folders.

**Do NOT copy raw data (`.csv`, `.xlsx`, `.txt`) into this local repository.** To run these scripts, you must define the `BASE_DIR`, `TRIAL_SERIES`, and other specific parameters in the `USER CONFIGURATION` block of `Master_Engine.R`.

---

## 🏗️ The "Wheel and Engine" Architecture

This pipeline is structured into a Control Center (The Wheel) and modular processing scripts (The Engines). **Future users should only ever open and interact with `Master_Engine.R`.**

### 1. The Control Center: `Master_Engine.R`
This is the single entry point for the entire pipeline. Users set their global paths, database toggles, and trial specifications here. Once configured, you simply highlight and run the specific Execution Blocks you need. It automatically sources the required Engine files in the background.

### 2. The Engines (Located in the `R/` directory)
These files contain the underlying logic and should generally not be edited by standard users.

* **`DP_batch_process_Master.R` (Dataplan Engine):** Ingests raw ASCII/TXT files, merges design files and spatial matrices, enforces global zero-padding on assessment ages, flags outliers/shrinkage, and outputs the final `_Full_Data_With_Flags.csv` and XML files. It automatically detects and applies surgical ID re-alignments if a `TRAVERSAL_HELPER_MASTER.csv` is present.
* **`Data_orientation_corrections.R` (Traversal Engine):** Scans for spatial assessor errors via a 16-way brute force simulation. It finds the traversal path that maximizes biological correlation against a trusted anchor. Generates a "recipe book" of fixes (`TRAVERSAL_HELPER_MASTER.csv`) for the Dataplan Engine to apply.
* **`Pedigree_diagnostics.R` (Pre-Flight Checks):** A diagnostic engine run before building the final pedigree. It compares your pending trial data against the existing DMS database, summarizing unique families, parent overlaps, and extracting instances of Open Pollination (OP) and Polymixes (PO) for manual review.
* **`Pedigree.R` (Pedigree Builder):** Generates final import files (`Groups`, `Genotypes`, `Families`) to facilitate batch uploads to the DMS. It cross-references local trial data against static DB downloads to prevent duplicate uploads. This script seamlessly processes Open Pollinated (OP) and Polymix (PO) families using strict regex evaluations.
* **`ASCII_diagnostics.R` (Utilities Engine):** Standalone tools for ground-truthing data, including functions to summarize ASCII inventory files, scan for duplicate measurements (e.g., repeating AV readings), and format raw spatial matrix files.

---

## 🧬 Supported Family Designations

The pipeline natively supports standard Controlled Pollination (CP), alongside advanced parental configurations:
* **Open Pollinated (OP):** Handled natively by the pipeline.
* **Polymixes (PO):** The pipeline explicitly supports Polymix families. These are expected to be formatted as either standard `PO`, or `POXXY` (where `PO` is followed by at least two numbers and one letter). They are processed using the exact same logic as OP families.

---

## ⚙️ Workflow: Detect, Diagnose, & Chain

To ensure the highest data integrity in multi-tree trials, the following workflow is recommended via `Master_Engine.R`:

1. **The Baseline Run (Detect):** Configure your paths in `Master_Engine.R` and run **Block 1**. This produces an initial `_Full_Data_With_Flags.csv` and the diagnostic `_graphs.pdf`. *(Note: Output files will NOT have the `_Corrected` suffix during a standard baseline run).*
2. **Visual Inspection:** Review the correlation plots in the generated PDF. Look for traits with poor correlations or severe left-skewing against expected 1:1 baselines. 
3. **Establish an Anchor:** Identify a stable trait to use as your Absolute Anchor. Ideally, this trait should be manually verified by comparing the physical paper records against the CEDD database export.
4. **The Fix (Diagnose):** In `Master_Engine.R`, configure the Traversal settings and run **Block 2**. Test the suspect trait against your verified anchor. This saves recommendations to the Master Helper.
5. **The Chain (Speculative):** Toggle `USE_CORRECTED_DATA <- TRUE` in the config block. You can now use your newly corrected trait as an anchor for further diagnostics. *Note: Chaining relies heavily on visually identifying consistent assessor error patterns and aborts automatically if a double-scramble is detected.*
6. **The Final Pass (Correct):** Run **Block 1** (Master Script) one final time. It will seamlessly ingest the `TRAVERSAL_HELPER_MASTER.csv`, apply the fixes simultaneously, and output files cleanly marked with a `_Corrected` suffix.

---

## 📂 Folder Structure & Requirements

### 📁 Expected Trial Folder Input
The script dynamically trawls the `TRIAL_SERIES` folder for:
* **`*_ASCII.csv/xlsx`** *(Required)*: The raw PowerBI measurement export.
* **`*_Matrix.csv`** *(Soft Requirement)*: The physical spatial layout of the trial. Required for downstream spatial analyses and final upload to DMS. **Matrix File Preference:** If both a `.csv` and `.xlsx` exist in the folder, the pipeline enforces a preference check and explicitly loads the `.csv` version to prevent crash conditions. Loading the spatial matrix is handled silently to prevent R from printing a wall of auto-generated placeholder column names to the console.
* **`*_DF.txt/xlsx`** *(Optional)*: The design file containing Crosses and Blocks. **Note:** See strict parser rules below if using `.txt`.
* **`*_AD_<age>.csv/xlsx`** *(Optional)*: Additional Data files to be merged.
* **`TRAVERSAL_HELPER_MASTER.csv`** *(Optional)*: The orientation "recipe book" generated by the Traversal Engine.

### 🛑 Strict Parser Rules & Troubleshooting

To ensure pipeline stability and prevent crashes during ingestion, adhere to the following data formatting assumptions:
* **Matrix CSV Preferences:** Always prefer `.csv` for Matrix files. `.xlsx` Matrix files can introduce hidden metadata or formatting that causes read failures. 
* **Strict `_DF.txt` Assumptions:** If supplying Design Files as text (`.txt`), they must be strictly formatted. Ensure they use consistent delimiters (e.g., tabs), avoid trailing whitespace or hidden characters, and maintain consistent column headers that match expected pipeline inputs.
* **Polymix Regex Enforcement:** Ensure any PO families strictly match `PO` or `POXXY` regex formats, otherwise they may be flagged as invalid entries or drop out during pedigree building.

### 📊 Main Pipeline Outputs
*(The `_Corrected` suffix is automatically appended ONLY if spatial logic was successfully applied).*

* `*_Full_Data_With_Flags[_Corrected].csv`: Master wide-format dataset with all spatial, pedigree, and corrected measurement data.
* `*_graphs[_Corrected].pdf`: Comprehensive diagnostic report featuring Spearman Rho coefficients to verify the success of ID re-alignment.
* `*_Stats[_Corrected].csv`: Summary statistics (N, Mean, CV%, etc.) for all valid traits.
* `*_Trait_Correlations[_Corrected].csv`: A tidy table of all trait-to-trait Spearman rank correlations for the specific trial.
* `MASTER_Trait_Correlations.csv`: A synthesized global table combining correlation data across all processed trials in the series.
* `*_Traits.xml`: Standardized XML trait definitions for Dataplan upload.

### 🛑 Validation Flags
Any flags generated during processing are automatically consolidated into the `Validation_record` column. This includes:
* **Extreme Outliers:** Values exceeding Mean ± 4 SD.
* **Temporal Shrinkage:** Any repeated measurements (most often Ht or Dm) where latter-age stems are recorded as smaller than the previous age. 
* **ID Re-alignment:** If a tree was moved by the diagnostic logic, it will state: *"ID Realigned from [Path] (Anchored to [Trait])"*.
