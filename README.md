# DATAPLAN Prep: Unified Master Pipeline

This repository contains the R-based data processing pipeline for standardizing, validating, and formatting raw tree breeding data exported from the CEDD prior to upload to our Data Management System (DMS). 

It takes raw ASCII exports, merges them with design files and spatial matrices, applies validation rules (outlier detection, temporal shrinkage checks, spatial coordinate mapping), and outputs clean, wide-format datasets, graphical summaries, correlation tables, and `.xml` files.

After running this pipeline and performing preliminary data validations (See Validation Record SOP), the verified pedigree data (filtered by downloaded copies of DMS pedigree files) series trial data & trait definitions will be ready to import to DMS. See the `ASReml` Repo for next steps: performing univariate analyses and applying spatial corrections.    

## ⚠️ Important Setup Instructions (Read First)

This project uses a **hybrid file architecture**:
1. **The Code:** Lives locally on your `C:` drive (managed via this GitHub repository).
2. **The Data:** Lives externally on the shared network drive (`Z:`) or synced Teams/SharePoint folders.

**Do NOT copy raw data (`.csv`, `.xlsx`, `.txt`) into this local repository.** To run these scripts, you must define the `BASE_DIR`, `TRIAL_SERIES`, and other specific parameters in the `USER CONFIGURATION` block of `Master_Engine.R`.

**Pedigree Quirks:** The Pedigree system in our DMS strictly prevents uploading families, groups or genotypes that already exist in the system. Therefore, use the toggle `HAS_EXISTING_DB<-TRUE` and the setting `EXISTING_DIR    <- file.path(BASE_DIR, "SERIES NAME")` to filter the current series of trials against `DMS_all_fams.xlsx`/`DMS_all_genotypes.xlsx`/`DMS_all_groups.xlsx`, which need to be located in the parent directory of the previous series (e.g. Backwards Selected Fullsib P96-P99 experiments). These files can easily be downloaded from DMS by going to the pedigree pages for each of families, genotypes and groups and downlaoding all available data. 

Also, please note that the two files specifying controls and OP families (currently) are separate from other pedigree files, and should be updated prior to running the `Pedigree.R` script (see `Pedigree_diagnostics.R` for tools to summarize these for a new series). The reason for this is that they require input on origin and status for import. 

---

## 🏗️ The "Wheel and Engine" Architecture

This pipeline is structured into a Control Center (The Wheel) and modular processing scripts (The Engines). **Future users should only ever open and interact with `Master_Engine.R` for day-to-day usage.**

### 1. The Control Center: `Master_Engine.R`
This is the single entry point for the entire pipeline. Users set their global paths, database toggles, and trial specifications here. Once configured, you simply highlight and run the specific Execution Blocks you need. It automatically sources the required Engine files in the background.

### 2. The Engines (Located in the `R/` directory)
These files contain the underlying logic and should generally not be edited by standard users.

* **`DP_batch_process_Master.R` (Dataplan Engine):** Ingests raw ASCII/TXT files, merges design files and spatial matrices, enforces global zero-padding on assessment ages, flags outliers/shrinkage, and outputs the final `_Full_Data_With_Flags.csv` and XML files. It automatically detects and applies surgical ID re-alignments if a `TRAVERSAL_HELPER_MASTER.csv` is present.
* **`Data_orientation_corrections.R` (Traversal Engine):** Scans for spatial assessor errors via a maximally 16-way brute force simulation (also works automatically with alpha-row designs: i.e. 8x1 plots). It finds the traversal path that maximizes biological correlation against a trusted anchor. Generates a "recipe book" of fixes (`TRAVERSAL_HELPER_MASTER.csv`) for the Dataplan Engine to apply.
* **`Pedigree_diagnostics.R` (Pre-Flight Checks):** A diagnostic engine run before building the final pedigree. It compares your pending trial data against the existing DMS database, summarizing unique families, parent overlaps, and extracting instances of Open Pollination (OP) and Polymixes (PO) for manual review.
* **`Pedigree.R` (Pedigree Builder):** Generates final import files (`Groups`, `Genotypes`, `Families`) to facilitate batch uploads to the DMS. It cross-references local trial data against static DB downloads to prevent duplicate uploads. This script seamlessly processes Open Pollinated (OP) and Polymix (PO) families using strict regex evaluations.
* **`ASCII_diagnostics.R` (Utilities Engine):** Standalone tools for ground-truthing data, including functions to summarize ASCII inventory files, scan for duplicate measurements (e.g., repeating AV readings), and format raw spatial matrix files.

### 3. Other code (Also in `R/` directory)
* **`Bespoke_code_legacy.R` (Trial Specific code):** this script is included only for data legacy and reproducibility. It includes code that, for example, performs a complete mirroring of Kielder 162 data and testing whether Cr_07 in Radnor 55 is non-normally distrubuted. It is not intended to work with the refactored scripts.  
* **`Create_demo_dir.R` (Training/debugging tool):** this script contains some simple code to copy the requisite input files (using quite a few assumptions about naming conventions/file types) to new "Demo" directories. It was retained b/c I went to the bother of making it, and it can provide a fresh, clean directory to test out new features, debug code, etc.
*  **`read_BrSt_sheets.R` (Data processing tool):** this script contains code to read messy field data, using similar logic to the Automatic Resitograph processing repo, and convert it into machine-readible wide format data. It was designed specifically for our group's fieldsheets from Kintyre, so will likely require updating prior to using on a new dataset.  

---

## 🧬 Supported Family Designations

The pipeline natively supports standard Controlled Pollination (CP), alongside advanced parental configurations:
* **Open Pollinated (OP):** Handled natively by the pipeline.
* **Polymixes (PO):** The pipeline explicitly supports Polymix families. These are expected to be formatted as either standard `PO`, or `POXXY` (where `PO` is followed by at least two numbers and one letter). They are processed using the exact same logic as OP families.
* Other non-standard paternal specifications (basically any text proceeding the "_" that typically separates maternal and paternal IDS) will be read in as is, but may have inaccurate parent type or descriptions.  

---

## ⚙️ Workflow: Detect, Diagnose, & Chain

To ensure data integrity and validation due-diligence, the following workflow is recommended via `Master_Engine.R`:

1. **The Baseline Run (Detect):** Configure your paths in `Master_Engine.R` and run **Block 1**. This produces an initial `_Full_Data_With_Flags.csv` and the diagnostic `_graphs.pdf`. *(Note: Output files will NOT have the `_Corrected` suffix during a standard baseline run).*  
2. **Visual Inspection:** Review the generated PDF. Verify that trait names are being imported correctly (change in ASCII if erroneous), that any new traits have an entry in `trait_trans.csv` so metadata is produced accurately, that traits have the correct units (can change in .xml prior to upload), and that the spatial layout plots are correct. Inspect all plots for any other anomolies.
3. **Inspect data orientation among traits** Look for traits with poor correlations or severe left-skewing, as opposed to expected correlations. Most growth traits should have strong correlations among one another (typically 0.6-0.8): anything lower than that is likely problematic in some way. Note that some correlations may be negative (e.g. denisty and diameter), use the   `EXECPT_NEGATIVE_COR <- TRUE` toggle in these cases to identify cases with best negative correlation. 
4. **Establish an Anchor:** Identify a verified trait to use as your Absolute Anchor. Ideally, this trait should be verified by comparing the physical paper records against the ASCII files, but this is not always an option. 
5. **The Fix (Diagnose):** In `Master_Engine.R`, configure the Traversal settings and run **Block 2**. Test the suspect trait against your verified anchor. This saves recommendations to the Master Helper.
6. **The Chain (Speculative: USE WITH CAUTION):** Toggle `USE_CORRECTED_DATA <- TRUE` in the config block. The `Data_orientation_corrections.R` script will use your newly corrected trait as an anchor for further diagnostics. *Note: Chaining relies heavily on visually identifying consistent assessor error patterns and aborts automatically if a double-scramble is detected. Again, included for legacy and traceablilty, use with caution*
7. **The Final Pass (Correct):** Run **Block 1** (Master Script) one final time. It will seamlessly ingest the `TRAVERSAL_HELPER_MASTER.csv`, apply the fixes simultaneously, and output files cleanly marked with a `_Corrected` suffix.

---

## 📂 Folder Structure & Requirements

### 📁 Expected Trial Folder Input
The script dynamically trawls the `TRIAL_SERIES` folder for:
* **`*_ASCII.csv/xlsx`** *(Required)*: The raw PowerBI measurement export.
* **`*_Matrix.csv`** *(Soft Requirement)*: The physical spatial layout of the trial. Required for downstream spatial analyses and final upload to DMS. **Matrix File Preference:** If both a `.csv` and `.xlsx` exist in the folder, the pipeline enforces a preference check and explicitly loads the `.csv` version to prevent crash conditions. Loading the spatial matrix is handled silently to prevent R from printing a wall of auto-generated placeholder column names to the console.
* **`*_DF.txt/xlsx/extensionless`** *(Optional)*: The design file containing Crosses and Blocks. 
* **`*_AD_<age>.csv/xlsx`** *(Optional)*: Additional Data files to be merged, organized as a wide format data frame.
* **`*_<trait>_<age>.txt`** *(Optional)*: Additional data files to be merged, organizd as LONG format CEDD download. Pulls in units and meta from header, and uses name of file for trait name. 
* **`TRAVERSAL_HELPER_MASTER.csv`** *(Optional)*: The orientation "recipe book" generated by the Traversal Engine.

### 🛑 Strict Parser Rules & Troubleshooting

To ensure pipeline stability and prevent crashes during ingestion, adhere to the following data formatting assumptions:
* **Matrix CSV Preferences:** This code preferentially improts `_Matrix.csv` over `_layout.xslx` for Matrix files, however barring a `_Matrix.csv` it will fall back to the `_layout.xlxs` import. This was to a) prevent JV from renaming all thier `_layout.xlsx` sheets, while also avoiding the `_layout.xlsx` sheets that contain the excel formulas whichJB used to generate the `_Matrix.csv`
*  **NB: USERS SHOULD DEFAULT TO PREPARING AND USING `_Matrix.csv` files for clarity.   

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
