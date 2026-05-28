# 1. Define your base paths
base_dir <- "C:/Users/james.baker/Forest Research/TW CBC-TBA-NextGenBritishConifers - Share"
demo_dir <- file.path(base_dir, "Demo")

parent_folders <- c(
  file.path(base_dir, "Sitka/Backwards Selected Fullsib P96-P99 experiments"),
  file.path(base_dir, "Sitka/High GCA Fullsib P85-P87 experiments")
)

# 2. Define the function to find, structure, and copy trial files
setup_demo_trial <- function(trial_name) {
  
  # Remove spaces to make searching robust
  search_term <- gsub(" ", "", tolower(trial_name))
  source_trial_dir <- NULL
  
  # Search the parent folders for the specific trial directory
  for (parent in parent_folders) {
    if (dir.exists(parent)) {
      subdirs <- list.dirs(parent, recursive = FALSE)
      match_idx <- grep(search_term, gsub(" ", "", tolower(basename(subdirs))))
      
      if (length(match_idx) > 0) {
        source_trial_dir <- subdirs[match_idx[1]]
        break
      }
    }
  }
  
  if (is.null(source_trial_dir)) {
    stop(paste("Could not find a folder matching", trial_name))
  }
  
  # 3. Create the nested directory structure: Demo > Parent Folder > Trial Folder
  parent_folder_name <- basename(dirname(source_trial_dir))
  target_trial_dir <- file.path(demo_dir, parent_folder_name, basename(source_trial_dir))
  
  # recursive = TRUE ensures it builds the Parent folder if it doesn't exist yet
  if (!dir.exists(target_trial_dir)) {
    dir.create(target_trial_dir, recursive = TRUE)
  }
  
  # 4. Define regex patterns for the files you need
  patterns <- c(
    "(?i)Matrix\\.csv$",
    "(?i)_ASCII\\.(csv|xlsx)$",
    "(?i)_DF\\.(txt|xlsx|csv)$",
    "(?i)[A-Za-z]+_[0-9]{2}\\.(txt|csv|xlsx)$" 
  )
  
  # 5. Find and copy the files
  all_files <- list.files(source_trial_dir, full.names = TRUE)
  files_to_copy <- c()
  
  for (pat in patterns) {
    matches <- grep(pat, all_files, value = TRUE)
    files_to_copy <- unique(c(files_to_copy, matches))
  }
  
  if (length(files_to_copy) == 0) {
    warning("No matching files found in ", source_trial_dir)
    return()
  }
  
  # Execute the copy
  file.copy(from = files_to_copy, to = target_trial_dir, overwrite = TRUE)
  message(paste("Successfully copied", length(files_to_copy), "files to", target_trial_dir))
}

# --- Execute ---
setup_demo_trial("Brecon 8")
setup_demo_trial("Kielder 162")
setup_demo_trial("Ae 58")
setup_demo_trial("Moray 55")
