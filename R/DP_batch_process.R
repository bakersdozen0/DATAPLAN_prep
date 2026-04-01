#### Library ####
library(here)
library(tidyverse)
library(readxl)
library(fs)
library(gridExtra)
library(grid)


# PART 0: HELPER FUNCTIONS #### 

# --- 1. Parse Design File (UPDATED: Force Numeric Plot IDs & Control Logic) ---
parse_long_design_file <- function(filepath, exp_prefix = "Exp") {
  raw_lines <- readLines(filepath)
  
  trt_start <- grep("TREATMENTS", raw_lines)
  plot_start <- grep("PLOTS", raw_lines)
  
  if(length(trt_start) == 0 || length(plot_start) == 0) return(NULL)
  
  # Parse TREATMENTS
  trt_lines <- raw_lines[(trt_start + 1):(plot_start - 1)]
  trt_lines <- trt_lines[grep("=", trt_lines)]
  
  trt_df <- tibble(line = trt_lines) %>%
    # Flag if the line starts with an asterisk (indicating a control) before we chop it up
    mutate(Is_Control = str_detect(str_trim(line), "^\\*")) %>%
    extract(line, into = c("Design_ID", "Raw_Cross_Name"), 
            regex = "^\\s*\\*?\\s*(\\d+)=(.*)$", convert = TRUE) %>%
    mutate(Raw_Cross_Name = str_trim(Raw_Cross_Name)) %>%
    # Cleanly separate the Mother and Father into their own columns for standard crosses
    extract(
      Raw_Cross_Name, 
      into = c("Maternal_ID", "Paternal_ID"), 
      regex = "^SS\\s*(\\d+)\\s*SS\\s*(.*)$", 
      remove = FALSE
    ) %>%
    mutate(
      Maternal_ID = str_trim(Maternal_ID),
      Paternal_ID = str_trim(Paternal_ID),
      # Extract the name for the controls by stripping the starting "SS "
      Control_Name = if_else(
        Is_Control,
        str_trim(str_replace(Raw_Cross_Name, "^SS\\s*", "")),
        NA_character_
      )
    )
  
  # Parse PLOTS
  plot_lines <- raw_lines[(plot_start + 1):length(raw_lines)]
  plot_lines <- plot_lines[grep(":", plot_lines)]
  
  plot_df <- tibble(line = plot_lines) %>%
    extract(line, into = c("Plot", "Design_ID", "Block"), 
            regex = "^\\s*(\\d+):\\s*(\\d+)\\s+(\\d+)", convert = TRUE)
  
  # Join and standardize PLOT to NUMERIC
  final_design <- plot_df %>%
    left_join(trt_df, by = "Design_ID") %>%
    mutate(
      Plot = as.numeric(Plot), # <--- FORCE NUMERIC HERE
      Cross_Name = case_when(
        Design_ID == 0 ~ paste0(exp_prefix, "_Filler"),
        
        # Map controls exactly like the multi-tree scripts (ss_ prefix)
        !is.na(Control_Name) ~ paste0("ss_", Control_Name),
        
        # Open Pollinated
        !is.na(Paternal_ID) & str_detect(Paternal_ID, "(?i)OP") ~ paste0("ss", Maternal_ID, "_OPCB"),
        
        # Standard valid cross
        !is.na(Maternal_ID) & !is.na(Paternal_ID) ~ paste0("ss", Maternal_ID, "_ss", Paternal_ID),
        
        # Fallback for broken formats
        TRUE ~ str_replace_all(Raw_Cross_Name, "\\s+", "") %>% str_replace("^SS", "ss_")
      )
    ) %>%
    select(Plot, Design_ID, Block, Cross_Name)
  
  return(final_design)
}

# --- 2. PDF Reporting Function (UPDATED: Color Correlations + % Concentric Maps) ---
generate_pdf_report <- function(long_data, wide_data, exp_name, multi_age_prefixes, all_traits) {
  
  # 1. Get ALL continuous traits first (Excluding Survival)
  all_cont_traits <- long_data %>% 
    filter(Is_Ordinal == FALSE, !str_detect(Trait, "(?i)Sur_")) %>% 
    distinct(Trait) %>% pull(Trait)
  
  # 2. Split into "Integer-Based" (Pilodyn) and "Standard Continuous"
  # Pilodyn is technically continuous but measured in integers, so bins=30 looks bad.
  pil_traits  <- all_cont_traits[str_detect(all_cont_traits, "(?i)Pil")]
  cont_traits <- setdiff(all_cont_traits, pil_traits) # Everything else (Hgt, DBH, etc.)
  
  ord_traits  <- long_data %>% 
    filter(Is_Ordinal == TRUE, !str_detect(Trait, "(?i)Sur_"))  %>% 
    distinct(Trait) %>% pull(Trait)
  
  # --- A1. Distributions (Standard Continuous) ---
  # Uses bins = 30
  plot_data_cont <- long_data %>% filter(Trait %in% cont_traits, !is.na(Value_Num))
  if (nrow(plot_data_cont) > 0) {
    p1 <- ggplot(plot_data_cont, aes(x = Value_Num, fill = as.character(Flag_Outlier))) +
      geom_histogram(bins = 30, color = "black", alpha = 0.7) +
      facet_wrap(~Trait, scales = "free") +
      scale_fill_manual(values = c("FALSE" = "lightblue", "TRUE" = "orange"), name = "Is Outlier?") +
      labs(title = paste(exp_name, "- Distributions (Continuous)"), subtitle = "Standard (bins=30)", x = "Value") + 
      theme_minimal() + theme(legend.position = "bottom")
    print(p1)
  }
  
  # --- A2. Distributions (Pilodyn / Integer) ---
  # Uses binwidth = 1 to show every unique integer value
  plot_data_pil <- long_data %>% filter(Trait %in% pil_traits, !is.na(Value_Num))
  if (nrow(plot_data_pil) > 0) {
    p1_pil <- ggplot(plot_data_pil, aes(x = Value_Num, fill = as.character(Flag_Outlier))) +
      # [FIX] Use binwidth = 1 to create one bin per integer
      geom_histogram(binwidth = 1, color = "black", alpha = 0.7, center = 0) +
      facet_wrap(~Trait, scales = "free") +
      scale_fill_manual(values = c("FALSE" = "lightblue", "TRUE" = "orange"), name = "Is Outlier?") +
      labs(title = paste(exp_name, "- Distributions (Pilodyn)"), subtitle = "Integer-based (binwidth=1)", x = "Value") + 
      theme_minimal() + theme(legend.position = "bottom")
    print(p1_pil)
  }
  
  # --- B. Distributions (Ordinal) ---
  plot_data_ord <- long_data %>% filter(Trait %in% ord_traits, !is.na(Value_Num))
  if (nrow(plot_data_ord) > 0) {
    p2 <- ggplot(plot_data_ord, aes(x = factor(Value_Num), fill = as.character(Flag_Outlier))) +
      geom_bar(color = "black", alpha = 0.7) +
      facet_wrap(~Trait, scales = "free") +
      scale_fill_manual(values = c("FALSE" = "lightgreen", "TRUE" = "red"), name = "Invalid?") +
      labs(title = paste(exp_name, "- Distributions (Ordinal)"), subtitle = "Excluding Survival", x = "Score") +
      theme_minimal() + theme(legend.position = "bottom")
    print(p2)
  }
  
  # --- C. Shrinkage Checks (Scatter Plots) ---
  if (length(all_traits) > 0) {
    
    # 1. Extract base traits directly from the translated list (e.g., "HGT_10" -> "HGT")
    # This ignores traits without an age suffix.
    base_traits <- unique(str_remove(all_traits[str_detect(all_traits, "_\\d+$")], "_\\d+$"))
    
    for (base_t in base_traits) {
      # Skip traits where shrinkage is irrelevant or expected
      if (str_detect(base_t, "(?i)(Sur|Resi|Detrended|Ampli|Resistance|Resist|Crack|Form)")) next
      
      # 2. STRICTLY match only traits with this exact base and an age number
      # e.g., "^HGT_\\d+$" finds HGT_10 and HGT_15, but ignores HGT_ADJ_10
      regex_pattern <- paste0("^", base_t, "_\\d+$")
      cols <- str_sort(all_traits[str_detect(all_traits, regex_pattern)], numeric = TRUE)
      
      if (length(cols) < 2) next
      
      for (i in 1:(length(cols) - 1)) {
        younger <- cols[i]; older <- cols[i+1]
        if(!all(c(younger, older) %in% names(wide_data))) next
        
        plot_df <- wide_data %>%
          # Single-tree: only selecting Plot
          select(Plot, X = all_of(older), Y = all_of(younger)) %>%
          filter(!is.na(X), !is.na(Y)) %>%
          mutate(Is_Shrinkage = Y > X, Status = if_else(Is_Shrinkage, "Shrinkage", "Normal"))
        
        if(nrow(plot_df) == 0) next
        
        p_shrink <- ggplot(plot_df, aes(x = X, y = Y)) +
          geom_abline(slope = 1, intercept = 0, color = "gray50", linetype = "dashed") +
          geom_point(aes(color = Status), size = 2, alpha = 0.6) +
          scale_color_manual(values = c("Normal"="black", "Shrinkage"="red")) +
          labs(title = "Shrinkage Check", subtitle = paste(younger, ">", older), x = older, y = younger) +
          theme_minimal()
        print(p_shrink)
      }
    }
  }
  
  # --- D. Concentric Temporal Check (Final: Every 5th Label) ---
  if (all(c("Prow", "Ppos") %in% names(wide_data))) {
    
    # 1. Get Master List of Plots
    all_plots <- unique(wide_data$Plot)
    
    # 2. Filter Data: Exclude Survival-only ages
    physical_data <- long_data %>%
      mutate(Age_Num = suppressWarnings(as.numeric(as.character(Age)))) %>%
      filter(
        !str_detect(Trait, "(?i)Sur_"), 
        !is.na(Value_Num),              
        !is.na(Age_Num),                
        Age_Num > 0                     
      )
    
    # 3. Identify Ages to Plot
    plot_ages <- sort(unique(physical_data$Age_Num), decreasing = TRUE)
    
    if (length(plot_ages) > 0) {
      
      age_label <- paste(plot_ages, collapse = ", ")
      
      # 4. Calculate Counts
      counts_per_tree <- physical_data %>%
        group_by(Plot, Age_Num) %>%
        summarise(Count = n(), .groups = "drop")
      
      # 5. FILL GAPS
      full_grid <- expand_grid(Plot = all_plots, Age_Num = plot_ages) %>%
        left_join(counts_per_tree, by = c("Plot", "Age_Num")) %>%
        replace_na(list(Count = 0))
      
      # 6. Calculate Mode
      mode_per_age <- full_grid %>%
        filter(Count > 0) %>% 
        group_by(Age_Num) %>%
        summarise(
          Standard_Count = as.numeric(names(sort(table(Count), decreasing=TRUE)[1])), 
          .groups = "drop"
        )
      
      # 7. Join and Calculate Percent
      plot_data <- full_grid %>%
        left_join(mode_per_age, by = "Age_Num") %>%
        mutate(
          Standard_Count = replace_na(Standard_Count, 1), 
          Percent = (Count / Standard_Count) * 100
        ) %>%
        left_join(wide_data %>% select(Plot, Prow, Ppos), by = "Plot") %>%
        filter(!is.na(Prow), !is.na(Ppos))
      
      # 8. Define Axis Breaks (Show every 5th Row/Pos)
      min_row <- min(plot_data$Prow, na.rm = TRUE)
      max_row <- max(plot_data$Prow, na.rm = TRUE)
      min_pos <- min(plot_data$Ppos, na.rm = TRUE)
      max_pos <- max(plot_data$Ppos, na.rm = TRUE)
      
      # 9. Plotting
      sizes <- seq(0.95, 0.25, length.out = length(plot_ages))
      
      p_concentric <- ggplot() +
        # [UPDATED] Show label every 5 units
        scale_y_reverse(breaks = seq(min_row, max_row, by = 5)) +
        scale_x_continuous(breaks = seq(min_pos, max_pos, by = 5)) +
        
        coord_fixed() +
        theme_minimal() +
        
        theme(axis.text.x = element_text(angle = 90, vjust = 0.5, size = 8),
              axis.text.y = element_text(size = 8)) +
        
        labs(title = paste(exp_name, "- Data Completeness"),
             subtitle = paste0("Ages Plotted: ", age_label, "\nOuter Ring = Oldest -> Inner Square = Youngest\nRed = 0% | Yellow = 50% | Green = 100%+"),
             x = "Position", y = "Row", fill = "% Data") +
        
        scale_fill_gradient2(
          low = "red", 
          mid = "yellow", 
          high = "forestgreen", 
          midpoint = 50, 
          limits = c(0, 100),
          oob = scales::squish
        )
      
      for (i in seq_along(plot_ages)) {
        curr_age <- plot_ages[i]
        curr_size <- sizes[i]
        layer_data <- plot_data %>% filter(Age_Num == curr_age)
        
        p_concentric <- p_concentric + 
          geom_tile(data = layer_data, 
                    aes(x = Ppos, y = Prow, fill = Percent), 
                    width = curr_size, height = curr_size, 
                    color = "black", size = 0.1) 
      }
      print(p_concentric)
    }
  }
  
  # --- H. Resurrection Check (Zombie) - [UPDATED: TEXT ONLY] ---
  sur_cols <- sort(names(wide_data)[str_detect(names(wide_data), "(?i)Sur_\\d+")])
  
  if (length(sur_cols) >= 2) {
    sur_ages <- as.numeric(str_extract(sur_cols, "\\d+"))
    sur_map  <- data.frame(Col = sur_cols, Age = sur_ages) %>% arrange(Age)
    
    for (i in 1:(nrow(sur_map) - 1)) {
      younger_col <- sur_map$Col[i]
      older_col   <- sur_map$Col[i+1]
      
      zombies <- wide_data %>%
        filter(.data[[younger_col]] == 0 & .data[[older_col]] == 1) %>%
        select(Plot) %>% 
        distinct() %>%
        arrange(Plot)
      
      if (nrow(zombies) > 0) {
        
        # 1. [HASHED OUT] Plotting Code Removed
        # p_zom <- ggplot(zombies, aes(x = Pos, y = Row)) + ...
        
        # 2. Prepare Text List
        bad_ids <- sort(unique(zombies$Plot))
        id_string <- paste(bad_ids, collapse = ", ")
        
        # Wrap text so it doesn't run off the page (approx 80 chars wide)
        wrapped_text <- str_wrap(id_string, width = 80)
        
        # Construct a full title + list string
        final_text <- paste0(
          "RESURRECTION CHECK (Zombie Trees)\n",
          "Logic: Marked Dead (0) at ", younger_col, " -> Alive (1) at ", older_col, "\n",
          "Total Flagged: ", length(bad_ids), "\n\n",
          "Plot IDs:\n", wrapped_text
        )
        
        t_grob <- textGrob(final_text, 
                           x = 0.5, y = 0.5, # Center on page
                           just = "center", 
                           gp = gpar(fontsize = 10, fontface = "plain"))
        
        # 3. [REPLACEMENT] Print ONLY the text grob
        grid.arrange(t_grob)
      }
    }
  }
  
  # --- E. Correlations (Updated: Color Outliers) ---
  if (length(cont_traits) >= 2) {
    gen_pairs <- combn(cont_traits, 2, simplify = FALSE)
    for (pair in gen_pairs) {
      t1 <- pair[1]; t2 <- pair[2]
      if(!all(c(t1, t2) %in% names(wide_data))) next
      
      # 1. Identify Outliers for these specific traits
      outlier_plots <- long_data %>% 
        filter(Trait %in% c(t1, t2), Flag_Outlier == TRUE) %>% 
        pull(Plot) %>% unique()
      
      # 2. Build Plot Data
      plot_df <- wide_data %>%
        select(Plot, X_Val = all_of(t1), Y_Val = all_of(t2)) %>%
        filter(!is.na(X_Val), !is.na(Y_Val)) %>%
        mutate(Status = if_else(Plot %in% outlier_plots, "Outlier", "Normal"))
      
      if(nrow(plot_df) == 0) next
      
      p_corr <- ggplot(plot_df, aes(x=X_Val, y=Y_Val)) + 
        geom_smooth(method = "lm", se = FALSE, color = "gray", linetype = "dashed", alpha=0.5) +
        # Color points based on Outlier Status
        geom_point(aes(color = Status), alpha = 0.6) + 
        scale_color_manual(values = c("Normal" = "black", "Outlier" = "orange")) +
        labs(title="Correlation Check", x = t1, y = t2) + 
        theme_minimal()
      print(p_corr)
    }
  }
  
  # --- F. Heatmaps (All Traits) ---
  if (all(c("Prow", "Ppos") %in% names(wide_data))) {
    heatmap_traits <- c(cont_traits, ord_traits)
    
    for (trait in heatmap_traits) {
      if (!trait %in% names(wide_data)) next
      map_data <- wide_data %>% select(Prow, Ppos, Value = all_of(trait)) %>% filter(!is.na(Value))
      
      if (nrow(map_data) == 0) next
      p_map <- ggplot(map_data, aes(x = Ppos, y = Prow, fill = Value)) +
        geom_tile(color = "white", size = 0.1) +
        scale_fill_distiller(palette = "Spectral", direction = -1) +
        scale_y_reverse() +
        labs(title = paste(exp_name, "- Spatial Map:", trait)) +
        theme_minimal() + coord_fixed()
      print(p_map)
    }
  }
}

# --- 3. Spatial Matrix Processor ---
prowppos <- function(matrix_data) {
  data <- list()
  plotchecklist <- c()
  matrix_data <- as.matrix(matrix_data)
  for (i in seq_len(nrow(matrix_data))) {
    for (j in seq_len(ncol(matrix_data))) {
      plot <- matrix_data[i, j]
      if (is.na(plot) || plot == 0) next
      tree <- if (!(plot %in% plotchecklist)) 1 else sum(plotchecklist == plot) + 1
      data <- append(data, list(c(plot, tree, i, j)))
      plotchecklist <- c(plotchecklist, plot)
    }
  }
  if(length(data) == 0) return(NULL)
  df <- as.data.frame(do.call(rbind, data))
  colnames(df) <- c("Plot", "Tree", "Prow", "Ppos")
  return(df)
}


# --- Load Translation Map (UPDATED FOR StdTraits SHEET) ---
trait_path <- here("PPGTraits_UK.xlsm")

if (file.exists(trait_path)) {
  trait_map <- read_excel(trait_path, sheet = "StdTraits") %>%
    # 1. Select & Rename Columns based on your file structure
    select(
      trait_code_FR = `Eng Desc`,  # Source (ASCII Header)
      trait_code_DP = Trait,       # Destination (Dataplan Code)
      xml_group     = Group,       # XML Group
      xml_desc_base = Description, # XML Description
      xml_units     = Units        # XML Units
    ) %>%
    
    # 2. Clean Data
    mutate(
      trait_code_FR = as.character(trait_code_FR),
      trait_code_DP = as.character(trait_code_DP),
      xml_group     = as.character(xml_group),
      xml_desc_base = as.character(xml_desc_base),
      xml_units     = as.character(xml_units)
    ) %>%
    
    # 3. Filter: We only want rows that have an 'Eng Desc' to map FROM
    filter(!is.na(trait_code_FR), !is.na(trait_code_DP)) %>%
    arrange(desc(nchar(trait_code_FR)))
  
} else {
  stop(paste("Error: Trait file not found at:", trait_path))
}


# PART 1: LOCATE TRIAL FOLDERS ####

all_dirs <- list.dirs(path = here(), recursive = FALSE, full.names = FALSE)
experiments_to_process <- setdiff(all_dirs, c("00_Scripts", "Archive", ".git", ".Rproj.user"))

message(paste("Found", length(experiments_to_process), "folders to check."))

# Instead of searching the directory, we force the list to be just one folder.
experiments_to_process <- c("Kintyre 17","Kintyre 18")

# PART 2: THE BATCH LOOP ####

for (curr_exp in experiments_to_process) {
  
  tryCatch({
    message(paste("\n=========================================="))
    message(paste("Processing Folder:", curr_exp))
    
    exp_path <- here("Backwards Selected Fullsib P96-P99 experiments",curr_exp)
    file_prefix <- str_replace_all(curr_exp, " ", "_")
    
    # Initialize variables
    genetic_info <- NULL
    spatial_info <- NULL
    exp_data_long <- NULL
    
    # ------------------------------------------------------------------------
    # STEP 1: LOAD RAW FILES & DESIGN
    # ------------------------------------------------------------------------
    meas_filename <- paste0(file_prefix, "_ASCII.xlsx")
    meas_path <- file.path(exp_path, meas_filename)
    
    if (!file.exists(meas_path)) {
      message(paste("  -> SKIPPING: Measurement file not found:", meas_filename))
      next 
    }
    
    design_filename <- paste0(file_prefix, "_DF.txt")
    design_path <- file.path(exp_path, design_filename)
    
    # --- Load Design ---
    if (file.exists(design_path)) {
      message(paste("  -> Loading Design:", design_filename))
      genetic_info <- parse_long_design_file(design_path)
      
      if(!is.null(genetic_info)) {
        
        # --- ROBUST BLOCK FIXING ---
        valid_data <- genetic_info %>% 
          filter(Block != "0") %>% 
          mutate(Plot_Num = as.numeric(Plot), Block_Num = as.numeric(Block))
        
        if(nrow(valid_data) > 0) {
          block_starts <- valid_data %>%
            group_by(Block_Num) %>%
            summarise(Start_Plot = min(Plot_Num), .groups = "drop") %>%
            arrange(Block_Num)
          
          if(nrow(block_starts) > 1) {
            strides <- diff(block_starts$Start_Plot)
            block_size <- as.numeric(names(sort(table(strides), decreasing = TRUE)[1]))
          } else {
            block_size <- max(valid_data$Plot_Num)
          }
          
          genetic_info <- genetic_info %>%
            mutate(
              Block = if_else(Block == "0", 
                              as.character(ceiling(Plot / block_size)), 
                              as.character(Block))
            )
        }
        
        # --- FAMILY NAMING ---
        genetic_info <- genetic_info %>%
          mutate(
            Clean_ID = str_remove_all(str_remove(Cross_Name, "^\\*\\s*"), "\\s+"),
            Is_Full_Cross = str_count(Clean_ID, "SS") >= 2,
            Family_name = case_when(
              Design_ID == "0" ~ paste0(file_prefix, "_Filler"),
              Is_Full_Cross ~ str_replace(Clean_ID, "^SS(.+)SS(.+)$", "ss\\1_ss\\2"),
              TRUE ~ str_replace(Clean_ID, "^SS", "ss")
            )
          ) %>%
          select(Plot, Block, Family_name)
      }
    }
    
    # --- Load Spatial Matrix ---
    matrix_files <- dir_ls(exp_path, regexp = "(?i)_Matrix\\.csv$")
    if (length(matrix_files) > 0) {
      raw_matrix <- read.csv(matrix_files[1], header = FALSE)
      raw_matrix[is.na(raw_matrix)] <- 0
      spatial_info <- prowppos(raw_matrix)
      if (!is.null(spatial_info)) {
        spatial_info <- spatial_info %>%
          mutate(Plot = as.numeric(Plot)) %>% # Force Numeric
          select(Plot, Prow, Ppos) %>%
          distinct(Plot, .keep_all = TRUE)
      }
    }
    
    # ------------------------------------------------------------------------
    # STEP 2: MEASUREMENTS & RENAMING
    # ------------------------------------------------------------------------
    raw_data <- read_excel(meas_path, col_types = "text")
    
    if(!is.null(genetic_info)) {
      raw_data <- raw_data %>% 
        mutate(Plot_Check = suppressWarnings(as.numeric(Plot))) %>%
        filter(Plot_Check %in% genetic_info$Plot) %>%
        select(-Plot_Check)
    }
    
    exp_data_long <- raw_data %>%
      select(Plot_Orig = Plot, Value_Char = Measurment, TraitCode = AssessmentType_long, 
             UnitCode = AssessmentUnit_long, Age = AssessmentYear_long, Date_Raw = AssessmentDate) %>%
      filter(!is.na(Plot_Orig)) %>%
      mutate(
        Plot = suppressWarnings(as.numeric(Plot_Orig)),
        Value_Num = suppressWarnings(as.numeric(Value_Char)),
        Prefix = str_to_upper(str_trim(replace_na(TraitCode, ""))),
        Age = as.character(Age),
        Date = as.Date(as.character(Date_Raw), format = "%d%m%Y"),
        Trait_Orig = paste0(Prefix, replace_na(UnitCode, ""), Age)
      )
    
    # --- RENAMING LOGIC ---
    unique_prefixes <- exp_data_long %>% distinct(Prefix)
    unique_prefixes$trait_code_DP <- map_chr(unique_prefixes$Prefix, function(p) {
      if(is.na(p) || p == "") return(NA_character_)
      match <- trait_map %>% filter(str_starts(p, trait_code_FR)) %>% slice(1) %>% pull(trait_code_DP)
      if(length(match) == 0) return(NA_character_) else return(match)
    })
    
    exp_data_long <- exp_data_long %>%
      left_join(unique_prefixes, by = "Prefix") %>%
      mutate(
        Trait = case_when(
          !is.na(trait_code_DP) & !is.na(Age) ~ paste0(trait_code_DP, "_", Age),
          !is.na(trait_code_DP) & is.na(Age)  ~ trait_code_DP,
          TRUE ~ Trait_Orig
        )
      ) %>%
      select(Plot, Trait, Value_Char, Value_Num, UnitCode, Trait_Orig, Age, Prefix, Date)
    
    # --- STEP 2.4: GENERALIZED ADDITIONAL DATA (AD FILES) ---
    # Looks for any file matching *_AD_*.csv or *_AD_*.xlsx
    ad_files <- dir_ls(exp_path, regexp = "(?i)_AD_(\\d+)\\.(csv|xlsx)$")
    
    if (length(ad_files) > 0) {
      for (ad_path in ad_files) {
        # Extract the age from the filename (e.g., "10" from "Wark_16_AD_10.xlsx")
        ad_age <- str_extract(basename(ad_path), "(?i)_AD_(\\d+)") %>% str_extract("\\d+")
        message(paste("  -> Loading Additional Data:", basename(ad_path)))
        
        # Read file depending on extension
        if (str_detect(basename(ad_path), "(?i)\\.csv$")) {
          ad_raw <- read_csv(ad_path, col_types = cols(.default = "c"))
        } else {
          ad_raw <- read_excel(ad_path, col_types = "text")
        }
        
        # Standardize standard column names (case-insensitive)
        names(ad_raw) <- str_to_upper(names(ad_raw))
        if("PLOT" %in% names(ad_raw)) ad_raw <- rename(ad_raw, Plot = PLOT)
        if("TREE" %in% names(ad_raw)) ad_raw <- rename(ad_raw, Tree = TREE)
        if("STEM" %in% names(ad_raw)) ad_raw <- rename(ad_raw, Tree = STEM)
        
        # Ensure we have a Plot column before proceeding
        if ("Plot" %in% names(ad_raw)) {
          # Pivot longer, ignoring metadata columns if they exist
          cols_to_exclude <- names(ad_raw)[names(ad_raw) %in% c("Plot", "Tree", "BLOCK", "REP")]
          
          ad_long_raw <- ad_raw %>%
            pivot_longer(cols = -any_of(cols_to_exclude), names_to = "Header_Original", values_to = "Value_Char") %>%
            filter(!is.na(Value_Char))
          
          if (nrow(ad_long_raw) > 0) {
            
            # Map headers to DP traits using the trait_map (same logic as the old resi block)
            unique_headers <- unique(ad_long_raw$Header_Original)
            header_mapping <- map_dfr(unique_headers, function(h) {
              clean_h_strict <- str_remove_all(str_to_upper(h), "[^A-Z0-9]")
              
              # Try exact match first, then prefix match
              match <- trait_map %>% 
                mutate(match_key_strict = str_remove_all(str_to_upper(trait_code_FR), "[^A-Z0-9]")) %>%
                filter(clean_h_strict == match_key_strict) %>% 
                slice(1)
              
              if(nrow(match) == 0) {
                match <- trait_map %>% 
                  mutate(match_key_strict = str_remove_all(str_to_upper(trait_code_FR), "[^A-Z0-9]")) %>%
                  filter(str_starts(clean_h_strict, match_key_strict)) %>% 
                  slice(1)
              }
              
              if(nrow(match) > 0) {
                tibble(Header_Original = h, Code_DP = match$trait_code_DP, Key_FR = match$trait_code_FR)
              } else {
                # Fallback if it isn't in the translation map: just use the letters from the header
                fallback_code <- str_to_upper(str_extract(h, "^[A-Z]+"))
                tibble(Header_Original = h, Code_DP = replace_na(fallback_code, h), Key_FR = h)
              }
            })
            
            ad_long <- ad_long_raw %>%
              left_join(header_mapping, by = "Header_Original") %>%
              mutate(
                Plot = suppressWarnings(as.numeric(Plot)),
                # If Tree doesn't exist in the file, assume it's plot-level data and assign to Tree 1
                Tree = if("Tree" %in% names(.)) suppressWarnings(as.numeric(Tree)) else 1,
                Value_Num = suppressWarnings(as.numeric(Value_Char)),
                Age = ad_age,
                Trait = paste0(Code_DP, "_", Age), 
                Prefix = Key_FR,
                Trait_Orig = Header_Original, 
                UnitCode = "", 
                Date = as.Date(NA)
              ) %>%
              select(any_of(names(exp_data_long))) 
            
            exp_data_long <- bind_rows(exp_data_long, ad_long)
          }
        } else {
          message("    -> WARNING: No 'Plot' column found in AD file. Skipping.")
        }
      }
    }
    
    # --- STEP 2.5: SURVIVAL CALC ---
    if(!is.null(genetic_info)) {
      raw_ages <- unique(na.omit(exp_data_long$Age))
      unique_ages <- raw_ages[!raw_ages %in% c("1", "01")]
      sur_list <- list()
      
      for(curr_age in unique_ages) {
        t_name <- paste0("Sur_", curr_age)
        if(t_name %in% unique(exp_data_long$Trait)) next
        alive_plots <- exp_data_long %>% filter(Age == curr_age, !is.na(Value_Num)) %>% pull(Plot) %>% unique()
        sur_df <- tibble(
          Plot = genetic_info$Plot,
          Trait = t_name, Age = curr_age,
          Value_Num = if_else(Plot %in% alive_plots, 1, 0),
          Value_Char = as.character(if_else(Plot %in% alive_plots, 1, 0)),
          UnitCode = "$1-1", Trait_Orig = paste0("Calculated_", t_name), Prefix = "SUR", Date = as.Date(NA)
        )
        sur_list[[t_name]] <- sur_df
      }
      if(length(sur_list) > 0) exp_data_long <- bind_rows(exp_data_long, bind_rows(sur_list))
    }
    
    # ----------------------------------------------------------------------------
    # STEP 3: FLAGS & LOGIC (FIXED: CULL & ZERO HANDLING)
    # ----------------------------------------------------------------------------
    
    # STRICT ZERO CLEANING
    exp_data_long <- exp_data_long %>%
      mutate(
        # Regex to identify continuous traits where 0 is an error (missing)
        Is_Target_For_Zero_Removal = str_detect(Trait, "(?i)^(Pil|Br|St)"),
        Value_Num = case_when(
          Is_Target_For_Zero_Removal & Value_Num == 0 ~ NA_real_,
          TRUE ~ Value_Num
        ),
        Value_Char = if_else(is.na(Value_Num), NA_character_, as.character(Value_Num))
      )
    
    # STATS & OUTLIERS
    data_with_outliers <- exp_data_long %>%
      select(-Is_Target_For_Zero_Removal) %>% 
      group_by(Trait) %>%
      mutate(
        n_unique = n_distinct(Value_Num, na.rm = TRUE),
        
        Is_Survival = str_detect(Trait, "(?i)Sur"),
        
        # Updated Regex to include "Cull" and "Stem" as Binary types
        # This prevents Cull=0 from being flagged as an "Ordinal 0" error.
        Is_Binary_Resi = str_detect(Trait, "(?i)(Resi|Drill|Amp|Den|Cull|Stem)") & n_unique <= 2,
        
        # Ordinal Definition: Low unique count, but NOT Survival or Binary
        Is_Ordinal  = !Is_Survival & !Is_Binary_Resi & n_unique < 9, 
        
        Mean_Val = mean(Value_Num, na.rm = TRUE),
        SD_Val   = sd(Value_Num, na.rm = TRUE),
        Lower_Limit = Mean_Val - (4 * SD_Val), 
        Upper_Limit = Mean_Val + (4 * SD_Val),
        
        Flag_Outlier = case_when(
          
          # Rule 1: Ordinal 0s (Check for NA first)
          !is.na(Value_Num) & Is_Ordinal & Value_Num == 0 ~ TRUE,
          
          # Rule 1: Ordinal 0s are invalid (1-8 scale), BUT "Cull" is now protected above
          Is_Ordinal & Value_Num == 0 ~ TRUE,
          
          # Rule 2: Binary/Survival (0-1) -> Values > 1 are invalid
          (Is_Survival | Is_Binary_Resi) & Value_Num > 1 ~ TRUE, 
          
          # Rule 3: Extreme Continuous Outliers
          !Is_Survival & !Is_Ordinal & !Is_Binary_Resi & !is.na(Value_Num) & 
            (Value_Num < Lower_Limit | Value_Num > Upper_Limit) ~ TRUE,
          
          TRUE ~ FALSE
        )
      ) %>%
      ungroup()
    
    # [NEW] Calculate Shrinkage Flags (Single-Tree)
    shrinkage_flags <- data_with_outliers %>%
      filter(!is.na(Value_Num), str_detect(Trait, "_\\d+$")) %>%
      filter(!str_detect(Trait, "(?i)(Sur|Resi|Detrended|Ampli|Resistance|Resist|Crack|Form)")) %>%
      mutate(
        Base_Trait = str_remove(Trait, "_\\d+$"),
        Age_Num = suppressWarnings(as.numeric(str_extract(Trait, "\\d+$")))
      ) %>%
      filter(!is.na(Age_Num)) %>%
      arrange(Plot, Base_Trait, Age_Num) %>%
      group_by(Plot, Base_Trait) %>%
      mutate(
        Prev_Val = lag(Value_Num),
        Prev_Age = lag(Age_Num),
        Is_Shrinkage = !is.na(Prev_Val) & (Value_Num < Prev_Val)
      ) %>%
      ungroup() %>%
      filter(Is_Shrinkage) %>%
      select(Plot, Trait, Prev_Age) %>%
      mutate(Shrinkage_Error = paste("Shrinkage detected (smaller than Age", Prev_Age, ")"))
    
    final_long_data <- data_with_outliers %>%
      left_join(shrinkage_flags %>% select(-Prev_Age), by = c("Plot", "Trait")) %>%
      mutate(
        Reject_Flag = case_when(
          is.na(Value_Num) ~ NA_real_,
          Flag_Outlier ~ 1,
          !is.na(Shrinkage_Error) ~ 1,      # <-- Shrinkage now triggers a reject!
          TRUE ~ 0
        )
      )
    
    # ----------------------------------------------------------------------------
    # STEP 4: OUTPUT GENERATION
    # ----------------------------------------------------------------------------
    
    # 1. Handle Duplicates
    final_long_dedup <- final_long_data %>%
      arrange(Plot, Trait, desc(Date)) %>%
      distinct(Plot, Trait, .keep_all = TRUE)
    
    # 2. Generate "Validation_record" Column (Single-Tree)
    # Aggregates outlier & shrinkage messages per Plot
    validation_summary <- final_long_dedup %>%
      filter(Reject_Flag == 1) %>%
      mutate(
        Error_Msg = case_when(
          Flag_Outlier & !is.na(Shrinkage_Error) ~ paste(Trait, "flagged as outlier AND", Shrinkage_Error),
          Flag_Outlier ~ paste(Trait, "flagged for being an extreme outlier"),
          !is.na(Shrinkage_Error) ~ paste(Trait, Shrinkage_Error),
          TRUE ~ paste(Trait, "flagged")
        )
      ) %>%
      group_by(Plot) %>%
      summarise(Validation_record = paste(Error_Msg, collapse = " | "), .groups = "drop")
    
    # 3. Generate "Alive" Column (Binary at Oldest Age) - [FIXED]
    alive_summary <- NULL
    # Get all survival traits present in this dataset
    sur_traits <- unique(final_long_dedup$Trait[str_detect(final_long_dedup$Trait, "(?i)Sur_")])
    
    if(length(sur_traits) > 0) {
      # Create a temporary lookup table: Trait Name vs Numeric Age
      sur_lookup <- tibble(
        Trait_Name = sur_traits,
        Age_Num = as.numeric(str_extract(sur_traits, "\\d+"))
      )
      
      # Find the specific TRAIT NAME corresponding to the max age
      # This avoids the "Sur_8" vs "Sur_08" mismatch
      target_row <- sur_lookup %>% 
        filter(Age_Num == max(Age_Num, na.rm = TRUE)) %>% 
        slice(1)
      
      if(nrow(target_row) > 0) {
        target_sur_trait <- target_row$Trait_Name
        
        # Extract status: 1 = Alive, 0 = Missing/Dead
        alive_summary <- final_long_dedup %>%
          filter(Trait == target_sur_trait) %>%
          select(Plot, Alive = Value_Num) %>%
          distinct()
      }
    }
    
    # 4. Prepare Wide Data (Text)
    data_wide_clean <- final_long_dedup %>% 
      select(Plot, Trait, Value_Char) %>% 
      pivot_wider(names_from = Trait, values_from = Value_Char)
    
    # 5. Prepare Wide Data (Numeric)
    exp_data_wide_numeric <- final_long_dedup %>% 
      select(Plot, Trait, Value_Num) %>% 
      pivot_wider(names_from = Trait, values_from = Value_Num)
    
    # 6. Join Spatial to Numeric (for PDF)
    if (!is.null(spatial_info)) {
      exp_data_wide_numeric <- left_join(exp_data_wide_numeric, spatial_info, by = "Plot")
    }
    
    # 7. Prepare Flags
    flags_wide <- final_long_dedup %>% 
      select(Plot, Trait, Reject_Flag) %>% 
      pivot_wider(names_from = Trait, values_from = Reject_Flag, names_glue = "{Trait}_reject")
    
    # 8. Master Join (Data + Flags + Validation + Alive + Design + Spatial)
    final_wide_with_flags <- left_join(data_wide_clean, flags_wide, by = "Plot")
    
    # Join [NEW] Validation Record
    final_wide_with_flags <- left_join(final_wide_with_flags, validation_summary, by = "Plot")
    
    # Join [NEW] Alive Column
    if(!is.null(alive_summary)) {
      final_wide_with_flags <- left_join(final_wide_with_flags, alive_summary, by = "Plot")
    }
    
    if (!is.null(genetic_info)) final_wide_with_flags <- left_join(final_wide_with_flags, genetic_info, by = "Plot")
    if (!is.null(spatial_info)) final_wide_with_flags <- left_join(final_wide_with_flags, spatial_info, by = "Plot")
    
    # 9. Sorting Columns
    # Define preferred order: Plot, Design info, Alive, Validation, Data...
    gen_cols <- c("Block", "Prow", "Ppos", "Family_name", "Alive", "Validation_record")
    existing_gen_cols <- intersect(names(final_wide_with_flags), gen_cols)
    remaining_cols <- sort(setdiff(names(final_wide_with_flags), c("Plot", existing_gen_cols)))
    
    final_wide_with_flags <- final_wide_with_flags %>% select(all_of(c("Plot", existing_gen_cols, remaining_cols)))
    
    write_csv(final_wide_with_flags, file.path(exp_path, paste0(curr_exp, "_Full_Data_With_Flags.csv")), na = "")
    
    # 10. PDF Report
    current_traits <- unique(final_long_dedup$Trait)
    prefixes_found <- unique(final_long_dedup$Prefix)
    
    try({
      pdf(file.path(exp_path, paste0(curr_exp, "_graphs.pdf")), width = 8, height = 11)
      generate_pdf_report(final_long_dedup, exp_data_wide_numeric, curr_exp, prefixes_found, current_traits)
      dev.off()
    })
    while(dev.cur() > 1) dev.off()
    
    # 11. Stats
    stats_df <- final_long_dedup %>%
      filter(Reject_Flag == 0) %>%
      group_by(Trait) %>%
      summarise(N_Valid = sum(!is.na(Value_Num)), Mean = mean(Value_Num, na.rm=TRUE), 
                Std_Dev = sd(Value_Num, na.rm=TRUE), Min = min(Value_Num, na.rm=TRUE), Max = max(Value_Num, na.rm=TRUE),
                CV_Pct = (sd(Value_Num, na.rm=TRUE)/mean(Value_Num, na.rm=TRUE))*100) %>%
      mutate(across(where(is.numeric), ~ round(., 2)))
    
    write_csv(stats_df, file.path(exp_path, paste0(curr_exp, "_Stats.csv")), na = "")
    
    # ----------------------------------------------------------------------------
    # STEP 6: XML GENERATION (UPDATED WITH StdTraits FIELDS)
    # ----------------------------------------------------------------------------
    xml_data <- data_with_outliers %>%
      group_by(Trait) %>%
      mutate(
        Calc_Vals = if_else(Value_Num > 0, Value_Num, NA_real_),
        Min_Val   = min(Calc_Vals, na.rm = TRUE),
        Max_Val   = max(Calc_Vals, na.rm = TRUE),
        
        Trait_Date = {
          valid_dates <- Date[!is.na(Date)]
          if(length(valid_dates) > 0) as.character(valid_dates[1]) else NA_character_
        }
      ) %>%
      ungroup() %>%
      distinct(Trait, Trait_Orig, Prefix, Age, UnitCode, Is_Ordinal, Min_Val, Max_Val, Trait_Date) %>%
      
      # Join with the trait_map to get Group/Desc/Units from Excel
      left_join(trait_map, by = c("Prefix" = "trait_code_FR")) %>%
      
      mutate(
        # 1. GROUP LOGIC: Use Excel Group -> Fallback to "Health" -> Default
        Group = case_when(
          !is.na(xml_group) & xml_group != "" ~ xml_group,
          str_detect(Prefix, "(?i)SUR") ~ "Health",
          TRUE ~ "Uncategorized"
        ),
        
        # 2. DESCRIPTION LOGIC: Use Excel Desc -> Fallback to Prefix
        Desc_Base = coalesce(xml_desc_base, Prefix),
        
        # 3. UNIT LOGIC: Use Excel Units -> Fallback to Calculation
        Real_Units = case_when(
          !is.na(xml_units) & xml_units != "none" & xml_units != "" ~ xml_units,
          Is_Ordinal & is.finite(Min_Val) ~ paste0("$", Min_Val, "-", Max_Val),
          str_detect(Trait, "(?i)Sur") ~ "0/$1",
          str_detect(Trait, "(?i)Pil") ~ "mm",
          str_detect(Trait, "(?i)Av") ~ "m/s",
          !is.na(UnitCode) & UnitCode != "" ~ str_to_lower(UnitCode),
          TRUE ~ "unitless"
        ),
        
        Date_String = if_else(!is.na(Trait_Date), paste(" Date:", Trait_Date), ""),
        
        # Construct Description: "Tree Height Age 10 (was HGT_10) Date: 2023-01-01"
        Full_Description = paste0(Desc_Base, " Age ", Age, " (was ", Trait_Orig, ")", Date_String),
        
        # Construct XML Tag
        XML_Tag = sprintf(
          '<trait group_name="%s" name="%s" trait_type="M" description="%s" data_type="N" is_solver_mappable="1" summary_rule="first_with_reject" units="%s" validate="none" />',
          Group, Trait, Full_Description, Real_Units
        )
      ) %>%
      arrange(Trait)
    
    xml_filename <- file.path(exp_path, paste0(curr_exp, "_Traits.xml"))
    
    xml_content <- c(
      "<trial>",
      "<traitlist>",
      xml_data$XML_Tag,
      "</traitlist>",
      "</trial>"
    )
    
    writeLines(xml_content, xml_filename)
    
    message("  -> Processed & Saved.")
    
  }, error = function(e) {
    message(paste("  -> ERROR processing", curr_exp, ": ", e$message))
  })
}


###### Part 3: Re-run PDF funtion on DP_ready.csv #######

# 1. SETUP
#    (Ensure "PART 0" helper functions are loaded)
base_path <- here() 
target_trials <- c("Moray_55") # Add your trial names here

for (curr_exp in target_trials) {
  
  message(paste("\nProcessing:", curr_exp))
  
  # --- Step 1: Define File Path ---
  csv_filename <- paste0(curr_exp, "_DP_ready.csv")
  csv_path <- file.path(base_path, curr_exp, csv_filename)
  
  if (!file.exists(csv_path)) {
    message(paste("  -> ERROR: File not found:", csv_path))
    next
  }
  
  # --- Step 2: Load Wide Data ---
  wide_data <- read_csv(csv_path, show_col_types = FALSE)
  
  # --- Step 3: Reshape to Long Data ---
  metadata_cols <- c("Ploc", "Plot", "Validation_record", "Alive", 
                     "Block", "Family_name", "Prow", "Ppos")
  
  # A. Pivot Measurement Values (Excluding Rejects)
  long_values <- wide_data %>%
    select(-any_of(metadata_cols)) %>%
    select(-ends_with("_reject")) %>% # Ensure we don't mix flags into values
    mutate(Plot = wide_data$Plot) %>% 
    pivot_longer(
      cols = -Plot,
      names_to = "Trait",
      values_to = "Value_Num"
    )
  
  # B. Pivot Rejection Flags
  long_flags <- wide_data %>%
    select(Plot, ends_with("_reject")) %>%
    pivot_longer(
      cols = -Plot,
      names_to = "Trait_Reject",
      values_to = "Reject_Status"
    ) %>%
    mutate(Trait = str_remove(Trait_Reject, "_reject")) %>%
    select(Plot, Trait, Reject_Status)
  
  # C. Join and Calculate Metadata
  long_data_final <- long_values %>%
    left_join(long_flags, by = c("Plot", "Trait")) %>%
    mutate(
      Flag_Outlier = coalesce(Reject_Status == 1, FALSE),
      Age = str_extract(Trait, "\\d+"),
      Prefix = str_to_upper(str_extract(Trait, "^[a-zA-Z]+")),
      Is_Survival = str_detect(Trait, "(?i)Sur_")
    ) %>%
    group_by(Trait) %>%
    mutate(
      n_vals = n_distinct(Value_Num, na.rm = TRUE),
      Is_Ordinal = !Is_Survival & n_vals < 9
    ) %>%
    ungroup() %>%
    filter(!is.na(Value_Num)) 
  
  # --- Step 4: Prepare Helper Lists ---
  all_traits <- unique(long_data_final$Trait)
  multi_age_prefixes <- unique(long_data_final$Prefix)
  
  # --- Step 5: Generate PDF (WITH FIX) ---
  pdf_file <- file.path(base_path, curr_exp, paste0(curr_exp, "_Report_Reprocessed.pdf"))
  
  message("  -> Generating PDF...")
  
  tryCatch({
    pdf(pdf_file, width = 8, height = 11)
    
    generate_pdf_report(
      long_data = long_data_final,
      
      # [FIX IS HERE] 
      # We filter out "_reject" columns so the Zombie Check 
      # doesn't see "Sur_01" and "Sur_01_reject" as two different ages.
      wide_data = wide_data %>% select(-ends_with("_reject")), 
      
      exp_name = curr_exp,
      multi_age_prefixes = multi_age_prefixes,
      all_traits = all_traits
    )
    
    dev.off()
    message("  -> Success! PDF saved.")
    
  }, error = function(e) {
    message(paste("  -> ERROR generating PDF:", e$message))
    while(dev.cur() > 1) dev.off() 
  })
}
