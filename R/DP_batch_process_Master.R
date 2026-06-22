  
  
# # # # # # # # # # # # # # # # # # # # # # # 
# PART 0: HELPER FUNCTIONS #### 
# # # # # # # # # # # # # # # # # # # # # # # 
# --- 1a. Parse Design File (TXT / No Extension) ---
parse_long_design_file <- function(filepath, exp_prefix, spp_code) {
  raw_lines <- readLines(filepath)
  trt_start <- grep("TREATMENTS", raw_lines)
  plot_start <- grep("PLOTS", raw_lines)
  if(length(trt_start) == 0 || length(plot_start) == 0) return(NULL)
  
  spp_regex <- paste0("^\\s*", spp_code, "\\s*(\\d+)\\s*", spp_code, "\\s*(.*)$")
  control_regex <- paste0("^\\*?\\s*", spp_code, "\\s*")
  prefix_low <- tolower(spp_code)
  
  trt_lines <- raw_lines[(trt_start + 1):(plot_start - 1)]
  trt_lines <- trt_lines[grep("=", trt_lines)]
  
  trt_df <- tibble(line = trt_lines) %>%
    mutate(Is_Control = str_detect(str_trim(line), "^\\*")) %>%
    extract(line, into = c("Design_ID", "Cross_Name"), regex = "^\\s*\\*?\\s*(\\d+)=(.*)$", convert = TRUE) %>%
    mutate(Cross_Name = str_trim(Cross_Name)) %>%
    extract(Cross_Name, into = c("Maternal_ID", "Paternal_ID"), regex = spp_regex, remove = FALSE) %>%
    mutate(
      Maternal_ID = str_trim(Maternal_ID), Paternal_ID = str_trim(Paternal_ID),
      Control_Name = if_else(Is_Control, str_trim(str_replace(Cross_Name, control_regex, "")), NA_character_)
    )
  
  plot_lines <- raw_lines[(plot_start + 1):length(raw_lines)]
  plot_lines <- plot_lines[grep("^[0-9]+,[0-9]+,[0-9]+", plot_lines)]
  plot_df <- tibble(line = plot_lines) %>%
    extract(line, into = c("Plot", "Block", "SubBlock", "Design_ID"), regex = "^(\\d+),(\\d+),(\\d+)\\s+(\\d+)", convert = TRUE)
  
  final_design <- plot_df %>% left_join(trt_df, by = "Design_ID") %>%
    mutate(
      Plot = as.numeric(Plot),
      Family_name = case_when(
        Design_ID == 0 ~ paste0(exp_prefix, "_Filler"),
        !is.na(Control_Name) ~ paste0(prefix_low, Control_Name),
        !is.na(Paternal_ID) & str_detect(Paternal_ID, "(?i)OP") ~ paste0(prefix_low, Maternal_ID, "_", Paternal_ID),        !is.na(Maternal_ID) & !is.na(Paternal_ID) ~ paste0(prefix_low, Maternal_ID, "_", prefix_low, Paternal_ID),
        TRUE ~ str_replace_all(Cross_Name, "\\s+", "") %>% str_replace(paste0("^", spp_code), paste0(prefix_low, "_"))
      )
    ) %>% select(Plot, Block, SubBlock, Family_name)
  return(final_design)
}

# --- 1b. Parse Design File (XLSX) ---
parse_xlsx_design_file <- function(filepath, exp_prefix, spp_code) {
  raw_design <- read_excel(filepath, col_types = "text")
  if("Rep" %in% names(raw_design) && !"Block" %in% names(raw_design)) raw_design <- raw_design %>% rename(Block = Rep)
  
  spp_regex <- paste0("^(?i)\\s*", spp_code, "\\s*(\\d+)\\s*", spp_code, "\\s*(.*)$")
  control_regex <- paste0("^(?i)\\s*", spp_code, "\\s*") 
  prefix_low <- tolower(spp_code)
  
  clean_design <- raw_design %>% select(any_of(c("Plot", "Block", "Seedlot", "SubBlock"))) %>% filter(!is.na(Plot)) %>%
    mutate(
      Seedlot = str_replace_all(Seedlot, "[[:cntrl:]]", ""), Seedlot = if_else(str_trim(Seedlot) == "", NA_character_, Seedlot),
      Is_Control = str_detect(str_trim(Seedlot), "^\\*"), Clean_Seedlot = str_replace(Seedlot, "^[^=]*=\\s*", "") 
    ) %>%
    extract(Clean_Seedlot, into = c("Maternal_ID", "Paternal_ID"), regex = spp_regex, remove = FALSE) %>%
    mutate(
      Plot = suppressWarnings(as.numeric(Plot)), Maternal_ID = str_trim(Maternal_ID), Paternal_ID = str_trim(Paternal_ID),
      Control_Name = if_else(Is_Control, str_trim(str_replace(Clean_Seedlot, control_regex, "")), NA_character_),
      Family_name = case_when(
        !is.na(Control_Name) ~ paste0(prefix_low, Control_Name),
        !is.na(Paternal_ID) & str_detect(Paternal_ID, "(?i)OP") ~ paste0(prefix_low, Maternal_ID, "_OPCB"),
        !is.na(Maternal_ID) & !is.na(Paternal_ID) ~ paste0(prefix_low, Maternal_ID, "_", prefix_low, Paternal_ID),
        is.na(Seedlot) ~ paste0(exp_prefix, "_Filler"), TRUE ~ paste0(exp_prefix, "_Filler")
      )
    )
  if(!"SubBlock" %in% names(clean_design)) clean_design$SubBlock <- NA_character_
  clean_design <- clean_design %>% select(Plot, Block, SubBlock, Family_name)
  return(clean_design)
}
  
# --- 1c. Parse CEDD ASCII File (txt) ---
parse_cedd_raw_txt <- function(filepath, trait_name, trait_prefix_raw, txt_age) {
    
    # 1. Read the entire file
    raw_text <- readChar(filepath, file.info(filepath)$size)
    
    # 2. Extract Data Blocks
    block_pattern <- "(?m)^\\s*(\\d+):\\s*(.*?)\\s*\\$"
    matches <- str_match_all(raw_text, block_pattern)[[1]]
    
    if (nrow(matches) == 0) return(tibble())
    
    # 3. UNIT EXTRACTION LOGIC
    # Read just the first 10 lines to hunt for the CEDD header code (e.g., HTCM03)
    head_lines <- readLines(filepath, n = 10, warn = FALSE)
    header_block <- paste(head_lines, collapse = " ")
    
    # Build a dynamic regex: Look for the Prefix, capture the middle letters, ending with Age
    unit_regex <- paste0("(?i)\\b", trait_prefix_raw, "([A-Za-z]+)", txt_age, "\\b")
    extracted_unit <- str_match(header_block, unit_regex)[, 2]
    
    # Clean up the unit (default to "" if nothing is found)
    final_unit <- if (!is.na(extracted_unit)) str_to_lower(extracted_unit) else ""
    
    # 4. Process each block
    plot_data_list <- list()
    for (i in 1:nrow(matches)) {
      plot_id <- as.character(matches[i, 2])
      raw_values_str <- matches[i, 3]
      
      clean_str <- str_squish(str_replace_all(raw_values_str, "[\\r\\n\\t]+", " "))
      val_tokens <- str_split(clean_str, " ")[[1]]
      val_tokens <- val_tokens[val_tokens != ""] 
      
      if (length(val_tokens) > 0) {
        plot_df <- tibble(
          Plot = plot_id,
          Tree = 1:length(val_tokens), 
          Raw_Val = val_tokens
        ) %>%
          mutate(
            Final_Val = if_else(Raw_Val == "M", NA_real_, suppressWarnings(as.numeric(Raw_Val))),
            Extracted_Unit = final_unit # Attach the unit we found!
          ) %>%
          select(Plot, Tree, Extracted_Unit, !!sym(trait_name) := Final_Val)
        
        plot_data_list[[i]] <- plot_df
      }
    }
    
    parsed_df <- bind_rows(plot_data_list)
    return(parsed_df)
  }
  
# --- 2. PDF Reporting Function ---
generate_pdf_report <- function(long_data, wide_data, exp_name, multi_age_prefixes, all_traits) {
    
    all_cont_traits <- long_data %>% filter(Is_Ordinal == FALSE, !str_detect(Trait, "(?i)Sur_")) %>% distinct(Trait) %>% pull(Trait)
    pil_traits  <- all_cont_traits[str_detect(all_cont_traits, "(?i)Pil")]
    cont_traits <- setdiff(all_cont_traits, pil_traits)
    ord_traits  <- long_data %>% filter(Is_Ordinal == TRUE, !str_detect(Trait, "(?i)Sur_")) %>% distinct(Trait) %>% pull(Trait)
    
    # Distributions (Standard Continuous)
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
    
    # Distributions (Pilodyn)
    plot_data_pil <- long_data %>% filter(Trait %in% pil_traits, !is.na(Value_Num))
    if (nrow(plot_data_pil) > 0) {
      p1_pil <- ggplot(plot_data_pil, aes(x = Value_Num, fill = as.character(Flag_Outlier))) +
        geom_histogram(binwidth = 1, color = "black", alpha = 0.7, center = 0) +
        facet_wrap(~Trait, scales = "free") +
        scale_fill_manual(values = c("FALSE" = "lightblue", "TRUE" = "orange"), name = "Is Outlier?") +
        labs(title = paste(exp_name, "- Distributions (Pilodyn)"), subtitle = "Integer-based (binwidth=1)", x = "Value") + 
        theme_minimal() + theme(legend.position = "bottom")
      print(p1_pil)
    }
    
    # Distributions (Ordinal)
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
    
    # Shrinkage Checks
    if (length(all_traits) > 0) {
      base_traits <- unique(str_remove(all_traits[str_detect(all_traits, "_\\d+$")], "_\\d+$"))
      for (base_t in base_traits) {
        if (str_detect(base_t, "(?i)(Sur|Resi|Detrended|Ampli|Resistance|Resist|Crack|Form|Br|St|Pil)")) next      
        regex_pattern <- paste0("^", base_t, "_\\d+$")
        cols <- str_sort(all_traits[str_detect(all_traits, regex_pattern)], numeric = TRUE)
        
        if (length(cols) < 2) next
        for (i in 1:(length(cols) - 1)) {
          younger <- cols[i]; older <- cols[i+1]
          if(!all(c(younger, older) %in% names(wide_data))) next
          
          unit_younger <- long_data %>% filter(Trait == younger) %>% drop_na(UnitCode) %>% pull(UnitCode) %>% head(1)
          unit_older <- long_data %>% filter(Trait == older) %>% drop_na(UnitCode) %>% pull(UnitCode) %>% head(1)
          if(length(unit_younger) > 0 && length(unit_older) > 0 && unit_younger != unit_older) next
          
          plot_df <- wide_data %>% select(Plot, Tree, X = all_of(older), Y = all_of(younger)) %>%
            filter(!is.na(X), !is.na(Y)) %>%
            mutate(Is_Shrinkage = Y > X, Status = if_else(Is_Shrinkage, "Shrinkage", "Normal"))
          
          if(nrow(plot_df) == 0) next
          p_shrink <- ggplot(plot_df, aes(x = X, y = Y)) +
            geom_abline(slope = 1, intercept = 0, color = "gray50", linetype = "dashed") +
            geom_point(aes(color = Status), size = 2, alpha = 0.6) +
            scale_color_manual(values = c("Normal"="black", "Shrinkage"="red")) +
            labs(title = "Shrinkage Check", subtitle = paste(younger, "vs", older, "(Units Match)"), x = older, y = younger) +
            theme_minimal()
          print(p_shrink)
        }
      }
    }
    
    # Concentric Temporal Check
    if (all(c("Prow", "Ppos") %in% names(wide_data))) {
      all_units <- wide_data %>% select(Plot, Tree) %>% distinct()
      physical_data <- long_data %>% mutate(Age_Num = suppressWarnings(as.numeric(as.character(Age)))) %>%
        filter(!str_detect(Trait, "(?i)Sur_"), !is.na(Value_Num), !is.na(Age_Num), Age_Num > 0)
      
      plot_ages <- sort(unique(physical_data$Age_Num), decreasing = TRUE)
      
      if (length(plot_ages) > 0) {
        age_label <- paste(plot_ages, collapse = ", ")
        counts_per_tree <- physical_data %>% group_by(Plot, Tree, Age_Num) %>% summarise(Count = n(), .groups = "drop")
        full_grid <- expand_grid(all_units, Age_Num = plot_ages) %>% left_join(counts_per_tree, by = c("Plot", "Tree", "Age_Num")) %>% replace_na(list(Count = 0))
        mode_per_age <- full_grid %>% filter(Count > 0) %>% group_by(Age_Num) %>% summarise(Standard_Count = as.numeric(names(sort(table(Count), decreasing=TRUE)[1])), .groups = "drop")
        
        plot_data <- full_grid %>% left_join(mode_per_age, by = "Age_Num") %>%
          mutate(Standard_Count = replace_na(Standard_Count, 1), Percent = (Count / Standard_Count) * 100) %>%
          left_join(wide_data %>% select(Plot, Tree, Prow, Ppos) %>% distinct(), by = c("Plot", "Tree")) %>%
          mutate(Prow = suppressWarnings(as.numeric(Prow)), Ppos = suppressWarnings(as.numeric(Ppos))) %>%
          filter(!is.na(Prow), !is.na(Ppos))
        
        if(nrow(plot_data) > 0) {
          min_row <- min(plot_data$Prow, na.rm = TRUE); max_row <- max(plot_data$Prow, na.rm = TRUE)
          min_pos <- min(plot_data$Ppos, na.rm = TRUE); max_pos <- max(plot_data$Ppos, na.rm = TRUE)
          sizes <- seq(0.95, 0.25, length.out = length(plot_ages))
          
          p_concentric <- ggplot() +
            scale_y_reverse(breaks = seq(min_row, max_row, by = 5)) +
            scale_x_continuous(breaks = seq(min_pos, max_pos, by = 5)) + coord_fixed() + theme_minimal() +
            theme(axis.text.x = element_text(angle = 90, vjust = 0.5, size = 8), axis.text.y = element_text(size = 8)) +
            labs(title = paste(exp_name, "- Data Completeness"), subtitle = paste0("Ages Plotted: ", age_label, "\nOuter Ring = Oldest -> Inner Square = Youngest"), x = "Position", y = "Row", fill = "% Data") +
            scale_fill_gradient2(low = "red", mid = "yellow", high = "forestgreen", midpoint = 50, limits = c(0, 100), oob = scales::squish)
          
          for (i in seq_along(plot_ages)) {
            curr_age <- plot_ages[i]; curr_size <- sizes[i]
            layer_data <- plot_data %>% filter(Age_Num == curr_age)
            p_concentric <- p_concentric + geom_tile(data = layer_data, aes(x = Ppos, y = Prow, fill = Percent), width = curr_size, height = curr_size, color = "black", size = 0.1) 
          }
          print(p_concentric)
        }
      }
    }
    
    # Resurrection Check (Zombie) - Text Based
    sur_cols <- sort(names(wide_data)[str_detect(names(wide_data), "(?i)Sur_\\d+")])
    if (length(sur_cols) >= 2) {
      sur_ages <- as.numeric(str_extract(sur_cols, "\\d+"))
      sur_map  <- data.frame(Col = sur_cols, Age = sur_ages) %>% arrange(Age)
      
      for (i in 1:(nrow(sur_map) - 1)) {
        younger_col <- sur_map$Col[i]; older_col <- sur_map$Col[i+1]
        zombies <- wide_data %>% filter(.data[[younger_col]] == 0 & .data[[older_col]] == 1) %>% select(Plot, Tree) %>% distinct() %>% arrange(Plot, Tree)
        
        if (nrow(zombies) > 0) {
          bad_ids <- paste(zombies$Plot, zombies$Tree, sep="-")
          id_string <- paste(bad_ids, collapse = ", ")
          wrapped_text <- str_wrap(id_string, width = 80)
          final_text <- paste0("RESURRECTION CHECK (Zombie Trees)\nLogic: Marked Dead (0) at ", younger_col, " -> Alive (1) at ", older_col, "\nTotal Flagged: ", length(bad_ids), "\n\nPlot-Tree IDs:\n", wrapped_text)
          t_grob <- textGrob(final_text, x = 0.5, y = 0.5, just = "center", gp = gpar(fontsize = 10, fontface = "plain"))
          grid.arrange(t_grob)
        }
      }
    }
    
    # Correlations
    if (length(all_cont_traits) >= 2) {
      gen_pairs <- combn(all_cont_traits, 2, simplify = FALSE)
      for (pair in gen_pairs) {
        t1 <- pair[1]; t2 <- pair[2]
        if(!all(c(t1, t2) %in% names(wide_data))) next
        
        outlier_plots <- long_data %>% filter(Trait %in% c(t1, t2), Flag_Outlier == TRUE) %>% select(Plot, Tree) %>% distinct() %>% mutate(Status = "Outlier")
        
        # FIX: Added X_Val != 0 and Y_Val != 0 to explicitly drop erroneous zeros from the math and the plot
        plot_df <- wide_data %>% select(Plot, Tree, X_Val = all_of(t1), Y_Val = all_of(t2)) %>%
          filter(!is.na(X_Val), !is.na(Y_Val), X_Val != 0, Y_Val != 0) %>% 
          left_join(outlier_plots, by = c("Plot", "Tree")) %>% mutate(Status = replace_na(Status, "Normal"))
        
        if(nrow(plot_df) < 3) next
        
        # Calculate the correlation coefficient
        cor_val <- round(cor(plot_df$X_Val, plot_df$Y_Val, method = "spearman"), 3)
        
        p_corr <- ggplot(plot_df, aes(x=X_Val, y=Y_Val)) + 
          geom_smooth(method = "lm",formula = y ~ x, se = FALSE, color = "gray", linetype = "dashed", alpha=0.5) +
          geom_point(aes(color = Status), alpha = 0.6) + 
          scale_color_manual(values = c("Normal" = "black", "Outlier" = "orange")) +
          labs(
            title = "Correlation Check", 
            subtitle = paste("Spearman Rho:", cor_val, "| N:", nrow(plot_df)), # Added to print on the PDF
            x = t1, 
            y = t2
          ) + 
          theme_minimal()
        print(p_corr)
      }
    }
    
    # Heatmaps
    if (all(c("Prow", "Ppos") %in% names(wide_data))) {
      heatmap_traits <- c(cont_traits, ord_traits, pil_traits)
      for (trait in heatmap_traits) {
        if (!trait %in% names(wide_data)) next
        map_data <- wide_data %>% select(Prow, Ppos, Value = all_of(trait)) %>% mutate(Prow = as.numeric(Prow), Ppos = as.numeric(Ppos)) %>% filter(!is.na(Value), !is.na(Prow), !is.na(Ppos))
        if (nrow(map_data) == 0) next
        p_map <- ggplot(map_data, aes(x = Ppos, y = Prow, fill = Value)) + geom_tile(color = "white", size = 0.1) + scale_fill_distiller(palette = "Spectral", direction = -1) + scale_y_reverse() + labs(title = paste(exp_name, "- Spatial Map:", trait)) + theme_minimal() + coord_fixed()
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
        plot <- as.character(matrix_data[i, j])
        if (is.na(plot) || str_trim(plot) == "0" || str_trim(plot) == "") next
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
  
# --- 4. Map interior tree sequence to full plot sequence ---
map_interior_trees <- function(tree_idx, subset_size, full_size) {
  
  # ADDED: subset_size == 1 bypass. 
  # If max tree ID is 1, it's a plot-wide summary, not a 1x1 interior grid.
  if (is.na(subset_size) || is.na(full_size) || subset_size == full_size || subset_size == 1) return(tree_idx)
  
  n <- sqrt(full_size); m <- sqrt(subset_size)
  if (n %% 1 != 0 || m %% 1 != 0) return(tree_idx)
  
  offset <- (n - m) / 2
  row_m <- ceiling(tree_idx / m); col_m <- (tree_idx - 1) %% m + 1
  row_n <- row_m + offset; col_n <- col_m + offset
  mapped_idx <- (row_n - 1) * n + col_n
  
  return(mapped_idx)
}
  
# --- 5. Traversal/ Data orientation correction ---
get_grid_layout <- function(n_rows, n_cols, interior_only = FALSE) {
    if (n_rows == 1 && n_cols == 8) return(matrix(1:8, nrow=1))
    if (n_rows == 8 && n_cols == 1) return(matrix(1:8, ncol=1))
    
    mat <- matrix(1:(n_rows*n_cols), nrow=n_rows, byrow=TRUE)
    if(interior_only && n_rows >= 3 && n_cols >= 3) mat <- mat[2:(n_rows-1), 2:(n_cols-1)]
    return(mat)
  }

get_traversal_path <- function(layout_mat, start_corner="top_left", direction="horizontal", snake=FALSE) {
  if (length(layout_mat) == 8) {
    if (start_corner %in% c("reversed", "bottom_right")) return(rev(as.vector(layout_mat)))
    return(as.vector(layout_mat))
  }
  n_rows <- nrow(layout_mat); n_cols <- ncol(layout_mat)
  row_indices <- if (start_corner %in% c("bottom_left", "bottom_right")) rev(seq_len(n_rows)) else seq_len(n_rows)
  col_indices <- if (start_corner %in% c("top_right", "bottom_right")) rev(seq_len(n_cols)) else seq_len(n_cols)
  
  path <- c()
  if (direction == "horizontal") {
    for (i_idx in seq_along(row_indices)) {
      i <- row_indices[i_idx]; cols <- col_indices
      if (snake && i_idx %% 2 == 0) cols <- rev(cols)
      for (j in cols) path <- c(path, layout_mat[i, j])
    }
  } else {
    for (j_idx in seq_along(col_indices)) {
      j <- col_indices[j_idx]; rows <- row_indices
      if (snake && j_idx %% 2 == 0) rows <- rev(rows)
      for (i in rows) path <- c(path, layout_mat[i, j])
    }
  }
  return(path)
}

# # # # # # # # # # # # # # # # # # # # # # # 
# PART 1: LOAD TRANSLATION MAP & FOLDERS ####
# # # # # # # # # # # # # # # # # # # # # # # 

run_dataplan_pipeline <- function(base_dir, trial_series, traits_file, plot_type,target_trials=NULL) {
  
  root_data_dir <- file.path(base_dir, trial_series)
  
  if (file.exists(traits_file)) {
    trait_map <- read_csv(traits_file, show_col_types = FALSE) %>%
      select(trait_code_FR = `Eng Desc`, trait_code_DP = Trait, xml_group = Group, xml_desc_base = Description, xml_units = Units) %>%
      mutate(across(everything(), as.character)) %>%
      filter(!is.na(trait_code_FR), !is.na(trait_code_DP)) %>%
      arrange(desc(nchar(trait_code_FR)))
  } else {
    stop(paste("Error: Trait translation file not found at:", traits_file))
  }
  
  all_dirs <- list.dirs(path = root_data_dir, recursive = FALSE, full.names = FALSE)
  experiments_to_process <- setdiff(all_dirs, c("00_Scripts", "Archive", ".git", ".Rproj.user"))
  message(paste("Found", length(experiments_to_process), "folders to check."))
  
  # --- TARGETED RUN LOGIC ---
  if (!is.null(target_trials)) {
    valid_targets <- intersect(target_trials, experiments_to_process)
    if (length(valid_targets) == 0) {
      stop("Error: None of the specified target_trials were found in the directory.")
    }
    experiments_to_process <- valid_targets
    message(paste("Targeted Run Mode: Only processing", length(experiments_to_process), "specified trial(s)."))
  } else {
    message(paste("Found", length(experiments_to_process), "folders to check."))
  }
  # # # # # # # # # # # # # # # # # # # # # # # 
  # PART 2: MAIN PROCESSING LOOP ####
  # # # # # # # # # # # # # # # # # # # # # # # 
  
  # Initialize a list to hold correlations for the master summary
  master_correlations_list <- list()
  
  for (curr_exp in experiments_to_process) {
    
    tryCatch({
      message(paste("\n=========================================="))  
      message(paste("Processing Folder:", curr_exp, "| Mode:", plot_type))
      
      exp_path <- file.path(root_data_dir, curr_exp)
      file_prefix <- str_replace_all(curr_exp, " ", "_")
      
      genetic_info <- NULL
      spatial_info <- NULL
      exp_data_long <- NULL
      
      # --- 1. Identify Raw ASCII File ---
      ascii_files <- fs::dir_ls(exp_path, regexp = paste0("(?i)", file_prefix, ".*_ASCII\\.(csv|xlsx)$"))
      
      if (length(ascii_files) == 0) {
        message(paste("  -> SKIPPING: Measurement file not found for:", curr_exp))
        next 
      }
      meas_path <- ascii_files[1] 
      if (length(ascii_files) > 1) message(paste("  -> WARNING: Multiple ASCII files. Loading:", basename(meas_path)))
      
      # --- 2. Dynamically Extract Species Code ---
      # Smartly read the first 5 rows depending on file type
      if (str_detect(meas_path, "(?i)\\.csv$")) {
        temp_data <- read_csv(meas_path, n_max = 5, show_col_types = FALSE) %>% clean_names()
      } else {
        temp_data <- read_excel(meas_path, n_max = 5) %>% clean_names()
      }
      
      # Look for the snake_case versions since clean_names() was applied
      spp_col <- intersect(names(temp_data), c("species", "species_code"))[1]
      
      if (!is.na(spp_col)) {
        current_spp <- toupper(as.character(temp_data[[spp_col]][1]))
      } else {
        current_spp <- "SS" # Fallback
        message("  -> WARNING: No species column found in ASCII, defaulting to: ", current_spp)
      }
      
      # --- 3. Load Design ---
      design_files <- fs::dir_ls(exp_path, regexp = "(?i)_DF(\\.txt|\\.xlsx|\\.)?$")
      if (length(design_files) > 0) {
        design_path <- if(any(str_detect(design_files, "(?i)\\.xlsx$"))) design_files[str_detect(design_files, "(?i)\\.xlsx$")][1] else design_files[1]
        message(paste("  -> Loading Design:", basename(design_path)))
        
        if (str_detect(design_path, "(?i)\\.xlsx$")) {
          genetic_info <- parse_xlsx_design_file(design_path, file_prefix, current_spp) 
        } else {
          genetic_info <- parse_long_design_file(design_path, file_prefix, current_spp) 
        }
        
        # Robust Block Fixing
        if(!is.null(genetic_info)) {
          valid_data <- genetic_info %>% filter(Block != "0") %>% mutate(Plot_Num = as.numeric(Plot), Block_Num = as.numeric(Block))
          if(nrow(valid_data) > 0) {
            block_starts <- valid_data %>% group_by(Block_Num) %>% summarise(Start_Plot = min(Plot_Num), .groups = "drop") %>% arrange(Block_Num)
            block_size <- if(nrow(block_starts) > 1) as.numeric(names(sort(table(diff(block_starts$Start_Plot)), decreasing = TRUE)[1])) else max(valid_data$Plot_Num)
            genetic_info <- genetic_info %>% mutate(Block = if_else(Block == "0", as.character(ceiling(Plot / block_size)), as.character(Block)))
          }
        }
      } else {
        message("  -> NO DESIGN FILE FOUND. Proceeding without genetic info.")
      }
      
      # --- 4. Load Spatial Matrix ---
      matrix_files <- dir_ls(exp_path, regexp = "(?i)_(Matrix\\.csv|layout\\.xlsx)$")
      if (length(matrix_files) > 0) {
        matrix_path <- matrix_files[1]
        
        # 2. Fork the reader based on the file extension
        if (str_detect(matrix_path, "(?i)\\.xlsx$")) {
          # Read Excel (requires readxl library, which is already loaded in your script)
          raw_matrix <- readxl::read_excel(matrix_path, col_names = FALSE, col_types = "text") %>% as.data.frame()
        } else {
          # Read CSV
          raw_matrix <- read.csv(matrix_path, header = FALSE, colClasses = "character")
        }
      
        raw_matrix[is.na(raw_matrix)] <- "0"
        spatial_info <- prowppos(raw_matrix)
        
        if (!is.null(spatial_info)) {
          spatial_info <- spatial_info %>% 
            mutate(
              Plot = as.character(Plot), 
              Tree = suppressWarnings(as.numeric(Tree)), 
              Prow = suppressWarnings(as.numeric(Prow)), 
              Ppos = suppressWarnings(as.numeric(Ppos))
            ) %>% 
            filter(!is.na(Plot))
          
          # 1. Identify any plots containing the word "Filler"
          filler_plots <- spatial_info %>% 
            filter(str_detect(Plot, "(?i)Filler")) %>% 
            distinct(Plot)
          
          if (nrow(filler_plots) > 0) {
            # 2. Find the highest existing numeric plot number (from design or spatial)
            existing_plots <- c(
              suppressWarnings(as.numeric(spatial_info$Plot)),
              if (!is.null(genetic_info)) suppressWarnings(as.numeric(genetic_info$Plot)) else c()
            )
            max_plot <- max(c(0, existing_plots), na.rm = TRUE)
            
            # 3. Create a mapping table: Filler -> New Numeric ID
            filler_mapping <- filler_plots %>%
              mutate(New_Plot = as.character(max_plot + row_number()))
            
            # 4. Update spatial_info with the new numeric IDs
            spatial_info <- spatial_info %>%
              left_join(filler_mapping, by = "Plot") %>%
              mutate(Plot = if_else(!is.na(New_Plot), New_Plot, Plot)) %>%
              select(Plot, Tree, Prow, Ppos)
            
            # 5. Build the genetic info additions using the NEW plot IDs
            filler_addition <- filler_mapping %>%
              mutate(
                Plot = New_Plot, # OVERWRITE the old text "Filler" with the new numeric ID
                Block = "0", 
                SubBlock = NA_character_, 
                Family_name = paste0(file_prefix, "_Filler")
              ) %>%
              select(Plot, Block, SubBlock, Family_name)
            
            # 6. Append to genetic_info
            if (is.null(genetic_info)) {
              genetic_info <- filler_addition
            } else {
              genetic_info <- bind_rows(
                genetic_info %>% mutate(Plot = as.character(Plot)), 
                filler_addition
              )
            }
          } else {
            spatial_info <- spatial_info %>% select(Plot, Tree, Prow, Ppos)
          }
        }
      } else {
        message("  -> WARNING: No spatial matrix file found. Proceeding without coords.")
      }
      
      
      # --- 5. Measurements & Renaming ---
      if (str_detect(meas_path, "(?i)\\.csv$")) raw_data <- read_csv(meas_path, col_types = cols(.default = "c")) %>% clean_names() else raw_data <- read_excel(meas_path, col_types = "text") %>% clean_names()
      
      # --- DYNAMIC COLUMN RENAMING ---
      # 1. Measurement
      meas_col <- intersect(c("measurment", "measurement", "value", "value_char"), names(raw_data))[1]
      if (!is.na(meas_col)) raw_data <- raw_data %>% rename(Measurement = !!sym(meas_col))
      
      # 2. Assessment Trait
      assess_col <- intersect(c("assessment", "assessment_type_long", "assessment_type"), names(raw_data))[1]
      if (!is.na(assess_col)) raw_data <- raw_data %>% rename(Assessment = !!sym(assess_col))
      
      # 3. Assessment Year
      year_col <- intersect(c("assessment_year", "assessment_year_long", "age"), names(raw_data))[1]
      if (!is.na(year_col)) raw_data <- raw_data %>% rename(`Assessment Year` = !!sym(year_col))
      
      # 4. Assessment Unit
      unit_col <- intersect(c("unit", "assessment_unit_long", "assessment_unit"), names(raw_data))[1]
      if (!is.na(unit_col)) raw_data <- raw_data %>% rename(Unit = !!sym(unit_col)) else raw_data <- raw_data %>% mutate(Unit = NA_character_)
      
      # 5. Remarks
      remark_col <- intersect(c("remarks", "comment", "comments"), names(raw_data))[1]
      if (!is.na(remark_col)) raw_data <- raw_data %>% rename(Remarks = !!sym(remark_col)) else raw_data <- raw_data %>% mutate(Remarks = NA_character_)
      
      # 6. Plot
      if("plot" %in% names(raw_data)) raw_data <- raw_data %>% rename(Plot = plot)
      
      # 
      # THE MASTER "TREE ID" INJECTION 
      # 
      if (plot_type == "SINGLE") {
        # Single-Tree: Force every record to be Tree 1
        raw_data <- raw_data %>% mutate(InferredTreePosition = 1)
      } else if (plot_type == "MULTI") {
        if ("inferred_tree_position" %in% names(raw_data)) {
          raw_data <- raw_data %>% rename(InferredTreePosition = inferred_tree_position)
        } else {
          message("    -> Generating sequential Tree IDs for Multi-Tree Plot data...")
          raw_data <- raw_data %>% 
            group_by(Assessment, `Assessment Year`, Plot) %>% 
            mutate(InferredTreePosition = row_number()) %>% 
            ungroup()
        }
      }
      
      # Strip out PowerBI footers and blank rows BEFORE checking for duplicates!
      raw_data <- raw_data %>% filter(!is.na(Plot), !is.na(Measurement))
      
      # Duplicate Checks
      duplicate_check <- raw_data %>% group_by(Plot, InferredTreePosition, Assessment, `Assessment Year`) %>% mutate(Measure_Count = n()) %>% ungroup()
      if (any(duplicate_check$Measure_Count > 1)) {
        message("    -> WARNING: Resolving repeat measurements by prioritizing 'Direction SE'.")
        raw_data <- duplicate_check %>% group_by(Plot, InferredTreePosition, Assessment, `Assessment Year`) %>% arrange(desc(str_detect(replace_na(Remarks, ""), "(?i)Direction SE"))) %>% slice(1) %>% ungroup() %>% select(-Measure_Count)
      } else {
        raw_data <- duplicate_check %>% select(-Measure_Count)
      }
      
      # Base Long Format
      exp_data_long <- raw_data %>%
        select(Plot_Orig = Plot, Tree_Orig = InferredTreePosition, Value_Char = Measurement, TraitCode = Assessment, UnitCode = Unit, Age = `Assessment Year`) %>%
        filter(!is.na(Plot_Orig)) %>%
        mutate(
          Plot = as.character(Plot_Orig),
          Tree_Orig_Num = suppressWarnings(as.numeric(Tree_Orig)), 
          Value_Num = suppressWarnings(as.numeric(Value_Char)),
          Prefix = str_to_upper(str_trim(replace_na(TraitCode, ""))),
          Age = str_pad(as.character(Age), width = 2, pad = "0"), # <--- FORCES LEADING ZEROS
          Date = as.Date(NA), Trait_Orig = paste0(Prefix, replace_na(UnitCode, ""), Age)
        )
      
      # Map Trees (Bypass for Single-Tree)
      if(plot_type == "MULTI") {
        exp_data_long <- exp_data_long %>%
          group_by(Plot) %>% mutate(Max_Trees_Planted = max(Tree_Orig_Num, na.rm = TRUE)) %>% 
          group_by(Plot, Trait_Orig) %>% mutate(Trees_in_Subset = max(Tree_Orig_Num, na.rm = TRUE), Tree = purrr::pmap_dbl(list(Tree_Orig_Num, Trees_in_Subset, Max_Trees_Planted), map_interior_trees)) %>%
          ungroup() %>% select(-Tree_Orig_Num, -Max_Trees_Planted, -Trees_in_Subset)
      } else {
        exp_data_long <- exp_data_long %>% mutate(Tree = Tree_Orig_Num) %>% select(-Tree_Orig_Num)
      }
      
      # Trait Map Renaming
      unique_prefixes <- exp_data_long %>% distinct(Prefix)
      unique_prefixes$trait_code_DP <- map_chr(unique_prefixes$Prefix, function(p) {
        if(is.na(p) || p == "") return(NA_character_)
        match <- trait_map %>% filter(str_starts(p, trait_code_FR)) %>% slice(1) %>% pull(trait_code_DP)
        if(length(match) == 0) return(NA_character_) else return(match)
      })
      
      exp_data_long <- exp_data_long %>% left_join(unique_prefixes, by = "Prefix") %>%
        mutate(Trait = case_when(!is.na(trait_code_DP) & !is.na(Age) ~ paste0(trait_code_DP, "_", Age), !is.na(trait_code_DP) & is.na(Age) ~ trait_code_DP, TRUE ~ Trait_Orig)) %>%
        select(Plot, Tree, Trait, Value_Char, Value_Num, UnitCode, Trait_Orig, Age, Prefix, Date) 
      
      # Additional Data (AD Files)
      ad_files <- dir_ls(exp_path, regexp = "(?i)_AD_(\\d+)\\.(csv|xlsx)$")
      if (length(ad_files) > 0) {
        for (ad_path in ad_files) {
          ad_age <- str_extract(basename(ad_path), "(?i)_AD_(\\d+)") %>% str_extract("\\d+")
          ad_age <- str_pad(ad_age, width = 2, pad = "0") # <--- FORCES LEADING ZEROS
          message(paste("  -> Loading Additional Data:", basename(ad_path)))
          ad_raw <- if (str_detect(basename(ad_path), "(?i)\\.csv$")) read_csv(ad_path, col_types = cols(.default = "c")) else read_excel(ad_path, col_types = "text")
          
          names(ad_raw) <- str_to_upper(names(ad_raw))
          if("PLOT" %in% names(ad_raw)) ad_raw <- rename(ad_raw, Plot = PLOT)
          if("TREE" %in% names(ad_raw)) ad_raw <- rename(ad_raw, Tree = TREE)
          if("STEM" %in% names(ad_raw)) ad_raw <- rename(ad_raw, Tree = STEM)
          
          if ("Plot" %in% names(ad_raw)) {
            cols_to_exclude <- names(ad_raw)[names(ad_raw) %in% c("Plot", "Tree", "BLOCK", "REP")]
            ad_long_raw <- ad_raw %>% pivot_longer(cols = -any_of(cols_to_exclude), names_to = "Header_Original", values_to = "Value_Char") %>% filter(!is.na(Value_Char))
            
            if (nrow(ad_long_raw) > 0) {
              unique_headers <- unique(ad_long_raw$Header_Original)
              header_mapping <- map_dfr(unique_headers, function(h) {
                clean_h <- str_remove_all(str_to_upper(h), "[^A-Z0-9]")
                match <- trait_map %>% mutate(m = str_remove_all(str_to_upper(trait_code_FR), "[^A-Z0-9]")) %>% filter(clean_h == m) %>% slice(1)
                if(nrow(match) == 0) match <- trait_map %>% mutate(m = str_remove_all(str_to_upper(trait_code_FR), "[^A-Z0-9]")) %>% filter(str_starts(clean_h, m)) %>% slice(1)
                if(nrow(match) > 0) tibble(Header_Original = h, Code_DP = match$trait_code_DP, Key_FR = match$trait_code_FR) else tibble(Header_Original = h, Code_DP = replace_na(str_to_upper(str_extract(h, "^[A-Z]+")), h), Key_FR = h)
              })
              
              ad_long <- ad_long_raw %>% left_join(header_mapping, by = "Header_Original") %>%
                mutate(Plot = suppressWarnings(as.character(Plot)), Tree = if("Tree" %in% names(.)) suppressWarnings(as.numeric(Tree)) else 1,
                       Value_Num = suppressWarnings(as.numeric(Value_Char)), Age = ad_age, Trait = paste0(Code_DP, "_", Age), Prefix = Key_FR, Trait_Orig = Header_Original, UnitCode = "", Date = as.Date(NA)) %>%
                select(any_of(names(exp_data_long))) 
              exp_data_long <- bind_rows(exp_data_long, ad_long)
            }
          }
        }
      }
      
      exp_data_long <- exp_data_long %>% mutate(Trait = str_to_title(Trait))
      
      # --------------------------------------------------------------------------
      # NEW: Raw ASCII Text Recovery (.txt files)
      # --------------------------------------------------------------------------
      txt_files <- dir_ls(exp_path, regexp = "(?i)/[A-Za-z]+_\\d+\\.txt$")
        if (length(txt_files) > 0) {
        for (txt_path in txt_files) {
          filename <- tools::file_path_sans_ext(basename(txt_path))
          
          trait_prefix_raw <- str_extract(filename, "^[A-Za-z]+")
          txt_age <- str_extract(filename, "\\d+$")
          txt_age <- str_pad(txt_age, width = 2, pad = "0")
          
          clean_prefix <- str_to_upper(trait_prefix_raw)
          match <- trait_map %>% mutate(m = str_to_upper(trait_code_FR)) %>% filter(clean_prefix == m) %>% slice(1)
          if(nrow(match) == 0) match <- trait_map %>% mutate(m = str_to_upper(trait_code_FR)) %>% filter(str_starts(clean_prefix, m)) %>% slice(1)
          
          trait_dp <- if(nrow(match) > 0) match$trait_code_DP else str_to_title(trait_prefix_raw)
          final_trait_name <- paste0(trait_dp, "_", txt_age)
          
          message(paste("  -> Recovering ASCII Data:", basename(txt_path), "as", final_trait_name))
          
          # Parse the multi-line text file (Now passes prefix and age to hunt for units!)
          txt_df <- parse_cedd_raw_txt(txt_path, final_trait_name, trait_prefix_raw, txt_age)
          
          if (nrow(txt_df) > 0) {
            # Convert to standard long format
            txt_long <- txt_df %>%
              pivot_longer(cols = c(-Plot, -Tree, -Extracted_Unit), names_to = "Trait", values_to = "Value_Num") %>%
              filter(!is.na(Value_Num)) %>%
              mutate(
                Plot = as.character(Plot),
                Tree = as.numeric(Tree),
                Value_Char = as.character(Value_Num),
                Age = txt_age,
                Prefix = trait_prefix_raw,
                Trait_Orig = final_trait_name,
                UnitCode = Extracted_Unit, # USE THE UNIT WE FOUND IN THE HEADER!
                Date = as.Date(NA)
              ) %>%
              select(any_of(names(exp_data_long)))
            
            exp_data_long <- bind_rows(exp_data_long, txt_long)
          }
        }
      }
      
      # =====================================================================
      # 5. ASSESSOR BEHAVIORAL DIAGNOSTIC & CORRECTIONS
      # =====================================================================
      applied_corrections <- FALSE
      
      if(plot_type == "MULTI") {
        helper_path <- file.path(exp_path, "TRAVERSAL_HELPER_MASTER.csv")
        exp_mirror_flags <- tibble(Plot = character(), Tree = numeric(), Mirror_Flag = character())
        
        if (file.exists(helper_path)) {
          message("   -> Found TRAVERSAL_HELPER_MASTER.csv! Applying surgical orientation fixes...")
          applied_corrections <- TRUE
          
          traversal_helper <- read_csv(helper_path, show_col_types = FALSE) %>% mutate(Plot = as.character(Plot))
          
          # Process each rule (Plot/Trait combination) individually
          for (i in 1:nrow(traversal_helper)) {
            rule <- traversal_helper[i, ]
            
            target_data <- exp_data_long %>% filter(Plot == rule$Plot, Trait == rule$Trait_ID, !is.na(Value_Num))
            if (nrow(target_data) == 0) {
              message(paste("      -> WARNING: Correction skipped. Could not find data for Plot", rule$Plot, "Trait", rule$Trait_ID))
              next
            }
            
            # --- DYNAMIC GRID RESIZING FOR THE MAIN PIPELINE ---
            p_rows <- 10; p_cols <- 10 # Fallback sizes
            if (!is.null(spatial_info)) {
              sp_data <- spatial_info %>% filter(Plot == rule$Plot)
              if(nrow(sp_data) > 0) {
                p_rows <- max(sp_data$Prow, na.rm=TRUE) - min(sp_data$Prow, na.rm=TRUE) + 1
                p_cols <- max(sp_data$Ppos, na.rm=TRUE) - min(sp_data$Ppos, na.rm=TRUE) + 1
              }
            }
            
            # Auto-detect interior status using the new grid math
            trait_trees <- unique(target_data$Tree) %>% as.numeric()
            layout_interior <- as.vector(get_grid_layout(p_rows, p_cols, interior_only = TRUE))
            
            # If it's a small plot (like 8x1), interior logic doesn't apply
            is_trait_interior <- (length(layout_interior) > 0 && length(trait_trees) > 0 && all(trait_trees %in% layout_interior))
            
            layout_mat <- get_grid_layout(p_rows, p_cols, interior_only = is_trait_interior)
            canonical_path <- get_traversal_path(layout_mat, "top_left", "horizontal", FALSE)
            tested_path <- get_traversal_path(layout_mat, rule$Best_Start, rule$Best_Dir, as.logical(rule$Best_Snake))
            
            translation_df <- tibble(Original_Tree = as.numeric(canonical_path), Mapped_Tree = as.numeric(tested_path))
            
            # Apply the surgical swap
            exp_data_long <- exp_data_long %>%
              left_join(translation_df, by = c("Tree" = "Original_Tree")) %>%
              mutate(Tree = if_else(Plot == rule$Plot & Trait == rule$Trait_ID & !is.na(Mapped_Tree), Mapped_Tree, Tree)) %>%
              select(-Mapped_Tree)
            
            # Log the action so it appears in the final Validation_record column
            anchor_text <- if("Anchor_Used" %in% names(rule)) paste(" (Anchored to", rule$Anchor_Used, ")") else ""
            
            exp_mirror_flags <- bind_rows(exp_mirror_flags, tibble(
              Plot = rule$Plot, Tree = trait_trees, 
              Mirror_Flag = sprintf("ID Realigned from %s %s%s", toupper(rule$Best_Start), toupper(rule$Best_Dir), anchor_text)
            ))
          }
          
          message(paste("   -> Applied", nrow(traversal_helper), "plot-specific traversal corrections."))
        } else {
          message("   -> No TRAVERSAL_HELPER found. Using standard orientation.")
        }
      } else {
        # If plot_type is SINGLE, we skip traversal logic entirely
        exp_mirror_flags <- tibble(Plot = character(), Tree = numeric(), Mirror_Flag = character())
      }
      
      # --- 6. Survival Calc ---
      if(!is.null(spatial_info)) {
        raw_ages <- unique(na.omit(exp_data_long$Age))
        unique_ages <- raw_ages[!raw_ages %in% c("1", "01")]
        sur_list <- list()
        for(curr_age in unique_ages) {
          t_name <- paste0("Sur_", curr_age)
          if(t_name %in% unique(exp_data_long$Trait)) next
          
          # NOTE: Included the "Fr" exemption we added earlier!
          alive_trees <- exp_data_long %>% filter(Age == curr_age, !is.na(Value_Num), (Value_Num != 0 | str_detect(Trait, "(?i)^Fr"))) %>% select(Plot, Tree) %>% distinct() %>% mutate(is_alive = TRUE)
          
          sur_df <- spatial_info %>% select(Plot, Tree) %>% left_join(alive_trees, by = c("Plot", "Tree")) %>%
            mutate(Trait = t_name, Age = curr_age, is_alive = replace_na(is_alive, FALSE), Value_Num = if_else(is_alive, 1, 0), Value_Char = as.character(Value_Num), UnitCode = "$1-1", Trait_Orig = paste0("Calculated_", t_name), Prefix = "SUR", Date = as.Date(NA)) %>% select(-is_alive)
          sur_list[[t_name]] <- sur_df
        }
        if(length(sur_list) > 0) exp_data_long <- bind_rows(exp_data_long, bind_rows(sur_list))
        
      } else {
        
        # --- FALLBACK SYNTHETIC ROSTER FOR MATRIX-LESS TRIALS ---
        message("   -> No Spatial Matrix found: Generating synthetic tree roster for survival calculations.")
        
        # 1. Build a roster from 1 to the highest Tree ID ever seen in each plot
        synthetic_roster <- exp_data_long %>%
          group_by(Plot) %>%
          summarise(Max_Tree = max(Tree, na.rm = TRUE), .groups = "drop") %>%
          uncount(Max_Tree, .id = "Tree")
        
        raw_ages <- unique(na.omit(exp_data_long$Age))
        unique_ages <- raw_ages[!raw_ages %in% c("1", "01")]
        sur_list <- list()
        
        for(curr_age in unique_ages) {
          t_name <- paste0("Sur_", curr_age)
          if(t_name %in% unique(exp_data_long$Trait)) next
          
          # Find who was alive this year
          alive_trees <- exp_data_long %>% filter(Age == curr_age, !is.na(Value_Num), (Value_Num != 0 | str_detect(Trait, "(?i)^Fr"))) %>% select(Plot, Tree) %>% distinct() %>% mutate(is_alive = TRUE)
          
          # Cross-reference against the synthetic roster instead of the spatial matrix
          sur_df <- synthetic_roster %>% left_join(alive_trees, by = c("Plot", "Tree")) %>%
            mutate(Trait = t_name, Age = curr_age, is_alive = replace_na(is_alive, FALSE), Value_Num = if_else(is_alive, 1, 0), Value_Char = as.character(Value_Num), UnitCode = "$1-1", Trait_Orig = paste0("Calculated_", t_name), Prefix = "SUR", Date = as.Date(NA)) %>% select(-is_alive)
          sur_list[[t_name]] <- sur_df
        }
        if(length(sur_list) > 0) exp_data_long <- bind_rows(exp_data_long, bind_rows(sur_list))
      }
      
      # --- 7. Flags & Logic ---
      data_with_outliers <- exp_data_long %>%
        mutate(Is_Target_For_Zero_Removal = str_detect(Trait, "(?i)^(Pil|Br|St)"), 
               Value_Num = if_else(Is_Target_For_Zero_Removal & Value_Num == 0, NA_real_, Value_Num), 
               Value_Char = if_else(is.na(Value_Num), NA_character_, as.character(Value_Num))) %>% 
        select(-Is_Target_For_Zero_Removal) %>% 
        group_by(Trait) %>% 
        mutate(
          n_unique = n_distinct(Value_Num, na.rm = TRUE), 
          Is_Survival = str_detect(Trait, "(?i)Sur"), 
          # ADDED Fr to the binary regex
          Is_Binary_Resi = str_detect(Trait, "(?i)(Resi|Drill|Amp|Den|Cull|Stem|Fr)") & n_unique <= 2, 
          Is_Ordinal = !Is_Survival & !Is_Binary_Resi & n_unique < 9, 
          Mean_Val = mean(Value_Num, na.rm = TRUE), 
          SD_Val = sd(Value_Num, na.rm = TRUE), 
          Lower_Limit = Mean_Val - (4 * SD_Val), 
          Upper_Limit = Mean_Val + (4 * SD_Val),
          Flag_Outlier = case_when(
            # EXEMPTED Fr from the Ordinal Zero flags
            !is.na(Value_Num) & Is_Ordinal & Value_Num == 0 & !str_detect(Trait, "(?i)^Fr") ~ TRUE, 
            Is_Ordinal & Value_Num == 0 & !str_detect(Trait, "(?i)^Fr") ~ TRUE, 
            (Is_Survival | Is_Binary_Resi) & Value_Num > 1 ~ TRUE, 
            !Is_Survival & !Is_Ordinal & !Is_Binary_Resi & !is.na(Value_Num) & (Value_Num < Lower_Limit | Value_Num > Upper_Limit) ~ TRUE, 
            TRUE ~ FALSE
          )
        ) %>% ungroup()
           
      global_age_seq <- data_with_outliers %>% filter(!is.na(Value_Num), str_detect(Trait, "_\\d+$")) %>% mutate(Base_Trait = str_remove(Trait, "_\\d+$"), Age_Num = suppressWarnings(as.numeric(str_extract(Trait, "\\d+$")))) %>% filter(!is.na(Age_Num)) %>% distinct(Base_Trait, Age_Num) %>% arrange(Base_Trait, Age_Num) %>% group_by(Base_Trait) %>% mutate(Expected_Prev_Age = lag(Age_Num)) %>% ungroup()
      
      
      # --- Shrinkage Flags: ---
      shrinkage_flags <- data_with_outliers %>% filter(!is.na(Value_Num), str_detect(Trait, "_\\d+$"), !Is_Ordinal, !Is_Binary_Resi, !str_detect(Trait, "(?i)(Sur|Resi|Detrended|Ampli|Resistance|Resist|Crack|Form|Br|St|Pil)")) %>%
        mutate(Base_Trait = str_remove(Trait, "_\\d+$"), Age_Num = suppressWarnings(as.numeric(str_extract(Trait, "\\d+$")))) %>% filter(!is.na(Age_Num)) %>% left_join(global_age_seq, by = c("Base_Trait", "Age_Num")) %>%
        arrange(Plot, Tree, Base_Trait, Age_Num) %>% group_by(Plot, Tree, Base_Trait) %>% mutate(Prev_Val = lag(Value_Num), Prev_Age = lag(Age_Num), Prev_Unit = lag(UnitCode), Is_Shrinkage = !is.na(Prev_Val) & !is.na(Expected_Prev_Age) & (Prev_Age == Expected_Prev_Age) & (UnitCode == Prev_Unit) & (Value_Num < Prev_Val)) %>%
        ungroup() %>% filter(Is_Shrinkage) %>% select(Plot, Tree, Trait, Prev_Age) %>% mutate(Shrinkage_Error = paste("Shrinkage detected (smaller than Age", Prev_Age, "with matching units)"))
      
      final_long_data <- data_with_outliers %>% left_join(shrinkage_flags %>% select(-Prev_Age), by = c("Plot", "Tree", "Trait")) %>% mutate(Reject_Flag = case_when(is.na(Value_Num) ~ NA_real_, Flag_Outlier ~ 1, !is.na(Shrinkage_Error) ~ 1, TRUE ~ 0))
      
      # --- 8. Output Generation ---
      final_long_dedup <- final_long_data %>% arrange(Plot, Tree, Trait, desc(Date)) %>% distinct(Plot, Tree, Trait, .keep_all = TRUE)
      
      validation_summary <- final_long_dedup %>% filter(Reject_Flag == 1) %>% mutate(Error_Msg = case_when(Flag_Outlier & !is.na(Shrinkage_Error) ~ paste(Trait, "flagged as outlier AND", Shrinkage_Error), Flag_Outlier ~ paste(Trait, "flagged for being an extreme outlier"), !is.na(Shrinkage_Error) ~ paste(Trait, Shrinkage_Error), TRUE ~ paste(Trait, "flagged"))) %>% group_by(Plot, Tree) %>% summarise(Validation_record = paste(Error_Msg, collapse = " | "), .groups = "drop")
      
      alive_summary <- NULL
      sur_traits <- unique(final_long_dedup$Trait[str_detect(final_long_dedup$Trait, "(?i)Sur_")])
      if(length(sur_traits) > 0) {
        target_row <- tibble(Trait_Name = sur_traits, Age_Num = as.numeric(str_extract(sur_traits, "\\d+"))) %>% filter(Age_Num == max(Age_Num, na.rm = TRUE)) %>% slice(1)
        if(nrow(target_row) > 0) alive_summary <- final_long_dedup %>% filter(Trait == target_row$Trait_Name) %>% select(Plot, Tree, Alive = Value_Num) %>% distinct()
      }
      
      data_wide_clean <- final_long_dedup %>% select(Plot, Tree, Trait, Value_Char) %>% pivot_wider(names_from = Trait, values_from = Value_Char)
      exp_data_wide_numeric <- final_long_dedup %>% select(Plot, Tree, Trait, Value_Num) %>% pivot_wider(names_from = Trait, values_from = Value_Num)
      if (!is.null(spatial_info)) exp_data_wide_numeric <- left_join(exp_data_wide_numeric, spatial_info, by = c("Plot", "Tree"))
      
      flags_wide <- final_long_dedup %>% select(Plot, Tree, Trait, Reject_Flag) %>% pivot_wider(names_from = Trait, values_from = Reject_Flag, names_glue = "{Trait}_reject")
      
      final_wide_with_flags <- left_join(data_wide_clean, flags_wide, by = c("Plot", "Tree")) %>% left_join(validation_summary, by = c("Plot", "Tree"))
      if(!is.null(alive_summary)) final_wide_with_flags <- left_join(final_wide_with_flags, alive_summary, by = c("Plot", "Tree"))
      if(!is.null(genetic_info)) final_wide_with_flags <- left_join(final_wide_with_flags, genetic_info %>% mutate(Plot=as.character(Plot)), by = "Plot")
      if(!is.null(spatial_info)) final_wide_with_flags <- left_join(final_wide_with_flags, spatial_info, by = c("Plot", "Tree"))
      
      # Drop plots that failed the genetic join (cannot be imported to DMS)
      if ("Family_name" %in% names(final_wide_with_flags)) {
        dropped_records <- sum(is.na(final_wide_with_flags$Family_name))
        if (dropped_records > 0) {
          message(paste("    -> WARNING: Dropping", dropped_records, "records missing Family/Block data (likely omitted from Design file)."))
          final_wide_with_flags <- final_wide_with_flags %>% filter(!is.na(Family_name))
        }
      }
      
      final_wide_with_flags <- final_wide_with_flags %>%
        mutate(Ploc = str_replace_all(curr_exp, " ", ""))
      
      metadata_cols <- c("Ploc","Validation_record", "Alive", "Block", "SubBlock", "Family_name", "Prow", "Ppos")
      existing_meta_cols <- intersect(metadata_cols, names(final_wide_with_flags))
      
      col_sorting_df <- tibble(col_name = setdiff(names(final_wide_with_flags), c("Plot", "Tree", metadata_cols))) %>%
        mutate(is_reject = str_detect(col_name, "(?i)_reject$"), base_full = str_remove(col_name, "(?i)_reject$"), age_num = suppressWarnings(as.numeric(str_extract(base_full, "\\d+$"))), age_sort = replace_na(age_num, 0), base_trait = str_remove(base_full, "_\\d+$"), is_survival = str_detect(base_trait, "(?i)^Sur")) %>% arrange(is_survival, age_sort, base_trait, is_reject)
      
      final_wide_with_flags <- final_wide_with_flags %>% select(all_of(c("Plot", "Tree", existing_meta_cols, col_sorting_df$col_name))) %>% arrange(Plot, Tree)
      
      # Merge Mirror Flags
      if(nrow(exp_mirror_flags) > 0) {
        exp_mirror_summary <- exp_mirror_flags %>% mutate(Plot = as.character(Plot), Tree = as.numeric(Tree)) %>% group_by(Plot, Tree) %>% summarise(Combined_Mirror_Flags = paste(na.omit(unique(Mirror_Flag)), collapse = " | "), .groups = "drop")
        final_wide_with_flags <- final_wide_with_flags %>% left_join(exp_mirror_summary, by = c("Plot", "Tree")) %>% mutate(Validation_record = case_when(!is.na(Combined_Mirror_Flags) & (is.na(Validation_record) | Validation_record == "") ~ Combined_Mirror_Flags, !is.na(Combined_Mirror_Flags) ~ paste(Validation_record, "|", Combined_Mirror_Flags), TRUE ~ Validation_record)) %>% select(-Combined_Mirror_Flags)
      }
      
      out_suffix <- if (applied_corrections) "_Corrected" else ""
      write_csv(final_wide_with_flags, file.path(exp_path, paste0(file_prefix, "_Full_Data_With_Flags", out_suffix, ".csv")), na = "")
      
      
      # PDF & Stats
      try({
        # FIX: Moved out_suffix before the file extension!
        pdf(file.path(exp_path, paste0(file_prefix, out_suffix, "_graphs.pdf")), width = 8, height = 11)
        generate_pdf_report(final_long_dedup, exp_data_wide_numeric, curr_exp, unique(final_long_dedup$Prefix), unique(final_long_dedup$Trait))
        dev.off()
      })
      while(dev.cur() > 1) dev.off()
      
      stats_df <- final_long_dedup %>% filter(Reject_Flag == 0) %>% group_by(Trait) %>% summarise(N_Valid = sum(!is.na(Value_Num)), Mean = mean(Value_Num, na.rm=TRUE), Std_Dev = sd(Value_Num, na.rm=TRUE), Min = min(Value_Num, na.rm=TRUE), Max = max(Value_Num, na.rm=TRUE), CV_Pct = (sd(Value_Num, na.rm=TRUE)/mean(Value_Num, na.rm=TRUE))*100) %>% mutate(across(where(is.numeric), ~ round(., 2)))
      write_csv(stats_df, file.path(exp_path, paste0(curr_exp, "_Stats", out_suffix, ".csv")), na = "")
      
      # --- Trait-Trait Correlations Table ---
      message("  -> Generating Trait-Trait Correlation Table...")
      
      # 1. Pull the numeric dataset and strictly filter for continuous biological traits
      # BULLETPROOF: Drops all known IDs/Coords, then forces everything else to be strictly numeric
      cor_data <- exp_data_wide_numeric %>% 
        select(-any_of(c("Plot", "Tree", "InferredTreePosition", "Row", "Position", "Prow", "Ppos", "Block", "SubBlock", "Validation_record"))) %>%
        select(-matches("(?i)^Sur_")) %>%
        select(where(is.numeric)) 
      
      if (ncol(cor_data) >= 2) {
        
        # 2. Filter 0s for physical traits (Heights/Diameters) to match PDF logic
        for (trt in names(cor_data)) {
          if (!str_detect(trt, "(?i)^(Pil|Br|St|Sur|Fr)")) {
            cor_data[[trt]] <- na_if(cor_data[[trt]], 0)
          }
        }
        
        # 3. Calculate correlations and N-obs
        cor_matrix <- cor(cor_data, use = "pairwise.complete.obs", method = "spearman")
        n_matrix <- crossprod(!is.na(cor_data))
        
        # 4. Melt into a tidy table
        cor_table <- as.data.frame(as.table(cor_matrix)) %>%
          rename(Trait_1 = Var1, Trait_2 = Var2, Spearman_Rho = Freq) %>%
          filter(as.character(Trait_1) < as.character(Trait_2)) %>% # Keep upper triangle only
          filter(!is.na(Spearman_Rho)) %>%
          mutate(Spearman_Rho = round(Spearman_Rho, 3))
        
        n_table <- as.data.frame(as.table(n_matrix)) %>%
          rename(Trait_1 = Var1, Trait_2 = Var2, N_Obs = Freq)
        
        cor_table <- cor_table %>%
          left_join(n_table, by = c("Trait_1", "Trait_2")) %>%
          arrange(desc(Spearman_Rho))
        
        # 5. Export Individual Trial Table
        write_csv(cor_table, file.path(exp_path, paste0(curr_exp, "_Trait_Correlations", out_suffix, ".csv")))
        
        # 6. Add to our Master List for the global summary
        master_correlations_list[[curr_exp]] <- cor_table %>% mutate(Trial = curr_exp)
        
      } else {
        message("  -> Skipping correlation table: Not enough continuous traits.")
      }
      
      # --- 9. XML Generation ---
      xml_data <- data_with_outliers %>% group_by(Trait) %>% mutate(Calc_Vals = if_else(Value_Num > 0, Value_Num, NA_real_), Min_Val = min(Calc_Vals, na.rm = TRUE), Max_Val = max(Calc_Vals, na.rm = TRUE), Trait_Date = if(length(Date[!is.na(Date)]) > 0) as.character(Date[!is.na(Date)][1]) else NA_character_) %>% ungroup() %>% distinct(Trait, Trait_Orig, Prefix, Age, UnitCode, Is_Ordinal, Min_Val, Max_Val, Trait_Date) %>% left_join(trait_map, by = c("Prefix" = "trait_code_FR")) %>%
        mutate(Group = case_when(!is.na(xml_group) & xml_group != "" ~ xml_group, str_detect(Prefix, "(?i)SUR") ~ "Health", TRUE ~ "Uncategorized"), Desc_Base = coalesce(xml_desc_base, Prefix), Real_Units = case_when(!is.na(xml_units) & xml_units != "none" & xml_units != "" ~ xml_units, Is_Ordinal & is.finite(Min_Val) ~ paste0("$", Min_Val, "-", Max_Val), str_detect(Trait, "(?i)Sur") ~ "0/$1", str_detect(Trait, "(?i)Pil") ~ "mm", str_detect(Trait, "(?i)Av") ~ "m/s", !is.na(UnitCode) & UnitCode != "" ~ str_to_lower(UnitCode), TRUE ~ "unitless"), Full_Description = paste0(Desc_Base, " Age ", Age, " (was ", Trait_Orig, ")", if_else(!is.na(Trait_Date), paste(" Date:", Trait_Date), "")), XML_Tag = sprintf('<trait group_name="%s" name="%s" trait_type="M" description="%s" data_type="N" is_solver_mappable="1" summary_rule="first_with_reject" units="%s" validate="none" />', Group, Trait, Full_Description, Real_Units)) %>% arrange(Trait)
      
      # FIX: Added out_suffix to the XML export as well!
      writeLines(c("<trial>", "<traitlist>", xml_data$XML_Tag, "</traitlist>", "</trial>"), file.path(exp_path, paste0(file_prefix, out_suffix, "_Traits.xml")))
      message("  -> Processed & Saved.")
      
    }, error = function(e) { message(paste("  -> ERROR processing", curr_exp, ": ", e$message)) })
  }
  
  # =====================================================================
  # 6. GLOBAL MULTI-TRIAL SYNTHESIS
  # =====================================================================
  message("\n==========================================")
  message("Generating Master Multi-Trial Correlation Table...")
  
  if (length(master_correlations_list) > 0) {
    # Combine all trials into one massive table (No summary math!)
    master_cor_df <- bind_rows(master_correlations_list) %>%
      # Reorder columns so Trial is the first thing you see
      select(Trial, Trait_1, Trait_2, Spearman_Rho, N_Obs) %>%
      # Sort alphabetically by Trial, then by highest correlation
      arrange(Trial, desc(Spearman_Rho))
    
    # Save to the root folder (e.g., "Sitka/High GCA Fullsib P85-P87 experiments")
    summary_path <- file.path(root_data_dir, paste0("MASTER_Trait_Correlations_", format(Sys.Date(), "%Y%m%d"), ".csv"))
    write_csv(master_cor_df, summary_path)
    
    message(paste("  -> Master correlation table saved to:", summary_path))
  } else {
    message("  -> No correlation data found across any trials to compile.")
  }
  message("==========================================\n")
  
}

