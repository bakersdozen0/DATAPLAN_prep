#### #### #### ####
#### #### #### ####
 
data_dir<-"//forestresearch.gov.uk/shares/CSFCC/Forest Resource and Product Assessment and Improvement/NRS-Tree Improvement/CONIFERS/SITKA SPRUCE/psi_DATAPLAN_prep"



library(usethis)
library(tidyverse)
library(readxl)
library(fs)
library(gridExtra)
library(grid)
library(here)

# # # # # # # # # # # # # # # # # # # # # # # 
# PART 0: HELPER FUNCTIONS #### 
# # # # # # # # # # # # # # # # # # # # # # # 

# --- 1. Parse Design File (TXT) ---
parse_long_design_file <- function(filepath, exp_prefix) {
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
    extract(line, into = c("Design_ID", "Cross_Name"), 
            regex = "^\\s*\\*?\\s*(\\d+)=(.*)$", convert = TRUE) %>%
    mutate(Cross_Name = str_trim(Cross_Name)) %>%
    # Cleanly separate the Mother and Father into their own columns for standard crosses
    extract(
      Cross_Name, 
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
        str_trim(str_replace(Cross_Name, "^SS\\s*", "")),
        NA_character_
      )
    )
  
  # Parse PLOTS
  plot_lines <- raw_lines[(plot_start + 1):length(raw_lines)]
  plot_lines <- plot_lines[grep(":", plot_lines)]
  
  # NEW: Regex captures Plot, Design_ID, SubBlock, AND Block
  plot_df <- tibble(line = plot_lines) %>%
    extract(line, into = c("Plot", "Design_ID", "SubBlock", "Block"), 
            regex = "^\\s*(\\d+):\\s*(\\d+)\\s+(\\d+)\\s+(\\d+)", convert = TRUE)
  
  # Join and build Family_name
  final_design <- plot_df %>%
    left_join(trt_df, by = "Design_ID") %>%
    mutate(
      Plot = as.numeric(Plot),
      Block = as.character(Block),
      SubBlock = as.character(SubBlock),
      Family_name = case_when(
        Design_ID == 0 ~ paste0(exp_prefix, "_Filler"),
        
        # Map controls exactly like the XLSX function
        !is.na(Control_Name) ~ paste0("ss", Control_Name),
        
        # Check the extracted Paternal_ID for OP, and build with Maternal_ID
        !is.na(Paternal_ID) & str_detect(Paternal_ID, "(?i)OP") ~ paste0("ss", Maternal_ID, "_OPCB"),
        !is.na(Maternal_ID) & !is.na(Paternal_ID) ~ paste0("ss", Maternal_ID, "_ss", Paternal_ID),
        
        # Fallback just in case formatting is totally broken
        TRUE ~ str_replace_all(Cross_Name, "\\s+", "") %>% str_replace("^SS", "ss_")
      )
    ) %>%  
    # NEW: Keep SubBlock in the final selection
    select(Plot, Block, SubBlock, Family_name)
  
  return(final_design)
}

# --- 1b. Parse XLSX Design File ---
parse_xlsx_design_file <- function(filepath, exp_prefix) {
  raw_design <- read_excel(filepath, col_types = "text")
  
  if("Rep" %in% names(raw_design) && !"Block" %in% names(raw_design)) {
    raw_design <- raw_design %>% rename(Block = Rep)
  }
  
  clean_design <- raw_design %>%
    select(any_of(c("Plot", "Block", "Seedlot", "SubBlock"))) %>%
    filter(!is.na(Plot)) %>%
    mutate(Seedlot = if_else(str_trim(Seedlot) == "", NA_character_, Seedlot)) %>%
    # 1. Extract standard crosses (those with two SS's)
    extract(
      Seedlot,
      into = c("Maternal_ID", "Paternal_ID"),
      regex = "^.*?=\\s*SS\\s*(\\d+)\\s*SS\\s*(.*)$",
      remove = FALSE
    ) %>%
    mutate(
      Plot = suppressWarnings(as.numeric(Plot)),
      Maternal_ID = str_trim(Maternal_ID),
      Paternal_ID = str_trim(Paternal_ID),
      
      # 2. Safely extract Control names (looks for the starting asterisk)
      Control_Name = if_else(
        str_detect(Seedlot, "^\\*"),
        str_trim(str_replace(Seedlot, "^\\*.*?=\\s*SS\\s*", "")),
        NA_character_
      ),
      
      # 3. Updated Family_name logic
      Family_name = case_when(
        # If it's a control, format it (e.g., "Control_70(7111)1")
        !is.na(Control_Name) ~ paste0("ss", Control_Name),
        
        # If it's an Open Pollinated cross
        !is.na(Paternal_ID) & str_detect(Paternal_ID, "(?i)OP") ~ paste0("ss", Maternal_ID, "_OPCB"),
        
        # If it's a valid standard cross
        !is.na(Maternal_ID) & !is.na(Paternal_ID) ~ paste0("ss", Maternal_ID, "_ss", Paternal_ID),
        
        # Only assign Filler if the Seedlot is genuinely missing/blank
        is.na(Seedlot) ~ paste0(exp_prefix, "_Filler"),
        
        # Catch-all fallback
        TRUE ~ paste0(exp_prefix, "_Filler")
      )
    ) %>%
    # Clean up the temporary control column
    select(any_of(c("Plot", "Block", "SubBlock", "Family_name")))
  
  return(clean_design)
}

# --- 2. PDF Reporting Function (UPDATED FOR MULTI-TREE & NUMERIC COORDS) ---
generate_pdf_report <- function(long_data, wide_data, exp_name, multi_age_prefixes, all_traits) {
  
  # 1. Get ALL continuous traits first (Excluding Survival)
  all_cont_traits <- long_data %>% 
    filter(Is_Ordinal == FALSE, !str_detect(Trait, "(?i)Sur_")) %>% 
    distinct(Trait) %>% pull(Trait)
  
  # 2. Split into "Integer-Based" (Pilodyn) and "Standard Continuous"
  pil_traits  <- all_cont_traits[str_detect(all_cont_traits, "(?i)Pil")]
  cont_traits <- setdiff(all_cont_traits, pil_traits)
  
  ord_traits  <- long_data %>% 
    filter(Is_Ordinal == TRUE, !str_detect(Trait, "(?i)Sur_"))  %>% 
    distinct(Trait) %>% pull(Trait)
  
  # --- A1. Distributions (Standard Continuous) ---
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
      if (str_detect(base_t, "(?i)(Sur|Resi|Detrended|Ampli|Resistance|Resist|Crack|Form|Br|St|Pil)")) next      
      # 2. STRICTLY match only traits with this exact base and an age number
      # e.g., "^HGT_\\d+$" finds HGT_10 and HGT_15, but ignores HGT_ADJ_10
      regex_pattern <- paste0("^", base_t, "_\\d+$")
      cols <- str_sort(all_traits[str_detect(all_traits, regex_pattern)], numeric = TRUE)
      
      if (length(cols) < 2) next
      
      for (i in 1:(length(cols) - 1)) {
        younger <- cols[i]; older <- cols[i+1]
        if(!all(c(younger, older) %in% names(wide_data))) next
        
        # --- NEW: Verify units match before plotting ---
        unit_younger <- long_data %>% filter(Trait == younger) %>% drop_na(UnitCode) %>% pull(UnitCode) %>% head(1)
        unit_older <- long_data %>% filter(Trait == older) %>% drop_na(UnitCode) %>% pull(UnitCode) %>% head(1)
        
        # If both have units but they don't match, skip generating this specific plot
        if(length(unit_younger) > 0 && length(unit_older) > 0 && unit_younger != unit_older) next
        # -----------------------------------------------
        
        plot_df <- wide_data %>%
          select(Plot, Tree, X = all_of(older), Y = all_of(younger)) %>%
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
  
  # --- D. Concentric Temporal Check ---
  if (all(c("Prow", "Ppos") %in% names(wide_data))) {
    
    # 1. Get Master List of Plot + Tree combinations
    all_units <- wide_data %>% select(Plot, Tree) %>% distinct()
    
    # 2. Filter Data
    physical_data <- long_data %>%
      mutate(Age_Num = suppressWarnings(as.numeric(as.character(Age)))) %>%
      filter(!str_detect(Trait, "(?i)Sur_"), !is.na(Value_Num), !is.na(Age_Num), Age_Num > 0)
    
    plot_ages <- sort(unique(physical_data$Age_Num), decreasing = TRUE)
    
    if (length(plot_ages) > 0) {
      age_label <- paste(plot_ages, collapse = ", ")
      
      # 4. Calculate Counts grouped by Plot AND Tree
      counts_per_tree <- physical_data %>%
        group_by(Plot, Tree, Age_Num) %>%
        summarise(Count = n(), .groups = "drop")
      
      # 5. Expand grid using Plot AND Tree, then join
      full_grid <- expand_grid(all_units, Age_Num = plot_ages) %>%
        left_join(counts_per_tree, by = c("Plot", "Tree", "Age_Num")) %>%
        replace_na(list(Count = 0))
      
      # 6. Calculate Mode
      mode_per_age <- full_grid %>%
        filter(Count > 0) %>% 
        group_by(Age_Num) %>%
        summarise(Standard_Count = as.numeric(names(sort(table(Count), decreasing=TRUE)[1])), .groups = "drop")
      
      # 7. Join to spatial data and FORCE numeric coords
      plot_data <- full_grid %>%
        left_join(mode_per_age, by = "Age_Num") %>%
        mutate(
          Standard_Count = replace_na(Standard_Count, 1), 
          Percent = (Count / Standard_Count) * 100
        ) %>%
        left_join(wide_data %>% select(Plot, Tree, Prow, Ppos) %>% distinct(), by = c("Plot", "Tree")) %>%
        mutate(
          Prow = suppressWarnings(as.numeric(Prow)), 
          Ppos = suppressWarnings(as.numeric(Ppos))
        ) %>%
        filter(!is.na(Prow), !is.na(Ppos))
      
      if(nrow(plot_data) > 0) {
        # 8. Define Axis Breaks
        min_row <- min(plot_data$Prow, na.rm = TRUE)
        max_row <- max(plot_data$Prow, na.rm = TRUE)
        min_pos <- min(plot_data$Ppos, na.rm = TRUE)
        max_pos <- max(plot_data$Ppos, na.rm = TRUE)
        
        # 9. Plotting
        sizes <- seq(0.95, 0.25, length.out = length(plot_ages))
        
        p_concentric <- ggplot() +
          scale_y_reverse(breaks = seq(min_row, max_row, by = 5)) +
          scale_x_continuous(breaks = seq(min_pos, max_pos, by = 5)) +
          coord_fixed() +
          theme_minimal() +
          theme(axis.text.x = element_text(angle = 90, vjust = 0.5, size = 8),
                axis.text.y = element_text(size = 8)) +
          labs(title = paste(exp_name, "- Data Completeness"),
               subtitle = paste0("Ages Plotted: ", age_label, "\nOuter Ring = Oldest -> Inner Square = Youngest"),
               x = "Position", y = "Row", fill = "% Data") +
          scale_fill_gradient2(low = "red", mid = "yellow", high = "forestgreen", midpoint = 50, limits = c(0, 100), oob = scales::squish)
        
        for (i in seq_along(plot_ages)) {
          curr_age <- plot_ages[i]
          curr_size <- sizes[i]
          layer_data <- plot_data %>% filter(Age_Num == curr_age)
          
          p_concentric <- p_concentric + 
            geom_tile(data = layer_data, aes(x = Ppos, y = Prow, fill = Percent), 
                      width = curr_size, height = curr_size, color = "black", size = 0.1) 
        }
        print(p_concentric)
      }
    }
  }
  
  # --- H. Resurrection Check (Zombie) ---
  sur_cols <- sort(names(wide_data)[str_detect(names(wide_data), "(?i)Sur_\\d+")])
  
  if (length(sur_cols) >= 2) {
    sur_ages <- as.numeric(str_extract(sur_cols, "\\d+"))
    sur_map  <- data.frame(Col = sur_cols, Age = sur_ages) %>% arrange(Age)
    
    for (i in 1:(nrow(sur_map) - 1)) {
      younger_col <- sur_map$Col[i]; older_col <- sur_map$Col[i+1]
      
      zombies <- wide_data %>%
        filter(.data[[younger_col]] == 0 & .data[[older_col]] == 1) %>%
        select(Plot, Tree) %>% 
        distinct() %>%
        arrange(Plot, Tree)
      
      if (nrow(zombies) > 0) {
        bad_ids <- paste(zombies$Plot, zombies$Tree, sep="-")
        id_string <- paste(bad_ids, collapse = ", ")
        wrapped_text <- str_wrap(id_string, width = 80)
        
        final_text <- paste0(
          "RESURRECTION CHECK (Zombie Trees)\n",
          "Logic: Marked Dead (0) at ", younger_col, " -> Alive (1) at ", older_col, "\n",
          "Total Flagged: ", length(bad_ids), "\n\n",
          "Plot-Tree IDs:\n", wrapped_text
        )
        
        t_grob <- textGrob(final_text, x = 0.5, y = 0.5, just = "center", gp = gpar(fontsize = 10, fontface = "plain"))
        grid.arrange(t_grob)
      }
    }
  }
  
  # --- E. Correlations ---
  if (length(all_cont_traits) >= 2) {
    gen_pairs <- combn(all_cont_traits, 2, simplify = FALSE)
    for (pair in gen_pairs) {
      t1 <- pair[1]; t2 <- pair[2]
      if(!all(c(t1, t2) %in% names(wide_data))) next
      
      outlier_plots <- long_data %>% 
        filter(Trait %in% c(t1, t2), Flag_Outlier == TRUE) %>% 
        select(Plot, Tree) %>% distinct() %>% mutate(Status = "Outlier")
      
      plot_df <- wide_data %>%
        select(Plot, Tree, X_Val = all_of(t1), Y_Val = all_of(t2)) %>%
        filter(!is.na(X_Val), !is.na(Y_Val)) %>%
        left_join(outlier_plots, by = c("Plot", "Tree")) %>%
        mutate(Status = replace_na(Status, "Normal"))
      
      if(nrow(plot_df) == 0) next
      
      p_corr <- ggplot(plot_df, aes(x=X_Val, y=Y_Val)) + 
        geom_smooth(method = "lm", se = FALSE, color = "gray", linetype = "dashed", alpha=0.5) +
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
      
      # Force Prow/Ppos to numeric just in case, preventing the unary minus error
      map_data <- wide_data %>% 
        select(Prow, Ppos, Value = all_of(trait)) %>% 
        mutate(Prow = as.numeric(Prow), Ppos = as.numeric(Ppos)) %>%
        filter(!is.na(Value), !is.na(Prow), !is.na(Ppos))
      
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
      # Force everything to be a character string safely
      plot <- as.character(matrix_data[i, j])
      
      # Safely skip empty cells, NAs, or 0s without triggering an error
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

# --- Map interior tree sequence to full plot sequence ---
map_interior_trees <- function(tree_idx, subset_size, full_size) {
  # If sizes match, or if it's an NA/empty row, return original index
  if (is.na(subset_size) || is.na(full_size) || subset_size == full_size) {
    return(tree_idx)
  }
  
  # Check if they are perfect squares (49, 25, 9)
  n <- sqrt(full_size)
  m <- sqrt(subset_size)
  
  if (n %% 1 != 0 || m %% 1 != 0) {
    # If not a square (e.g., uneven row planting), return as-is to avoid bad math
    return(tree_idx)
  }
  
  # Calculate how many rows/cols of trees are skipped on the edges
  offset <- (n - m) / 2
  
  # Convert the 1D subset index to 2D subset coordinates (1-based)
  row_m <- ceiling(tree_idx / m)
  col_m <- (tree_idx - 1) %% m + 1
  
  # Shift coordinates by the edge buffer
  row_n <- row_m + offset
  col_n <- col_m + offset
  
  # Convert shifted 2D coordinates back to the 1D full-plot index
  mapped_idx <- (row_n - 1) * n + col_n
  
  return(mapped_idx)
}

# --- 4. Detect and Reverse Mirrored Plots (Trait-by-Trait Vector Method) ---
fix_mirrored_traits_8x1 <- function(long_data) {
  
  # Initialize an empty dataframe to catch our flags
  mirror_flags <- tibble(Plot = character(), Tree = numeric(), Trait = character(), Mirror_Flag = character())
  
  exp_max_tree <- max(suppressWarnings(as.numeric(long_data$Tree)), na.rm = TRUE)
  if (exp_max_tree != 8) {
    message(paste("  -> Skipping mirror fix: Experiment max tree is", exp_max_tree, "(Not an 8x1 design)."))
    return(list(data = long_data, flags = mirror_flags)) # Return list
  }
  
  valid_plots <- long_data %>%
    mutate(Tree_Num = suppressWarnings(as.numeric(Tree))) %>%
    filter(!is.na(Tree_Num), Tree_Num <= 8) %>%
    pull(Plot) %>% unique()
  
  if(length(valid_plots) == 0) return(list(data = long_data, flags = mirror_flags))
  
  fixed_data <- long_data
  plots_fixed_count <- 0
  
  for (p in valid_plots) {
    plot_data <- long_data %>% 
      filter(Plot == p, !is.na(Value_Num), !str_detect(Trait, "(?i)Sur_")) %>%
      mutate(
        Age_Num = suppressWarnings(as.numeric(str_extract(Trait, "\\d+$"))),
        Tree_Num = suppressWarnings(as.numeric(Tree))
      ) %>%
      filter(!is.na(Age_Num), !is.na(Tree_Num))
    
    plot_ages <- sort(unique(plot_data$Age_Num))
    if(length(plot_ages) < 2) next
    
    plot_fixed_any <- FALSE
    baseline_alive <- plot_data %>% filter(Age_Num == plot_ages[1]) %>% pull(Tree_Num) %>% unique()
    
    for (i in 2:length(plot_ages)) {
      target_age <- plot_ages[i]
      traits_at_age <- plot_data %>% filter(Age_Num == target_age) %>% pull(Trait) %>% unique()
      
      trees_measured_this_age <- c()
      
      for (trt in traits_at_age) {
        measured_trees_raw <- plot_data %>% filter(Trait == trt) %>% pull(Tree_Num) %>% unique()
        res_raw <- sum(!measured_trees_raw %in% baseline_alive)
        
        if (res_raw > 0) {
          measured_trees_flip <- 9 - measured_trees_raw
          res_flip <- sum(!measured_trees_flip %in% baseline_alive)
          
          if (res_flip == 0) {
            message(paste("    -> ✅ Successfully flipped Plot", p, "for Trait:", trt))
            
            # Record the flag for the newly flipped Tree IDs
            new_flags <- tibble(
              Plot = p,
              Tree = measured_trees_flip,
              Trait = trt,
              # NEW: Explicitly state the TO and FROM tree positions
              Mirror_Flag = paste0("Mirrored data moved FROM Tree ", measured_trees_raw, 
                                   " TO Tree ", measured_trees_flip, " for Trait: ", trt)
            )
            
            mirror_flags <- bind_rows(mirror_flags, new_flags)
            
            # Apply permanent fix to the data
            fixed_data <- fixed_data %>%
              mutate(Tree = if_else(Plot == p & Trait == trt, 9 - Tree, Tree))
            
            trees_measured_this_age <- c(trees_measured_this_age, measured_trees_flip)
            plot_fixed_any <- TRUE
          } else {
            trees_measured_this_age <- c(trees_measured_this_age, measured_trees_raw)
          }
        } else {
          trees_measured_this_age <- c(trees_measured_this_age, measured_trees_raw)
        }
      }
      baseline_alive <- intersect(baseline_alive, unique(trees_measured_this_age))
    }
    if(plot_fixed_any) plots_fixed_count <- plots_fixed_count + 1
  }
  
  if (plots_fixed_count > 0) {
    message(paste("\n  -> Trait Reversal complete. Fixed instances in", plots_fixed_count, "problem plots."))
  }
  
  # Return both the cleaned data AND the flags
  return(list(data = fixed_data, flags = mirror_flags))
}

# # # # # # # # # # # # # # # # # # # # # # # 
# PART 1: LOAD TRANSLATION MAP & FOLDERS ####
# # # # # # # # # # # # # # # # # # # # # # # 

trait_path <- file.path(data_dir,"PPGTraits_UK.xlsm")

if (file.exists(trait_path)) {
  trait_map <- read_excel(trait_path, sheet = "StdTraits") %>%
    select(
      trait_code_FR = `Eng Desc`,  
      trait_code_DP = Trait,       
      xml_group     = Group,       
      xml_desc_base = Description, 
      xml_units     = Units        
    ) %>%
    mutate(
      trait_code_FR = as.character(trait_code_FR),
      trait_code_DP = as.character(trait_code_DP),
      xml_group     = as.character(xml_group),
      xml_desc_base = as.character(xml_desc_base),
      xml_units     = as.character(xml_units)
    ) %>%
    filter(!is.na(trait_code_FR), !is.na(trait_code_DP)) %>%
    arrange(desc(nchar(trait_code_FR)))
  
} else {
  stop(paste("Error: Trait file not found at:", trait_path))
}

all_dirs <- list.dirs(path = file.path(data_dir,"High GCA Fullsib P85-P87 experiments"), recursive = FALSE, full.names = FALSE)

experiments_to_process <- setdiff(all_dirs, c("00_Scripts", "Archive", ".git", ".Rproj.user"))
message(paste("Found", length(experiments_to_process), "folders to check."))

# or do single 
experiments_to_process <-c("Craigellachie 49")

# # # # # # # # # # # # # # # # # #
# PART 2: MAIN PROCESSING LOOP #########################
# # # # # # # # # # # # # # # # # # 

root_path<-file.path(data_dir,"High GCA Fullsib P85-P87 experiments")

for (curr_exp in experiments_to_process) {
  
  tryCatch({
    message(paste("\n=========================================="))
    message(paste("Processing Folder:", curr_exp))
    
    exp_path <- file.path(root_path, curr_exp)
    file_prefix <- str_replace_all(curr_exp, " ", "_")
    
    # Initialize variables
    genetic_info <- NULL
    spatial_info <- NULL
    exp_data_long <- NULL
    
    # 
    # STEP 1: LOAD RAW FILES & DESIGN ####
    # 
    # NEW: Dynamically hunt for either .csv or .xlsx
    ascii_files <- fs::dir_ls(exp_path, regexp = paste0("(?i)", file_prefix, ".*_ASCII\\.(csv|xlsx)$"))
    
    if (length(ascii_files) == 0) {
      message(paste("  -> SKIPPING: Measurement file not found for:", curr_exp))
      next 
    }
    
    # Check if multiple files exist and warn the user
    if (length(ascii_files) > 1) {
      message("  -> WARNING: Multiple ASCII measurement files found in folder!")
      message(paste("  -> Found:", paste(basename(ascii_files), collapse = " AND ")))
      meas_path <- ascii_files[1]
      message(paste("  -> ACTION: Decided to load ->", basename(meas_path)))
    } else {
      meas_path <- ascii_files[1] 
      message(paste("  -> Loading Measurement File:", basename(meas_path)))
    }
    
    # --- Load Design ####
    # Look for .txt, .xlsx, or extensionless/dot-only design files
    design_files <- fs::dir_ls(exp_path, regexp = "(?i)_DF(\\.txt|\\.xlsx|\\.)?$")
    
    if (length(design_files) > 0) {
      # Prioritize .xlsx files if multiple exist
      xlsx_files <- design_files[str_detect(design_files, "(?i)\\.xlsx$")]
      
      if (length(xlsx_files) > 0) {
        design_path <- xlsx_files[1] # Grab the xlsx
      } else {
        design_path <- design_files[1] # Fallback to txt, dot, or blank
      }
      message(paste("  -> Loading Design:", basename(design_path)))
      
      # Route to the correct parser based on file extension
      if (str_detect(design_path, "(?i)\\.xlsx$")) {
        genetic_info <- parse_xlsx_design_file(design_path, file_prefix) 
      } else {
        genetic_info <- parse_long_design_file(design_path, file_prefix) 
      }
      
      if(!is.null(genetic_info)) {
        # Calculate block sizes and fill in missing blocks ("0")
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
      }
    } else {
      message("  -> NO DESIGN FILE FOUND. Proceeding without genetic info.")
    }
    
    # --- Load Spatial Matrix ####
    matrix_files <- dir_ls(exp_path, regexp = "(?i)_Matrix\\.csv$")
    if (length(matrix_files) > 0) {
      # Safely read the matrix as characters to prevent "Filler" from turning into NA
      raw_matrix <- read.csv(matrix_files[1], header = FALSE, colClasses = "character")
      raw_matrix[is.na(raw_matrix)] <- "0"
      
      spatial_info <- prowppos(raw_matrix)
      if (!is.null(spatial_info)) {
        spatial_info <- spatial_info %>%
          mutate(
            Plot = as.character(Plot), # Keep as character
            Tree = suppressWarnings(as.numeric(Tree)),
            Prow = suppressWarnings(as.numeric(Prow)),
            Ppos = suppressWarnings(as.numeric(Ppos))
          ) %>% 
          filter(!is.na(Plot))
        
        # --- NEW: Extract Fillers and inject them into genetic_info ---
        filler_plots <- spatial_info %>% 
          filter(str_detect(Plot, "(?i)Filler")) %>% 
          distinct(Plot)
        
        spatial_info <- spatial_info %>% select(Plot, Tree, Prow, Ppos)
        
        if (nrow(filler_plots) > 0) {
          filler_addition <- filler_plots %>%
            mutate(
              Block = "0", 
              Family_name = paste0(file_prefix, "_Filler")
            )
          
          if (is.null(genetic_info)) {
            genetic_info <- filler_addition
          } else {
            genetic_info <- genetic_info %>% 
              mutate(Plot = as.character(Plot)) %>%
              bind_rows(filler_addition)
          }
        }
      }
    } else {
      # NEW: Warn the user if no matrix file exists!
      message("  -> WARNING: No spatial matrix file (_Matrix.csv) found. Proceeding without spatial info.")
    }
    
    # 
    # STEP 2: MEASUREMENTS & RENAMING ####
    # 
    if (str_detect(meas_path, "(?i)\\.csv$")) {
      raw_data <- read_csv(meas_path, show_col_types = FALSE) 
    } else {
      raw_data <- read_excel(meas_path) 
    }
    
    if(!is.null(genetic_info)) {
      # Ensure genetic_info Plot is character so it joins safely
      genetic_info <- genetic_info %>% mutate(Plot = as.character(Plot))
      
      raw_data <- raw_data %>% 
        mutate(Plot_Check = as.character(Plot)) %>%
        filter(Plot_Check %in% genetic_info$Plot) %>%
        select(-Plot_Check)
    }
    
    # Check for and resolve repeat measurements within the same age ####
    # Safely ensure a Remarks column exists so we can query it without error
    if (!"Remarks" %in% names(raw_data)) {
      raw_data <- raw_data %>% mutate(Remarks = NA_character_)
    }
    
    # Identify duplicates based on Plot, Tree, Assessment, and Year
    duplicate_check <- raw_data %>%
      group_by(Plot, InferredTreePosition, Assessment, `Assessment Year`) %>%
      mutate(Measure_Count = n()) %>%
      ungroup()
    
    # If any duplicates exist, warn the user and filter them
    if (any(duplicate_check$Measure_Count > 1)) {
      
      # Isolate the duplicated records
      dupes_only <- duplicate_check %>% filter(Measure_Count > 1)
      
      # Count unique tree/trait combinations that have duplicates
      num_dupes <- dupes_only %>% distinct(Plot, InferredTreePosition, Assessment, `Assessment Year`) %>% nrow()
      
      # Extract and format the specific traits affected
      affected_traits <- dupes_only %>% 
        mutate(Combined_Trait = paste0(str_to_upper(Assessment), "_", `Assessment Year`)) %>% 
        pull(Combined_Trait) %>% 
        unique() %>% 
        sort()
      
      # Print the rich diagnostic messages
      message(paste("    -> WARNING:", num_dupes, "instances of repeat measurements detected!"))
      message(paste("    -> Affected Traits:", paste(affected_traits, collapse = ", ")))
      message("    -> Prioritizing 'Direction SE' (or keeping the first entry if exact match)...")
      
      # Apply the deduplication
      raw_data <- duplicate_check %>%
        group_by(Plot, InferredTreePosition, Assessment, `Assessment Year`) %>%
        arrange(
          # Sort so that rows containing "Direction SE" (case-insensitive) jump to the top
          desc(str_detect(replace_na(Remarks, ""), "(?i)Direction SE"))
        ) %>%
        slice(1) %>% # Take the top row (either the SE one, or just the first one if neither match)
        ungroup() %>%
        select(-Measure_Count)
    } else {
      # No duplicates, just clean up the temporary column
      raw_data <- duplicate_check %>% select(-Measure_Count)
    }
    # 
    
    # --- Create base long format and apply spatial correction ####
    exp_data_long <- raw_data %>%
      select(Plot_Orig = Plot, Tree_Orig = InferredTreePosition, Value_Char = Measurement, TraitCode = Assessment, 
             UnitCode = Unit, Age = `Assessment Year`) %>%
      filter(!is.na(Plot_Orig)) %>%
      mutate(
        Plot = as.character(Plot_Orig), # <--- Keep character to save "Filler"
        Tree_Orig_Num = suppressWarnings(as.numeric(Tree_Orig)), 
        Value_Num = suppressWarnings(as.numeric(Value_Char)),
        Prefix = str_to_upper(str_trim(replace_na(TraitCode, ""))),
        Age = as.character(Age),
        Date = as.Date(NA), 
        Trait_Orig = paste0(Prefix, replace_na(UnitCode, ""), Age)
      ) %>%
    
      # 1. Find the absolute maximum planted trees for the plot
      group_by(Plot) %>%
      mutate(Max_Trees_Planted = max(Tree_Orig_Num, na.rm = TRUE)) %>% 
      # 2. Find the max trees measured for this specific assessment run
      group_by(Plot, Trait_Orig) %>%
      mutate(
        Trees_in_Subset = max(Tree_Orig_Num, na.rm = TRUE),
        # 3. Apply the mapping function to correct the tree position
        Tree = purrr::pmap_dbl(
          list(Tree_Orig_Num, Trees_in_Subset, Max_Trees_Planted),
          map_interior_trees
        )
      ) %>%
      ungroup() %>%
      # Clean up the temporary columns
      select(-Tree_Orig_Num, -Max_Trees_Planted, -Trees_in_Subset)
    
    # --- RENAMING LOGIC ####
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
      select(Plot, Tree, Trait, Value_Char, Value_Num, UnitCode, Trait_Orig, Age, Prefix, Date) 
    
    # --- STEP 2.4: GENERALIZED ADDITIONAL DATA (AD FILES) ####
    # Looks for any file matching *_AD_*.csv or *_AD_*.xlsx
    ad_files <- dir_ls(exp_path, regexp = "(?i)_AD_(\\d+)\\.(csv|xlsx)$")
    
    if (length(ad_files) > 0) {
      for (ad_path in ad_files) {
        # Extract the age from the filename (e.g., "10" from "Brecon_56_AD_10_resi.xlsx")
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
                Plot = suppressWarnings(as.character(Plot)),
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
    
    # Standardize trait cases and Fix Mirrored Traits ####
    exp_data_long <- exp_data_long %>% mutate(Trait = str_to_title(Trait))
    
    mirror_output <- fix_mirrored_traits_8x1(exp_data_long)
    exp_data_long <- mirror_output$data
    exp_mirror_flags <- mirror_output$flags # Save the flags for later!
    # 
    
    
    # --- STEP 2.5: SURVIVAL CALC ####
    if(!is.null(spatial_info)) {
      raw_ages <- unique(na.omit(exp_data_long$Age))
      unique_ages <- raw_ages[!raw_ages %in% c("1", "01")]
      sur_list <- list()
      
      for(curr_age in unique_ages) {
        t_name <- paste0("Sur_", curr_age)
        if(t_name %in% unique(exp_data_long$Trait)) next
        
        # 1. Create a clean list of alive trees 
        alive_trees <- exp_data_long %>% 
          # Check if tree has data at this 
          filter(Age == curr_age, !is.na(Value_Num)) %>% 
          
          # OPTIONAL: If your raw data explicitly uses '0' in a Value column to mean 'Dead', 
          # you must filter those out so they don't count as "data". Uncomment below:
          filter(Value_Num != 0) %>% 
          
          select(Plot, Tree) %>% 
          distinct() %>%
          mutate(is_alive = TRUE)
        
        # 2. Safely join against the spatial info
        sur_df <- spatial_info %>% 
          select(Plot, Tree) %>%
          left_join(alive_trees, by = c("Plot", "Tree")) %>%
          mutate(
            Trait = t_name, 
            Age = curr_age,
            # Replace any NAs (trees that didn't match) with FALSE
            is_alive = replace_na(is_alive, FALSE), 
            Value_Num = if_else(is_alive, 1, 0),
            Value_Char = as.character(Value_Num),
            UnitCode = "$1-1", 
            Trait_Orig = paste0("Calculated_", t_name), 
            Prefix = "SUR", 
            Date = as.Date(NA)
          ) %>% 
          select(-is_alive)
        
        sur_list[[t_name]] <- sur_df
      }
      if(length(sur_list) > 0) exp_data_long <- bind_rows(exp_data_long, bind_rows(sur_list))
    }
    
    #
    # STEP 3: FLAGS & LOGIC ####
    # 
    exp_data_long <- exp_data_long %>%
      mutate(
        Is_Target_For_Zero_Removal = str_detect(Trait, "(?i)^(Pil|Br|St)"),
        Value_Num = case_when(
          Is_Target_For_Zero_Removal & Value_Num == 0 ~ NA_real_,
          TRUE ~ Value_Num
        ),
        Value_Char = if_else(is.na(Value_Num), NA_character_, as.character(Value_Num))
      )
    
    data_with_outliers <- exp_data_long %>%
      select(-Is_Target_For_Zero_Removal) %>% 
      group_by(Trait) %>%
      mutate(
        n_unique = n_distinct(Value_Num, na.rm = TRUE),
        Is_Survival = str_detect(Trait, "(?i)Sur"),
        Is_Binary_Resi = str_detect(Trait, "(?i)(Resi|Drill|Amp|Den|Cull|Stem)") & n_unique <= 2,
        Is_Ordinal  = !Is_Survival & !Is_Binary_Resi & n_unique < 9, 
        Mean_Val = mean(Value_Num, na.rm = TRUE),
        SD_Val   = sd(Value_Num, na.rm = TRUE),
        Lower_Limit = Mean_Val - (4 * SD_Val), 
        Upper_Limit = Mean_Val + (4 * SD_Val),
        Flag_Outlier = case_when(
          !is.na(Value_Num) & Is_Ordinal & Value_Num == 0 ~ TRUE,
          Is_Ordinal & Value_Num == 0 ~ TRUE,
          (Is_Survival | Is_Binary_Resi) & Value_Num > 1 ~ TRUE, 
          !Is_Survival & !Is_Ordinal & !Is_Binary_Resi & !is.na(Value_Num) & 
            (Value_Num < Lower_Limit | Value_Num > Upper_Limit) ~ TRUE,
          TRUE ~ FALSE
        )
      ) %>%
      ungroup()
    
    # Map out the exact, adjacent age steps that occurred in the whole experiment
    global_age_seq <- data_with_outliers %>%
      filter(!is.na(Value_Num), str_detect(Trait, "_\\d+$")) %>%
      mutate(
        Base_Trait = str_remove(Trait, "_\\d+$"),
        Age_Num = suppressWarnings(as.numeric(str_extract(Trait, "\\d+$")))
      ) %>%
      filter(!is.na(Age_Num)) %>%
      distinct(Base_Trait, Age_Num) %>%
      arrange(Base_Trait, Age_Num) %>%
      group_by(Base_Trait) %>%
      mutate(Expected_Prev_Age = lag(Age_Num)) %>%
      ungroup()
    
    # Calculate Shrinkage Flags using the global sequence AND unit checks
    shrinkage_flags <- data_with_outliers %>%
      filter(!is.na(Value_Num), str_detect(Trait, "_\\d+$")) %>%
      filter(!Is_Ordinal, !Is_Binary_Resi) %>%
      filter(!str_detect(Trait, "(?i)(Sur|Resi|Detrended|Ampli|Resistance|Resist|Crack|Form|Br|St|Pil)")) %>%
      mutate(
        Base_Trait = str_remove(Trait, "_\\d+$"),
        Age_Num = suppressWarnings(as.numeric(str_extract(Trait, "\\d+$")))
      ) %>%
      filter(!is.na(Age_Num)) %>%
      left_join(global_age_seq, by = c("Base_Trait", "Age_Num")) %>%
      arrange(Plot, Tree, Base_Trait, Age_Num) %>%
      group_by(Plot, Tree, Base_Trait) %>%
      mutate(
        Prev_Val = lag(Value_Num),
        Prev_Age = lag(Age_Num),
        Prev_Unit = lag(UnitCode), # <--- NEW: Grab the unit from the previous measurement
        
        # Strict check: adjacent ages AND matching units
        Is_Shrinkage = !is.na(Prev_Val) & !is.na(Expected_Prev_Age) & 
          (Prev_Age == Expected_Prev_Age) & 
          (UnitCode == Prev_Unit) &  # <--- NEW: Ensure units are identical
          (Value_Num < Prev_Val)
      ) %>%
      ungroup() %>%
      filter(Is_Shrinkage) %>%
      select(Plot, Tree, Trait, Prev_Age) %>%
      mutate(Shrinkage_Error = paste("Shrinkage detected (smaller than Age", Prev_Age, "with matching units)"))
    
    final_long_data <- data_with_outliers %>%
      left_join(shrinkage_flags %>% select(-Prev_Age), by = c("Plot", "Tree", "Trait")) %>%
      mutate(
        Reject_Flag = case_when(
          is.na(Value_Num) ~ NA_real_,
          Flag_Outlier ~ 1,
         !is.na(Shrinkage_Error) ~ 1,      # Turn shrinkage on or off with this line
          TRUE ~ 0
        )
      )
    
    # 
    # STEP 4: OUTPUT GENERATION (TREE LEVEL) ####
    # 
    final_long_dedup <- final_long_data %>%
      arrange(Plot, Tree, Trait, desc(Date)) %>%
      distinct(Plot, Tree, Trait, .keep_all = TRUE)
    
    # [UPDATED] Validation summary now includes specific shrinkage messages
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
      group_by(Plot, Tree) %>%
      summarise(Validation_record = paste(Error_Msg, collapse = " | "), .groups = "drop")
    
    alive_summary <- NULL
    sur_traits <- unique(final_long_dedup$Trait[str_detect(final_long_dedup$Trait, "(?i)Sur_")])
    if(length(sur_traits) > 0) {
      sur_lookup <- tibble(Trait_Name = sur_traits, Age_Num = as.numeric(str_extract(sur_traits, "\\d+")))
      target_row <- sur_lookup %>% filter(Age_Num == max(Age_Num, na.rm = TRUE)) %>% slice(1)
      if(nrow(target_row) > 0) {
        alive_summary <- final_long_dedup %>%
          filter(Trait == target_row$Trait_Name) %>%
          select(Plot, Tree, Alive = Value_Num) %>%
          distinct()
      }
    }
    
    data_wide_clean <- final_long_dedup %>% 
      select(Plot, Tree, Trait, Value_Char) %>% 
      pivot_wider(names_from = Trait, values_from = Value_Char)
    
    exp_data_wide_numeric <- final_long_dedup %>% 
      select(Plot, Tree, Trait, Value_Num) %>% 
      pivot_wider(names_from = Trait, values_from = Value_Num)
    
    if (!is.null(spatial_info)) {
      exp_data_wide_numeric <- left_join(exp_data_wide_numeric, spatial_info, by = c("Plot", "Tree"))
    }
    
    flags_wide <- final_long_dedup %>% 
      select(Plot, Tree, Trait, Reject_Flag) %>% 
      pivot_wider(names_from = Trait, values_from = Reject_Flag, names_glue = "{Trait}_reject")
    
    final_wide_with_flags <- left_join(data_wide_clean, flags_wide, by = c("Plot", "Tree"))
    final_wide_with_flags <- left_join(final_wide_with_flags, validation_summary, by = c("Plot", "Tree"))
    
    if(!is.null(alive_summary)) {
      final_wide_with_flags <- left_join(final_wide_with_flags, alive_summary, by = c("Plot", "Tree"))
    }
    
    # Genetic info is Plot-level only
    if (!is.null(genetic_info)) final_wide_with_flags <- left_join(final_wide_with_flags, genetic_info, by = "Plot")
    if (!is.null(spatial_info)) final_wide_with_flags <- left_join(final_wide_with_flags, spatial_info, by = c("Plot", "Tree"))
    
    # Explicitly define the exact metadata order you requested
    metadata_cols <- c("Validation_record", "Alive", "Block", "SubBlock", "Family_name", "Prow", "Ppos")
    
    # Keep only the ones that actually exist in the dataframe (preserves your exact order)
    existing_meta_cols <- intersect(metadata_cols, names(final_wide_with_flags))
    
    # 1. Get all trait/reject columns (everything not Plot, Tree, or the metadata)
    raw_remaining <- setdiff(names(final_wide_with_flags), c("Plot", "Tree", metadata_cols))
    
    # 2. Build a sorting dataframe to parse and logically sort
    col_sorting_df <- tibble(col_name = raw_remaining) %>%
      mutate(
        # Identify if it's a reject column
        is_reject = str_detect(col_name, "(?i)_reject$"),
        # Strip the reject tag to get the core trait name + age
        base_full = str_remove(col_name, "(?i)_reject$"),
        # Extract the trailing numbers as a REAL numeric value
        age_num = suppressWarnings(as.numeric(str_extract(base_full, "\\d+$"))),
        # Replace NA ages (traits without an age tag) with 0 so they sort to the front
        age_sort = replace_na(age_num, 0),
        # Extract just the letters/base of the trait
        base_trait = str_remove(base_full, "_\\d+$"),
        # Identify if it is a survival trait
        is_survival = str_detect(base_trait, "(?i)^Sur")
      ) %>%
      # 3. Sort logic:
      # - Measurements first, Survival last (is_survival: FALSE -> TRUE)
      # - Then by Age (youngest to oldest)
      # - Then alphabetically by trait (e.g., br 06 comes before dm 06)
      # - Finally, the trait itself before its reject flag (is_reject: FALSE -> TRUE)
      arrange(is_survival, age_sort, base_trait, is_reject)
    
    # Extract the properly ordered column names
    remaining_cols <- col_sorting_df$col_name
    
    # Combine and apply to the dataframe
    final_wide_with_flags <- final_wide_with_flags %>% 
      select(all_of(c("Plot", "Tree", existing_meta_cols, remaining_cols))) %>%
      arrange(Plot, Tree)
    
    # 
    # ---  MERGE: Add Mirror Flags to Validation Record --- ####
    # 
    # 1. Summarize the flags to one row per Plot/Tree FIRST
    exp_mirror_summary <- exp_mirror_flags %>% 
      mutate(Plot = as.character(Plot), Tree = as.numeric(Tree)) %>%
      group_by(Plot, Tree) %>%
      summarise(
        Combined_Mirror_Flags = paste(na.omit(unique(Mirror_Flag)), collapse = " | "), 
        .groups = "drop"
      )
    
    # 2. Join the summary to your wide dataset safely (1-to-1 match)
    final_wide_with_flags <- final_wide_with_flags %>%
      left_join(exp_mirror_summary, by = c("Plot", "Tree")) %>%
      mutate(
        Validation_record = case_when(
          !is.na(Combined_Mirror_Flags) & (is.na(Validation_record) | Validation_record == "") ~ Combined_Mirror_Flags,
          !is.na(Combined_Mirror_Flags) ~ paste(Validation_record, "|", Combined_Mirror_Flags),
          TRUE ~ Validation_record
        )
      ) %>%
      # Clean up the temporary column
      select(-Combined_Mirror_Flags)
    # ---- ##
    
    write_csv(final_wide_with_flags, file.path(exp_path, paste0(file_prefix, "_Full_Data_With_Flags.csv")), na = "")
    
    # PDF & Stats
    current_traits <- unique(final_long_dedup$Trait)
    prefixes_found <- unique(final_long_dedup$Prefix)
    
    try({
      pdf(file.path(exp_path, paste0(file_prefix, "_graphs.pdf")), width = 8, height = 11)
      generate_pdf_report(final_long_dedup, exp_data_wide_numeric, curr_exp, prefixes_found, current_traits)
      dev.off()
    })
    while(dev.cur() > 1) dev.off()
    
    stats_df <- final_long_dedup %>%
      filter(Reject_Flag == 0) %>%
      group_by(Trait) %>%
      summarise(N_Valid = sum(!is.na(Value_Num)), Mean = mean(Value_Num, na.rm=TRUE), 
                Std_Dev = sd(Value_Num, na.rm=TRUE), Min = min(Value_Num, na.rm=TRUE), Max = max(Value_Num, na.rm=TRUE),
                CV_Pct = (sd(Value_Num, na.rm=TRUE)/mean(Value_Num, na.rm=TRUE))*100) %>%
      mutate(across(where(is.numeric), ~ round(., 2)))
    
    write_csv(stats_df, file.path(exp_path, paste0(curr_exp, "_Stats.csv")), na = "")
    #
    # STEP 5: XML GENERATION ####
    #
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
      left_join(trait_map, by = c("Prefix" = "trait_code_FR")) %>%
      mutate(
        Group = case_when(
          !is.na(xml_group) & xml_group != "" ~ xml_group,
          str_detect(Prefix, "(?i)SUR") ~ "Health",
          TRUE ~ "Uncategorized"
        ),
        Desc_Base = coalesce(xml_desc_base, Prefix),
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
        Full_Description = paste0(Desc_Base, " Age ", Age, " (was ", Trait_Orig, ")", Date_String),
        XML_Tag = sprintf(
          '<trait group_name="%s" name="%s" trait_type="M" description="%s" data_type="N" is_solver_mappable="1" summary_rule="first_with_reject" units="%s" validate="none" />',
          Group, Trait, Full_Description, Real_Units
        )
      ) %>%
      arrange(Trait)
    
    xml_filename <- file.path(exp_path, paste0(file_prefix, "_Traits.xml"))
    
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
