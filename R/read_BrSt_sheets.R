library(here)
library(tidyxl)
library(dplyr)
library(tidyr)
library(purrr)
library(stringr)
library(readxl)
library(writexl)

process_kintyre_assessments <- function(layout_path) {
  
  message("--- Processing Kintyre Assessment Sheet ---")
  
  # Read the cells. Ensure you point this to the actual .xlsx file, not a CSV.
  cells <- xlsx_cells(layout_path, sheets = "plot layout assess sheet")
  
  # 1. Identify distinct rows containing block headers
  block_rows <- cells %>%
    filter(grepl("^BLOCK\\s*\\d+", character, ignore.case = TRUE)) %>%
    pull(row) %>%
    unique() %>%
    sort()
  
  if(length(block_rows) == 0) stop("No BLOCK anchors found.")
  
  # Define vertical row boundaries for each chunk of data
  end_rows <- c(tail(block_rows, -1) - 1, max(cells$row))
  field_data_list <- list()
  
  # 2. Iterate vertically through row chunks
  for (i in seq_along(block_rows)) {
    
    curr_block_row <- block_rows[i]
    curr_data_start <- curr_block_row + 2 # Skip header and sub-header rows
    curr_data_end   <- end_rows[i]
    
    # 3. Find all blocks horizontally within this specific row
    blocks_in_row <- cells %>%
      filter(row == curr_block_row, grepl("^BLOCK\\s*\\d+", character, ignore.case = TRUE)) %>%
      arrange(col)
    
    # Iterate across the horizontal blocks
    for (b in 1:nrow(blocks_in_row)) {
      block_col <- blocks_in_row$col[b]
      block_name <- blocks_in_row$character[b]
      
      # Determine horizontal boundary for this specific block
      next_block_col <- ifelse(b < nrow(blocks_in_row), blocks_in_row$col[b + 1], max(cells$col) + 1)
      
      # Isolate the metadata cells for just this block
      header_cells <- cells %>% 
        filter(row == curr_block_row, col >= block_col, col < next_block_col)
      
      # Extract Date
      date_anchor <- header_cells %>% filter(grepl("Date", character, ignore.case = TRUE))
      date_val <- if(nrow(date_anchor) > 0) {
        val_cell <- header_cells %>% filter(col == date_anchor$col[1] + 1)
        # Handle native Excel date vs character strings
        if(!is.na(val_cell$date[1])) format(as.Date(val_cell$date[1]), "%Y-%m-%d") else coalesce(val_cell$character[1], as.character(val_cell$numeric[1]))
      } else NA_character_
      
      # Extract Assessor
      assessor_anchor <- header_cells %>% filter(grepl("Assessor", character, ignore.case = TRUE))
      assessor_val <- if(nrow(assessor_anchor) > 0) {
        val_cell <- header_cells %>% filter(col == assessor_anchor$col[1] + 1)
        coalesce(val_cell$character[1], as.character(val_cell$numeric[1]))
      } else NA_character_
      
      # Get the column indices for "Plot" sub-headers
      plot_headers <- cells %>%
        filter(row == curr_block_row + 1, col >= block_col, col < next_block_col, grepl("Plot", character, ignore.case = TRUE)) %>%
        pull(col)
      
      # 4. Extract data for each Plot group
      for (p_col in plot_headers) {
        
        # Grab Plot (col), Br (col+1), and St (col+2)
        data_cells <- cells %>%
          filter(row >= curr_data_start, row <= curr_data_end, col %in% c(p_col, p_col + 1, p_col + 2)) %>%
          mutate(
            val = coalesce(as.character(numeric), character),
            col_type = case_when(
              col == p_col ~ "Plot",
              col == p_col + 1 ~ "Br",
              col == p_col + 2 ~ "St"
            )
          ) %>%
          select(row, col_type, val) %>%
          pivot_wider(names_from = col_type, values_from = val)
        
        # Ensure the columns exist even if blank, then filter out empty rows
        if(!"Plot" %in% names(data_cells)) data_cells$Plot <- NA_character_
        if(!"Br" %in% names(data_cells)) data_cells$Br <- NA_character_
        if(!"St" %in% names(data_cells)) data_cells$St <- NA_character_
        
        data_cells <- data_cells %>%
          filter(!is.na(Plot)) %>%
          mutate(
            Block = block_name,
            Date = date_val,
            Assessor = assessor_val
          )
        
        field_data_list[[length(field_data_list) + 1]] <- data_cells
      }
    }
  }
  
  # 5. Bind everything and perform final cleanup
  final_tidy_df <- bind_rows(field_data_list) %>%
    select(Block, Date, Assessor, Plot, Br, St) %>%
    mutate(
      Plot = as.numeric(Plot) # Convert back to numeric for easy joining later
    ) %>%
    arrange(Plot)
  
  return(final_tidy_df)
}

# Execution:
# tidy_assessments <- process_kintyre_assessments("path/to/your/Kintyre 18_BrSt_filled.xlsx")
# glimpse(tidy_assessments)


# Example Usage:
K18_Br_St<-here("Backwards Selected Fullsib P96-P99 experiments","Kintyre 18","Kintyre 18_BrSt_filled.xlsx")
K17_Br_St<-here("Backwards Selected Fullsib P96-P99 experiments","Kintyre 17","Kintyre 17_BrSt_filled.xlsx")

tidy_data17 <- process_kintyre_assessments (K17_Br_St)
tidy_data18 <- process_kintyre_assessments (K18_Br_St)
print(head(tidy_data))

write_xlsx(tidy_data17, here("Backwards Selected Fullsib P96-P99 experiments","Kintyre 17", "Kintyre_17_BrSt.xlsx"))
write_xlsx(tidy_data18, here("Backwards Selected Fullsib P96-P99 experiments","Kintyre 18", "Kintyre_18_BrSt.xlsx"))

####
### Histograms ####
library(dplyr)
library(tidyr)
library(ggplot2)

# 1. Clean and pivot the data
plot_data <- tidy_data %>%
  mutate(
    # Force characters like "Missing" or "Dead" to NA, keeping only numbers
    Br = suppressWarnings(as.numeric(Br)),
    St = suppressWarnings(as.numeric(St))
  ) %>%
  # Pivot to long format for easy faceting
  pivot_longer(cols = c(Br, St), names_to = "Score_Type", values_to = "Score") %>%
  # Drop NAs AND explicitly filter for valid scores to prevent empty 7-9 bins
  filter(!is.na(Score), Score >= 1, Score <= 6) 

# 2. Pre-calculate the histogram bars and percentages
# This makes plotting labels infinitely easier than doing it inside ggplot
bar_data <- plot_data %>%
  group_by(Assessor, Score_Type, Score) %>%
  summarise(count = n(), .groups = "drop_last") %>%
  mutate(
    total_trees = sum(count),
    density = count / total_trees,                   # Since binwidth=1, proportion == density
    percent_label = sprintf("%.1f%%", density * 100) # Format nicely as "XX.X%"
  ) %>%
  ungroup()

# 3. Calculate ideal normal distributions for each Assessor & Score_Type combo
normal_curves <- plot_data %>%
  group_by(Assessor, Score_Type) %>%
  summarise(
    mean_score = mean(Score, na.rm = TRUE),
    sd_score = sd(Score, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  # Create a sequence strictly from 1 to 6
  mutate(Score = list(seq(1, 6, length.out = 100))) %>%
  unnest(Score) %>%
  mutate(Normal_Density = dnorm(Score, mean = mean_score, sd = sd_score))

# 4. Plot the data
ggplot() +
  # Use geom_col instead of geom_histogram for total control
  geom_col(data = bar_data, 
           aes(x = Score, y = density), 
           fill = "#4682B4", color = "white", alpha = 0.8, width = 1) +
  
  # Add the percentage labels on top of the bars
  geom_text(data = bar_data, 
            aes(x = Score, y = density, label = percent_label), 
            vjust = -0.5, size = 3.5, color = "black") +
  
  # Add the calculated normal curve overlays
  geom_line(data = normal_curves, 
            aes(x = Score, y = Normal_Density), 
            color = "#B22222", linetype = "dotted", linewidth = 1) +
  
  # Facet by Score_Type and Assessor
  facet_grid(Score_Type ~ Assessor) +
  
  # Strictly enforce the x-axis to only show 1 through 6
  scale_x_continuous(breaks = 1:6, limits = c(0.5, 6.5)) +
  
  # Add 15% extra space to the top of the y-axis so labels don't clip through the facet ceiling
  scale_y_continuous(expand = expansion(mult = c(0, 0.15))) +
  
  theme_minimal() +
  labs(
    title = "Distribution of Br and St Scores by Assessor - Kintyre 18",
    x = "Ordinal Score (1-6)",
    y = "Density / Proportion"
  ) +
  theme(
    strip.text = element_text(face = "bold", size = 12),
    panel.grid.minor = element_blank() # Clean up the background grid
  )


#### chi-2 tests #### 

library(dplyr)
library(tidyr)
library(purrr)
library(broom) # Very helpful for turning statistical test outputs into tidy dataframes

# 1. Define the expected probabilities for scores 1 through 6
# 1 = 2%, 2 = 15%, 3 = 33%, 4 = 33%, 5 = 15%, 6 = 2%
expected_probs <- c(0.02, 0.15, 0.33, 0.33, 0.15, 0.02)

# 2. Calculate the observed counts and run the tests
chi_sq_results <- plot_data %>%
  # Filter to ensure we only have valid integer scores 1 to 6
  filter(Score %in% 1:6) %>%
  
  # Get the count of each score for each Assessor and Score_Type
  count(Assessor, Score_Type, Score) %>%
  
  # CRITICAL STEP: Ensure every score 1-6 is present for every group. 
  # If an assessor didn't give any '6's, this adds a row for Score 6 with n = 0.
  group_by(Assessor, Score_Type) %>%
  complete(Score = 1:6, fill = list(n = 0)) %>%
  
  # Make absolutely sure they are sorted 1 to 6 so they match expected_probs perfectly
  arrange(Assessor, Score_Type, Score) %>%
  
  # Run the chi-squared goodness-of-fit test for each group
  summarise(
    # chisq.test(x = observed_counts, p = expected_probabilities)
    test_result = list(chisq.test(x = n, p = expected_probs)),
    .groups = "drop"
  ) %>%
  
  # Unnest the test results into clean columns (statistic, p.value, parameter, etc.)
  mutate(tidied = map(test_result, tidy)) %>%
  unnest(tidied) %>%
  
  # Select the most useful columns for your final table
  select(Assessor, Score_Type, statistic, p.value) %>%
  
  # Optional: Add a significance flag for easy reading (p < 0.05)
  mutate(
    Significant_Deviation = ifelse(p.value < 0.05, "Yes", "No"),
    # Format p-value to prevent scientific notation if desired
    p.value = round(p.value, 4) 
  )

# 3. View the results
print(chi_sq_results)
