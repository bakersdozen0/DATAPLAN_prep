


#### What's going on with Radnor 55 Cr_07? ####
## There's an overabundance of values ~200: 
## It's 199 at n=149
rd55<-read.csv(file.path(data_dir,"High GCA Fullsib P85-P87 experiments","Radnor 55","Radnor 55_Full_Data_With_Flags.csv"))

# Find the most frequent exact values in Cr_07
rd55 %>%
  count(Cr_07) %>%
  arrange(desc(n)) %>%
  head(10)

## test data and residuals of linear model for normalacy: 

# Load necessary libraries
library(ggplot2)
# install.packages("lme4") # Uncomment if you need to install it
library(lme4) 

# PART A: Test normality of raw Cr_07 data

# 1. Visual Check: Q-Q Plot
# If the data is normal, the points will hug the diagonal line tightly.
# The 199 spike will look like a horizontal flat line on this plot.
qqnorm(rd55$Cr_07, main = "Q-Q Plot: Raw Cr_07 Data")
qqline(rd55$Cr_07, col = "red", lwd = 2)

# 2. Statistical Check: Shapiro-Wilk Test
# Note: shapiro.test() fails if you have > 5000 rows. 
shapiro_raw <- shapiro.test(rd55$Cr_07)
print("Shapiro-Wilk Test for Raw Cr_07:")
print(shapiro_raw)

# PART B: Test normality of Model Residuals

# 1. Build the Linear Mixed-Effects Model
# Family_name is fixed; Prow and Ppos are random intercepts.
mod1 <- lmer(Cr_07 ~  + Family_name + (1 | Prow) + (1 | Ppos), data = rd55)

# 2. Extract the residuals
model_resids <- resid(mod1)

# 3. Visual Check: Q-Q Plot of Residuals
qqnorm(model_resids, main = "Q-Q Plot: Model Residuals")
qqline(model_resids, col = "blue", lwd = 2)

# 4. Statistical Check: Shapiro-Wilk on Residuals
shapiro_resids <- shapiro.test(model_resids)

print("Shapiro-Wilk Test for Model Residuals:")
print(shapiro_resids)


#### Correct Kielder 162 Ht_05 mirroring issue ####
## This trait is fully mirrored over the plot: compare patterns of dead tress between Ht_05 and Dm_10. 
library(tidyverse)

# 1. Point this directly at Kielder 162
target_csv <- "C:/Users/james.baker/Forest Research/TW CBC-TBA-NextGenBritishConifers - Share/Sitka/Backwards Selected Fullsib P96-P99 experiments/Kielder 162/Kielder_162_Full_Data_With_Flags.csv"

# Load the wide data
df <- read_csv(target_csv, show_col_types = FALSE)

# --- STORE "BEFORE" STATE FOR PLOTTING ---
df_before <- df %>% 
  select(Plot, Tree, Dm_09, Ht_05) %>% 
  mutate(State = "1. Before Fix (Mirrored Error)")

# 2. The critical fix: Use BOTH Plot and Tree as the unique identifiers!
spatial_map <- df %>%
  select(Plot, Tree, Prow, Ppos) %>%
  filter(!is.na(Prow) & !is.na(Ppos)) %>%
  distinct(Prow, Ppos, .keep_all = TRUE) 

# 3. Isolate the exact columns that need to be flipped
traits_to_flip <- intersect(names(df), c("Ht_05", "Sur_05", "Ht_05_reject"))

flip_data <- df %>%
  select(Plot, Tree, Prow, Ppos, all_of(traits_to_flip)) %>%
  filter(!is.na(Prow) & !is.na(Ppos)) %>%
  distinct(Plot, Tree, .keep_all = TRUE)

# 4. The Spatial Math: Find target coordinates
min_x <- min(flip_data$Ppos, na.rm = TRUE)
max_x <- max(flip_data$Ppos, na.rm = TRUE)

# Map the source data to the TARGET Plot/Tree at the mirrored destination
flip_mapping <- flip_data %>%
  mutate(Mirrored_Ppos = max_x - Ppos + min_x) %>%
  left_join(spatial_map %>% select(Target_Plot = Plot, Target_Tree = Tree, Prow, Ppos), 
            by = c("Prow" = "Prow", "Mirrored_Ppos" = "Ppos")) %>%
  filter(!is.na(Target_Plot))

# 5. Extract the data and assign it to the new Target Trees
fixed_traits <- flip_mapping %>%
  select(Plot = Target_Plot, Tree = Target_Tree, all_of(traits_to_flip)) %>%
  distinct(Plot, Tree, .keep_all = TRUE)

# 6. Merge the pristine data back together!
df_clean <- df %>% select(-all_of(traits_to_flip))

df_final <- df_clean %>%
  # Join securely using BOTH Plot and Tree
  left_join(fixed_traits, by = c("Plot", "Tree")) %>% 
  mutate(
    Validation_record = case_when(
      is.na(Validation_record) ~ "Ht_05 Globally Mirrored L-R",
      TRUE ~ paste(Validation_record, "| Ht_05 Globally Mirrored L-R")
    )
  )

# 7. Save the fixed dataset
out_path <- str_replace(target_csv, "(?i)\\.csv$", "_Corrected.csv")
write_csv(df_final, out_path, na = "")

message("Global Mirror applied flawlessly! Saved to: ", basename(out_path))

# 8. VERIFICATION PLOT
df_after <- df_final %>% 
  select(Plot, Tree, Dm_09, Ht_05) %>% 
  mutate(State = "2. After Fix (Corrected)")

plot_data <- bind_rows(df_before, df_after) %>%
  filter(!is.na(Dm_09) & !is.na(Ht_05))

p_verify <- ggplot(plot_data, aes(x = Dm_09, y = Ht_05)) +
  geom_point(alpha = 0.5, size = 1, color = "#2c3e50") +
  geom_smooth(method = "lm", formula = y ~ x, color = "#e74c3c", linetype = "dashed", se = FALSE) +
  facet_wrap(~State) +
  theme_bw() +
  labs(
    title = "Kielder 162: Ht_05 Global Spatial Correction",
    subtitle = "Verifying the Left-to-Right Field Mirror Fix",
    x = "Dm_09 (Trusted Baseline)",
    y = "Ht_05"
  ) +
  theme(
    strip.background = element_rect(fill = "#ecf0f1"),
    strip.text = element_text(face = "bold", size = 11)
  )

print(p_verify)

# 9. BEFORE & AFTER SPATIAL HEATMAP

# 1. Extract the BEFORE spatial data directly from the raw 'df'
spatial_before <- df %>% 
  select(Prow, Ppos, Value = Ht_05) %>% 
  mutate(State = "1. Before Fix (Mirrored Error)",
         Prow = as.numeric(Prow), 
         Ppos = as.numeric(Ppos)) %>%
  filter(!is.na(Value), !is.na(Prow), !is.na(Ppos))

# 2. Extract the AFTER spatial data directly from the corrected 'df_final'
spatial_after <- df_final %>% 
  select(Prow, Ppos, Value = Ht_05) %>% 
  mutate(State = "2. After Fix (Corrected)",
         Prow = as.numeric(Prow), 
         Ppos = as.numeric(Ppos)) %>%
  filter(!is.na(Value), !is.na(Prow), !is.na(Ppos))

# 3. Combine them into a single dataframe for faceting
heatmap_data <- bind_rows(spatial_before, spatial_after)

# 4. Generate the side-by-side Spatial Map using your custom styling
p_heatmap <- ggplot(heatmap_data, aes(x = Ppos, y = Prow, fill = Value)) + 
  geom_tile(color = "white", size = 0.1) + 
  scale_fill_distiller(palette = "Spectral", direction = -1, name = "Ht_05") + 
  scale_y_reverse() + 
  facet_wrap(~State) +
  labs(
    title = "Kielder 162: Ht_05 Spatial Heatmap Comparison",
    subtitle = "Visually verifying the left-to-right mirror correction",
    x = "Field Position (X)", 
    y = "Field Row (Y)"
  ) + 
  theme_minimal() + 
  coord_fixed() +
  theme(
    strip.background = element_rect(fill = "#ecf0f1"),
    strip.text = element_text(face = "bold", size = 11),
    legend.position = "right"
  )

# Display the heatmap in your RStudio Viewer
print(p_heatmap)


