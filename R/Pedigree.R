# =====================================================================
# MASTER PEDIGREE PIPELINE CONFIGURATION
# =====================================================================

# 1. Base Paths & Species Identifiers
BASE_DIR      <- "//forestresearch.gov.uk/shares/CSFCC/Forest Resource and Product Assessment and Improvement/NRS-Tree Improvement/CONIFERS/SITKA SPRUCE/psi_DATAPLAN_prep"


# SS On Z: 
# "//forestresearch.gov.uk/shares/CSFCC/Forest Resource and Product Assessment and Improvement/NRS-Tree Improvement/CONIFERS/SITKA SPRUCE/psi_DATAPLAN_prep"
# SS on Teams: 
# "C:/Users/james.baker/Forest Research/TW CBC-TBA-NextGenBritishConifers - Share/Sitka"
# SP on Teams:
#"C:/Users/james.baker/Forest Research/TW CBC-TBA-NextGenBritishConifers - Share/Scots_Pine"



SPECIES_CODE  <- "SS"         # Options: "SS" or "SP"
SPECIES_NAME  <- "CBCSitka"   # Options: "CBCSitka" or "CBCScots"

# 2. Input File Names (Located in BASE_DIR/Pedigree/)
FOUNDERS_FILE <- paste0(SPECIES_CODE, "_tibdb_clones.csv") # Automatically looks for SS_ or SP_
CONTROLS_FILE <- "dataplan_family_control_import.csv" ## List of control families (usually listed in experimental preamble) with origin 
OP_FAM_FILE   <- "Cycle 1 OP.xlsx" # list of OP families and their origins 

# 3. Directories containing the processed trial data and live DB exports
TARGET_DIR    <- file.path(BASE_DIR, "High GCA Fullsib P85-P87 experiments")
DB_EXPORT_DIR <- file.path(BASE_DIR, "Backwards Selected Fullsib P96-P99 experiments")

# =====================================================================

library(tidyverse)
library(readxl)
library(fs)
library(janitor)
library(here)
library(purrr)
library(stringr)
library(igraph)
library(ggraph)

cat("\n==========================================")
cat("\nBuilding Pedigree For:", SPECIES_NAME)
cat("\n==========================================\n")

# --- 1. LOAD FOUNDERS, CONTROLS, AND OP FAMILIES ---
prefix_low <- tolower(SPECIES_CODE)

founders <- read_csv(file.path(BASE_DIR, "Pedigree", FOUNDERS_FILE), show_col_types = FALSE) %>%
  mutate(Genotype_name = paste0(prefix_low, number))

controls <- read_csv(file.path(BASE_DIR, "Pedigree", CONTROLS_FILE), show_col_types = FALSE)

op_families <- read_excel(file.path(BASE_DIR, "Pedigree", OP_FAM_FILE)) %>% 
  mutate(across(where(is.character), str_trim))

# --- 2. THE PEDIGREE GENERATOR FUNCTION ---
build_pedigree <- function(target_dir, founders, controls, op_families, species_name) {
  is_high_gca <- str_detect(target_dir, "(?i)High GCA")
  
  # 1. EXTRACT CURATED TRIAL DATA
  trial_files <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)_DP_ready\\.csv$")
  
  trial_data <- trial_files %>%
    map_df(function(file) {
      df <- tryCatch(read_csv(file, show_col_types = FALSE, col_types = cols(.default = col_character())), error = function(e) NULL)
      if (is.null(df)) return(NULL)
      fam_col <- grep("(?i)^family_name$", names(df), value = TRUE)
      if (length(fam_col) == 0) return(NULL)
      
      trial_name <- str_replace_all(str_extract(basename(file), "^[^_]+"), " ", "_")
      df %>% rename(Family_name = all_of(fam_col[1])) %>% select(Family_name) %>% distinct() %>% 
        mutate(
          Family_name = str_trim(Family_name),
          Family_name = str_replace(Family_name, "(?i)_OPST$", "_OPCB"),
          Trial_id = trial_name
        )
    }) %>% na.omit()
  
  # 2. MASTER PARENT METADATA
  raw_mums <- str_extract(trial_data$Family_name, "^[^_]+")
  raw_dads <- str_extract(trial_data$Family_name, "(?<=_).*")
  all_parents <- unique(c(raw_mums, raw_dads, op_families$Mum_name))
  
  parent_meta <- tibble(Genotype_name = all_parents) %>%
    drop_na() %>%
    left_join(founders, by = "Genotype_name") %>%
    mutate(Origin = case_when(GEN == "QC" ~ "HG", GEN == "WC" ~ "WC", TRUE ~ "Unk"))
  
  # 3. BUILD GROUPS 
  grp_locat <- parent_meta %>% filter(!is.na(LOCAT)) %>% select(LOCAT, Origin) %>% distinct() %>%
    mutate(Group_name = paste0(LOCAT, "_", Origin, "++"), Species = species_name, Type = "UKLR++", Description = paste("Selection group from", LOCAT, "Origin", Origin))
  
  grp_cb <- op_families %>% filter(!is.na(LOCAT)) %>% select(LOCAT) %>% distinct() %>%
    mutate(Group_name = paste0(LOCAT, "+"), Species = species_name, Type = "Clone Bank+", Description = paste("Open pollinated clone bank at", LOCAT))
  
  groups_final <- bind_rows(
    grp_locat, grp_cb, 
    controls %>% select(Group_name = Family_name, Type, Description = Fam_description),
    tibble(Group_name = "Unknown", Species = species_name, Type = "Unknown", Description = "Dummy")
  ) %>% distinct(Group_name, .keep_all = TRUE) %>% mutate(Group_id = row_number())
  
  # 4. PROCESS OP FAMILIES
  fam_op <- op_families %>%
    mutate(Stage = 4, Dad_id = NA_character_) %>%
    select(Family_name, Mum_name, Mum_type, Dad_name, Dad_type, Fam_description, Stage, Dad_id)
  
  # 5. PROCESS CONTROLS AS FAMILIES 
  fam_controls <- controls %>%
    mutate(
      Mum_name = "Unknown", Mum_type = "G",
      Dad_name = "Unknown", Dad_type = "G",
      Stage = 1, Dad_id = NA_character_
    ) %>%
    select(Family_name, Mum_name, Mum_type, Dad_name, Dad_type, Fam_description, Stage, Dad_id)
  
  # 6. PROCESS CP FAMILIES
  fam_cp <- trial_data %>%
    filter(str_detect(Family_name, "_"), !str_detect(Family_name, "(?i)iller")) %>%
    filter(!tolower(Family_name) %in% tolower(fam_op$Family_name)) %>%
    mutate(
      Mum_name = str_extract(Family_name, "^[^_]+"), Mum_type = "I",
      Raw_Dad = str_extract(Family_name, "(?<=_).*"),
      Is_OPCB = str_detect(Raw_Dad, "(?i)OPCB"),
      Dad_type = if_else(Is_OPCB, "G", "I"),
      Dad_name = case_when(Is_OPCB ~ "Unknown", TRUE ~ Raw_Dad),
      Dad_id = NA_character_,
      Stage = 4,
      Fam_description = case_when(
        Is_OPCB ~ paste("Open pollinated family from", Mum_name, "in unknown clone bank"),
        is_high_gca ~ "High GCA parents control pollinated",
        TRUE ~ paste("Control pollinated family", Mum_name, "x", Dad_name)
      )
    ) %>%
    select(Family_name, Mum_name, Mum_type, Dad_name, Dad_type, Fam_description, Stage, Dad_id) %>%
    distinct()
  
  # 7. ASSEMBLE ALL FAMILIES
  fam_founders <- grp_locat %>% mutate(Family_name = paste0(Group_name, "_Founders"), Mum_name = Group_name, Mum_type = "G", Dad_name = Group_name, Dad_type = "G", Fam_description = paste("Dummy family for founders in", Group_name), Stage = 2, Dad_id = NA_character_) %>% select(names(fam_op))
  fam_fillers <- tibble(Trial_id = unique(trial_data$Trial_id)) %>% mutate(Family_name = paste0(Trial_id, "_Filler"), Mum_name = "Unknown", Mum_type = "G", Dad_name = "Unknown", Dad_type = "G", Fam_description = paste("Fillers for trial", Trial_id), Stage = 3, Dad_id = NA_character_) %>% select(names(fam_op))
  
  families_combined <- bind_rows(fam_controls, fam_fillers, fam_founders, fam_op, fam_cp) %>% distinct(Family_name, .keep_all = TRUE) %>% arrange(Stage, Family_name) %>% mutate(Family_id = row_number())
  
  # 8. BUILD GENOTYPES 
  genotypes_final <- parent_meta %>%
    filter(!is.na(LOCAT)) %>% 
    mutate(Family_name = paste0(LOCAT, "_", Origin, "++_Founders"), Geno_description = paste("Backward selected founder", Genotype_name, "in", LOCAT), Ortet_locat = LOCAT, Ortet_pyr = PYR, Ortet_origin = Origin, Ortet_lat = if("lat" %in% names(.)) lat else NA, Ortet_long = if("long" %in% names(.)) long else NA, Ortet_ngr_status = if("Status" %in% names(.)) Status else NA, Ortet_prec = if("prec" %in% names(.)) prec else NA, Ortet_tavg = if("tavg" %in% names(.)) tavg else NA, Ortet_elev = if("Elevation" %in% names(.)) Elevation else NA) %>%
    left_join(families_combined %>% select(Family_name, Family_id, Mum_name, Mum_type, Dad_name, Dad_type), by = "Family_name") %>%
    mutate(Genotype_id = row_number()) %>% 
    select(Genotype_id, Genotype_name, Family_name, Family_id, Mum_name, Mum_type, Dad_name, Dad_type, Geno_description, Ortet_locat, Ortet_pyr, Ortet_origin, Ortet_lat, Ortet_long, Ortet_ngr_status, Ortet_prec, Ortet_tavg, Ortet_elev)
  
  # 9. FINAL ID LINKAGE
  families_final <- families_combined %>%
    mutate(
      Mum_id = case_when(
        Mum_type == "G" ~ as.character(groups_final$Group_id[match(Mum_name, groups_final$Group_name)]),
        Mum_type == "I" ~ as.character(genotypes_final$Genotype_id[match(Mum_name, genotypes_final$Genotype_name)]),
        TRUE ~ NA_character_
      ),
      Dad_id = case_when(
        !is.na(Dad_id) ~ as.character(Dad_id), 
        Dad_type == "G" ~ as.character(groups_final$Group_id[match(Dad_name, groups_final$Group_name)]), 
        Dad_type == "I" ~ as.character(genotypes_final$Genotype_id[match(Dad_name, genotypes_final$Genotype_name)]),
        TRUE ~ NA_character_
      )
    ) %>%
    mutate(Mum_id = replace_na(Mum_id, "Unknown"), Dad_id = replace_na(Dad_id, "Unknown")) %>%
    select(Family_id, Family_name, Mum_name, Mum_id, Mum_type, Dad_name, Dad_id, Dad_type, Fam_description, Stage)
  
  return(list(groups = groups_final, genotypes = genotypes_final, families = families_final))
}

# --- 3. GENERATE RAW CYCLE 1 TABLES ---
c1_tables <- build_pedigree(
  target_dir = TARGET_DIR, 
  founders = founders, 
  controls = controls, 
  op_families = op_families,
  species_name = SPECIES_NAME
)

# --- 4. LOAD ACTUAL DATABASE EXPORTS ---
db_fams   <- read_excel(file.path(DB_EXPORT_DIR, "DMS_all_fams.xlsx"))
db_genos  <- read_excel(file.path(DB_EXPORT_DIR, "DMS_all_genotypes.xlsx"))
db_groups <- read_excel(file.path(DB_EXPORT_DIR, "DMS_Groups.xlsx"))

# Extract clean vectors of names currently in the database
db_fam_list   <- db_fams[[grep("(?i)family.*name", names(db_fams), value = TRUE)[1]]]
db_geno_list  <- db_genos[[grep("(?i)name", names(db_genos), value = TRUE)[1]]]
db_group_list <- db_groups[[grep("(?i)name", names(db_groups), value = TRUE)[1]]]

# --- 5. THE TRUE ANTI-JOIN ---
cat("\n--- BUILDING TRUE UPLOAD FILES ---\n")

# A. Families
true_families_export <- c1_tables$families %>% 
  filter(!Family_name %in% db_fam_list)

# B. Extract ALL parents required by these specific families
needed_mums_I <- true_families_export %>% filter(Mum_type == "I") %>% pull(Mum_name)
needed_dads_I <- true_families_export %>% filter(Dad_type == "I") %>% pull(Dad_name)
required_parents_I <- unique(c(needed_mums_I, needed_dads_I))

needed_mums_G <- true_families_export %>% filter(Mum_type == "G") %>% pull(Mum_name)
needed_dads_G <- true_families_export %>% filter(Dad_type == "G") %>% pull(Dad_name)
required_parents_G <- unique(c(needed_mums_G, needed_dads_G))

# C. Build Verified Genotypes
true_genotypes_export <- c1_tables$genotypes %>%
  filter(Genotype_name %in% required_parents_I) %>%
  filter(!Genotype_name %in% db_geno_list) %>%
  distinct(Genotype_name, .keep_all = TRUE)

# D. Build Verified Groups
true_groups_export <- c1_tables$groups %>%
  filter(Group_name %in% required_parents_G | Group_name %in% true_families_export$Family_name) %>%
  filter(!Group_name %in% db_group_list) %>%
  distinct(Group_name, .keep_all = TRUE)

cat("Verified Groups to upload:    ", nrow(true_groups_export), "\n")
cat("Verified Genotypes to upload: ", nrow(true_genotypes_export), "\n")
cat("Verified Families to upload:  ", nrow(true_families_export), "\n")

# --- 6. ORPHAN CHECK ---
cat("\n--- ORPHAN CHECK ---\n")
orphans_I <- required_parents_I[!(required_parents_I %in% db_geno_list | required_parents_I %in% true_genotypes_export$Genotype_name)]
orphans_G <- required_parents_G[!(required_parents_G %in% db_group_list | required_parents_G %in% true_groups_export$Group_name)]

if(length(orphans_I) == 0 && length(orphans_G) == 0) {
  cat("SUCCESS! All Individual and Group parents are accounted for. Safe to upload.\n")
} else {
  if(length(orphans_I) > 0) {
    cat("\nWARNING: Missing INDIVIDUAL parents:\n"); print(orphans_I)
  }
  if(length(orphans_G) > 0) {
    cat("\nWARNING: Missing GROUP parents:\n"); print(orphans_G)
  }
}

# --- 7. EXPORT VERIFIED FILES ---
write_csv(true_groups_export, file.path(BASE_DIR, "Pedigree", "Verified_Cycle1_Groups_Import.csv"))
write_csv(true_genotypes_export, file.path(BASE_DIR, "Pedigree", "Verified_Cycle1_Genotypes_Import.csv"))
write_csv(true_families_export, file.path(BASE_DIR, "Pedigree", "Verified_Cycle1_Families_Import.csv"))

cat("\n==========================================")
cat("\nPlotting Pedigree Networks...")
cat("\n==========================================\n")

# # # # # # # # # # # # # # # # # # # # # # # 
# PLOT 1: VERIFIED CYCLE 1 PEDIGREE ONLY ####
# # # # # # # # # # # # # # # # # # # # # # # 

families  <- read_csv(file.path(BASE_DIR,"Pedigree", "Verified_Cycle1_Families_Import.csv"), show_col_types = FALSE)
genotypes <- read_csv(file.path(BASE_DIR,"Pedigree", "Verified_Cycle1_Genotypes_Import.csv"), show_col_types = FALSE)
groups    <- read_csv(file.path(BASE_DIR,"Pedigree", "Verified_Cycle1_Groups_Import.csv"), show_col_types = FALSE)

crosses <- families %>% filter(Stage == 4, Mum_name != "Unknown", Dad_name != "Unknown")
edges_mum <- crosses %>% select(from = Mum_name, to = Family_name)
edges_dad <- crosses %>% select(from = Dad_name, to = Family_name)

edges_group_geno <- genotypes %>%
  mutate(Group_name = paste0(Ortet_locat, "_", Ortet_origin, "++")) %>%
  select(from = Group_name, to = Genotype_name) %>% distinct()

edges_origin_group <- groups %>%
  filter(Type == "UKLR++", Origin != "Unknown", !is.na(Origin)) %>%
  select(from = Origin, to = Group_name) %>% distinct()

pedigree_edges <- bind_rows(edges_origin_group, edges_group_geno, edges_mum, edges_dad) %>% 
  filter(!is.na(from), !is.na(to))

all_nodes <- data.frame(name = unique(c(pedigree_edges$from, pedigree_edges$to))) %>%
  mutate(
    Node_Type = case_when(
      name %in% edges_origin_group$from ~ "1. Origin",
      name %in% edges_group_geno$from ~ "2. Group",
      name %in% edges_mum$from | name %in% edges_dad$from ~ "3. Genotype (Parent)",
      TRUE ~ "4. Family (Offspring)"
    )
  )

pedigree_graph <- graph_from_data_frame(d = pedigree_edges, vertices = all_nodes, directed = TRUE)

p1 <- ggraph(pedigree_graph, layout = 'sugiyama') + 
  geom_edge_diagonal(arrow = arrow(length = unit(1.5, 'mm')), end_cap = circle(3, 'mm'), alpha = 0.3, color = "gray50") +
  geom_node_point(aes(color = Node_Type), size = 3) +
  geom_node_text(aes(label = name), vjust = 1.5, hjust = 0.5, size = 2.5, repel = TRUE) +
  theme_void() + theme(legend.position = "bottom") +
  scale_color_manual(
    values = c("1. Origin" = "#E41A1C", "2. Group" = "#377EB8", "3. Genotype (Parent)" = "#4DAF4A", "4. Family (Offspring)" = "#984EA3"),
    name = "Pedigree Level"
  ) +
  labs(title = paste("Verified", SPECIES_CODE, "Cycle 1 Pedigree Network"), subtitle = "Ready for Upload")

print(p1)

# # # # # # # # # # # # # # # # # # # # # # # 
# PLOT 2: ENTIRE PEDIGREE (Cycle 1 + DB) ####
# # # # # # # # # # # # # # # # # # # # # # # 

db_fams_char <- db_fams %>% mutate(across(everything(), as.character))
all_families <- bind_rows(
  c1_tables$families %>% mutate(across(everything(), as.character)), 
  db_fams_char
) %>% distinct(Family_name, .keep_all = TRUE)

db_genos_char <- db_genos %>% mutate(across(everything(), as.character))
all_genotypes <- bind_rows(
  c1_tables$genotypes %>% mutate(across(everything(), as.character)), 
  db_genos_char
) %>% distinct(Genotype_name, .keep_all = TRUE)

db_group_col <- grep("(?i)name", names(db_groups), value = TRUE)[1]
db_groups_clean <- db_groups %>% 
  rename(Group_name = !!sym(db_group_col)) %>% 
  mutate(across(everything(), as.character))

all_groups <- bind_rows(
  c1_tables$groups %>% mutate(across(everything(), as.character)), 
  db_groups_clean
) %>% distinct(Group_name, .keep_all = TRUE)

crosses_all <- all_families %>% filter(Stage == 4, Mum_name != "Unknown", Dad_name != "Unknown", !is.na(Mum_name))
edges_mum_all <- crosses_all %>% select(from = Mum_name, to = Family_name)
edges_dad_all <- crosses_all %>% select(from = Dad_name, to = Family_name)

edges_group_geno_all <- all_genotypes %>%
  filter(!is.na(Ortet_locat)) %>%
  mutate(Group_name = paste0(Ortet_locat, "_", Ortet_origin, "++")) %>%
  select(from = Group_name, to = Genotype_name) %>% distinct()

edges_origin_group_all <- all_groups %>%
  filter(Type == "UKLR++", !is.na(Origin), Origin != "Unknown") %>%
  select(from = Origin, to = Group_name) %>% distinct()

pedigree_edges_all <- bind_rows(edges_origin_group_all, edges_group_geno_all, edges_mum_all, edges_dad_all) %>% 
  filter(!is.na(from), !is.na(to))

all_nodes_all <- data.frame(name = unique(c(pedigree_edges_all$from, pedigree_edges_all$to))) %>%
  mutate(
    Node_Type = case_when(
      name %in% edges_origin_group_all$from ~ "1. Origin",
      name %in% edges_group_geno_all$from ~ "2. Group",
      name %in% edges_mum_all$from | name %in% edges_dad_all$from ~ "3. Genotype (Parent)",
      TRUE ~ "4. Family (Offspring)"
    )
  )

pedigree_graph_all <- graph_from_data_frame(d = pedigree_edges_all, vertices = all_nodes_all, directed = TRUE)

p2 <- ggraph(pedigree_graph_all, layout = 'sugiyama') + 
  geom_edge_diagonal(arrow = arrow(length = unit(1.5, 'mm')), end_cap = circle(3, 'mm'), alpha = 0.3, color = "gray50") +
  geom_node_point(aes(color = Node_Type), size = 2) +
  geom_node_text(aes(label = name), vjust = 1.5, hjust = 0.5, size = 2, repel = TRUE) +
  theme_void() + theme(legend.position = "bottom") +
  scale_color_manual(
    values = c("1. Origin" = "#E41A1C", "2. Group" = "#377EB8", "3. Genotype (Parent)" = "#4DAF4A", "4. Family (Offspring)" = "#984EA3"),
    name = "Pedigree Level"
  ) +
  labs(title = paste("Complete", SPECIES_CODE, "Pedigree Network"), subtitle = "Live Database + Local Extracted Data")

print(p2)