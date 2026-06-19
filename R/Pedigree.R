# =====================================================================
# MASTER PEDIGREE PIPELINE CONFIGURATION
# =====================================================================
run_pedigree_builder <- function(base_dir, pending_dir, existing_dir, species_code, species_name, founders_file, controls_file, op_fam_file, has_existing_db) {
  
  cat("\n==========================================")
  cat("\nBuilding Pedigree For:", species_name)
  cat("\n==========================================\n")
  
  # --- 1. LOAD FOUNDERS, CONTROLS, AND OP FAMILIES ---
  prefix_low <- tolower(species_code)
  
  # 1. Load Founders and force headers to lowercase, then fix the specific ones the script needs
  founders <- read_csv(file.path(base_dir, "Pedigree", founders_file), show_col_types = FALSE) %>%
    janitor::clean_names() %>% 
    rename(LOCAT = locat, GEN = gen, PYR = pyr) %>% 
    mutate(Genotype_name = paste0(prefix_low, number))
  
  # 2. Load Controls
  controls <- read_csv(file.path(base_dir, "Pedigree", controls_file), show_col_types = FALSE)
  
  # 3. Load OP Families and force headers to standard snake_case
  op_families <- read_excel(file.path(base_dir, "Pedigree", op_fam_file)) %>% 
    janitor::clean_names() %>% 
    mutate(across(everything(), as.character)) %>% 
    rename(
      Family_name = family_name, 
      Mum_name = mum_name, 
      Mum_type = mum_type, 
      Dad_name = dad_name, 
      Dad_type = dad_type, 
      Fam_description = fam_description,
      LOCAT = locat 
    ) %>% 
    mutate(across(where(is.character), str_trim))
  
  # --- 2. THE PEDIGREE GENERATOR FUNCTION ---
  build_pedigree <- function(target_dir, founders, controls, op_families, species_name) {
    is_target_batch <- str_detect(target_dir, "(?i)Target|High GCA")
    
    # 1. EXTRACT CURATED TRIAL DATA (UPDATED TO TARGET NEW CSV OUTPUTS)
    trial_files <- dir_ls(target_dir, recurse = TRUE, regexp = "(?i)Full_Data_With_Flags\\.csv$")
    
    if (length(trial_files) == 0) {
      trial_data <- tibble(Family_name = character(), Trial_id = character())
    } else {
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
        }) %>% drop_na(Family_name)
    }
    
    # Safety fallback if files exist but lack Family names
    if (nrow(trial_data) == 0) {
      trial_data <- tibble(Family_name = character(), Trial_id = character())
    }
    
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
    
    grp_controls <- controls %>% 
      select(Group_name, Type, Description = Fam_description) %>% 
      distinct()
    
    groups_final <- bind_rows(
      grp_locat, grp_cb, grp_controls,
      tibble(Group_name = "Unknown", Species = species_name, Type = "Unknown", Description = "Dummy")
    ) %>% distinct(Group_name, .keep_all = TRUE) %>% mutate(Group_id = row_number())
    
    # 4. PROCESS OP FAMILIES
    fam_op <- op_families %>%
      mutate(Stage = 4, Dad_id = NA_character_) %>%
      select(Family_name, Mum_name, Mum_type, Dad_name, Dad_type, Fam_description, Stage, Dad_id)
    
    # 5. PROCESS CONTROLS AS FAMILIES
    fam_controls <- controls %>%
      mutate(
        Mum_name = Group_name, Mum_type = "G", 
        Dad_name = Group_name, Dad_type = "G", 
        Stage = 2, Dad_id = NA_character_     
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
          is_target_batch ~ "Target batch parents control pollinated",
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
  
  # --- 3. GENERATE RAW PENDING TABLES ---
  pending_tables <- build_pedigree(
    target_dir = pending_dir, 
    founders = founders, 
    controls = controls, 
    op_families = op_families,
    species_name = species_name
  )
  
  # --- 4. LOAD EXISTING DATABASE EXPORTS (WITH TOGGLE) ---
  if (has_existing_db) {
    cat("\nLoading Existing Database Exports...\n")
    
    # Updated to safely read the specific DMS files we established in diagnostics
    db_fams   <- tryCatch(read_excel(file.path(existing_dir, "DMS_fams.xlsx")), error = function(e) tibble(Family_name = character()))
    db_genos  <- tryCatch(read_excel(file.path(existing_dir, "DMS_genotypes.xlsx")), error = function(e) tibble(Genotype_name = character(), Ortet_lat = numeric(), Ortet_origin = character()))
    db_groups <- tryCatch(read_excel(file.path(existing_dir, "DMS_groups.xlsx")), error = function(e) tibble(Group_name = character()))
    
    # Extract clean vectors of names currently in the database safely
    db_fam_list   <- if("Family_name" %in% names(db_fams)) db_fams$Family_name else character(0)
    db_geno_list  <- if(length(names(db_genos)) > 0) db_genos[[grep("(?i)name", names(db_genos), value = TRUE)[1]]] else character(0)
    db_group_list <- if(length(names(db_groups)) > 0) db_groups[[grep("(?i)name", names(db_groups), value = TRUE)[1]]] else character(0)
    
  } else {
    cat("\nFirst Run Mode: No database exports to load. Bypassing filter...\n")
    
    # Create empty dataframes and vectors so downstream logic (like Section 8) doesn't break
    db_fams   <- tibble(Family_name = character())
    db_genos  <- tibble(Genotype_name = character(), Ortet_lat = numeric(), Ortet_origin = character())
    db_groups <- tibble(Group_name = character())
    
    db_fam_list   <- character(0)
    db_geno_list  <- character(0)
    db_group_list <- character(0)
  }
  
  # --- 5. THE TRUE ANTI-JOIN ---
  cat("\n--- BUILDING UPLOAD EXPORTS ---\n")
  
  # A. Families
  true_families_export <- pending_tables$families %>% 
    filter(!Family_name %in% db_fam_list)
  
  # B. Extract ALL parents required
  needed_mums_I <- true_families_export %>% filter(Mum_type == "I") %>% pull(Mum_name)
  needed_dads_I <- true_families_export %>% filter(Dad_type == "I") %>% pull(Dad_name)
  required_parents_I <- unique(c(needed_mums_I, needed_dads_I))
  
  needed_mums_G <- true_families_export %>% filter(Mum_type == "G") %>% pull(Mum_name)
  needed_dads_G <- true_families_export %>% filter(Dad_type == "G") %>% pull(Dad_name)
  required_parents_G <- unique(c(needed_mums_G, needed_dads_G))
  
  # C. Build Verified Genotypes
  true_genotypes_export <- pending_tables$genotypes %>%
    filter(Genotype_name %in% required_parents_I) %>%
    filter(!Genotype_name %in% db_geno_list) %>%
    distinct(Genotype_name, .keep_all = TRUE)
  
  # D. Build Verified Groups
  true_groups_export <- pending_tables$groups %>%
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
  # Groups: Drop Group_id
  write_csv(true_groups_export %>% select(-Group_id), 
            file.path(base_dir, "Pedigree", paste0("Verified_", species_code, "_Groups_Import.csv")))
  
  # Genotypes: Drop Genotype_id and Family_id
  write_csv(true_genotypes_export %>% select(-Genotype_id, -Family_id), 
            file.path(base_dir, "Pedigree", paste0("Verified_", species_code, "_Genotypes_Import.csv")))
  
  # Families: Drop Family_id, Mum_id, and Dad_id
  write_csv(true_families_export %>% select(-Family_id, -Mum_id, -Dad_id), 
            file.path(base_dir, "Pedigree", paste0("Verified_", species_code, "_Families_Import.csv")))
  
  # --- 8. GENERATE COMPLETE FAMILIES ORIGIN FILE ---
  cat("\n--- CALCULATING COMPLETE FAMILY ORIGINS (MACRO-REGIONS & OPCB) ---\n")
  
  db_fams_char <- db_fams %>% mutate(across(everything(), as.character))
  universal_families <- bind_rows(
    pending_tables$families %>% mutate(across(everything(), as.character)),
    db_fams_char
  ) %>% distinct(Family_name, .keep_all = TRUE)
  
  db_genos_char <- db_genos %>% mutate(across(everything(), as.character))
  universal_genotypes <- bind_rows(
    pending_tables$genotypes %>% mutate(across(everything(), as.character)),
    db_genos_char
  ) %>% distinct(Genotype_name, .keep_all = TRUE)
  
  ro_families <- universal_families %>%
    filter(!str_detect(Family_name, "(?i)Founders")) %>% 
    left_join(universal_genotypes %>% select(Genotype_name, Mum_lat = Ortet_lat, Mum_orig = Ortet_origin), 
              by = c("Mum_name" = "Genotype_name")) %>%
    left_join(universal_genotypes %>% select(Genotype_name, Dad_lat = Ortet_lat, Dad_orig = Ortet_origin), 
              by = c("Dad_name" = "Genotype_name")) %>%
    left_join(controls %>% select(Family_name, Control_Region = Region) %>% distinct(), 
              by = "Family_name") %>%
    mutate(
      Mum_region = if_else(!is.na(Mum_lat) & as.numeric(Mum_lat) < 54, "South", "North"),
      Dad_region = if_else(!is.na(Dad_lat) & as.numeric(Dad_lat) < 54, "South", "North"),
      Dad_region = if_else(Dad_type == "G" & str_detect(Dad_name, "\\+"), Mum_region, Dad_region),
      Mum_ro = if_else(Mum_type == "I", paste(Mum_region, Mum_orig, sep="_"), NA_character_),
      Dad_ro = case_when(
        Dad_type == "I" ~ paste(Dad_region, Dad_orig, sep="_"),
        Dad_type == "G" & str_detect(Dad_name, "\\+") ~ paste(Dad_region, "Unk", sep="_"),
        TRUE ~ NA_character_
      )
    )
  
  cp_ros <- na.omit(unique(c(ro_families$Mum_ro, ro_families$Dad_ro)))
  control_ros <- na.omit(unique(ro_families$Control_Region))
  all_unique_ros <- unique(c(cp_ros, control_ros))
  
  for(ro in all_unique_ros) {
    col_name <- paste0("Ro_", tolower(ro))
    ro_families <- ro_families %>%
      mutate(
        !!sym(col_name) := if_else(is.na(Control_Region),
                                   0 + 
                                     if_else(!is.na(Mum_ro) & Mum_ro == ro, 0.5, 0) +
                                     if_else(!is.na(Dad_ro) & Dad_ro == ro, 0.5, 0),
                                   0
        )
      ) %>%
      mutate(
        !!sym(col_name) := if_else(!is.na(Control_Region) & Control_Region == ro, 1, !!sym(col_name))
      )
  }
  
  # Diagnostic Investigation: Ro_north_unk Contributors
  north_unk_investigation <- ro_families %>%
    filter(Mum_ro == "North_Unk" | Dad_ro == "North_Unk") %>%
    mutate(
      missing_latitude = is.na(Mum_lat) & is.na(Dad_lat),
      primary_cause = case_when(
        Dad_type == "G" & str_detect(Dad_name, "\\+") ~ "OPCB Pollen Cloud (Rule-Forced)",
        Mum_orig == "Unk" | is.na(Mum_orig) | Dad_orig == "Unk" | is.na(Dad_orig) ~ "Source DB Origin is Missing/Unk",
        TRUE ~ "Other Unknown Causes"
      )
    ) %>%
    count(primary_cause, missing_latitude) %>%
    mutate(percentage = round((n / sum(n)) * 100, 1)) %>%
    arrange(desc(n))
  
  print(north_unk_investigation)
  
  # 8.5 Add Fillers and finalize the export dataframe
  families_origin_export <- ro_families %>%
    mutate(Ro_filler = if_else(str_detect(Family_name, "(?i)Filler"), 1, 0)) %>%
    select(Family_name, starts_with("Ro_")) %>%
    mutate(across(starts_with("Ro_"), ~replace_na(.x, 0)))
  
  write_csv(families_origin_export, file.path(base_dir, "Pedigree", paste0("Complete_", species_code, "_Families_Origin.csv")))
  
  cat("Generated mathematically complete origins for", nrow(families_origin_export), "families.\n")
  
  # Extract the specific families with North_Unk for manual review
  families_to_review <- ro_families %>%
    filter(Mum_ro == "North_Unk" | Dad_ro == "North_Unk") %>%
    mutate(
      primary_cause = case_when(
        Dad_type == "G" & str_detect(Dad_name, "\\+") ~ "OPCB Pollen Cloud",
        Mum_orig == "Unk" | is.na(Mum_orig) | Dad_orig == "Unk" | is.na(Dad_orig) ~ "Missing/Unk in Source DB",
        TRUE ~ "Other"
      )
    ) %>%
    select(
      Family_name,
      primary_cause,
      Mum_name, Mum_type, Mum_lat, Mum_orig, Mum_ro,
      Dad_name, Dad_type, Dad_lat, Dad_orig, Dad_ro
    ) %>%
    arrange(primary_cause, Family_name)
  
  write_csv(families_to_review, file.path(base_dir, "Pedigree", "HighGCA_Unknown_Origins_Review.csv"))
  cat("\nExported", nrow(families_to_review), "families to 'HighGCA_Unknown_Origins_Review.csv'\n")
  
  cat("\n==========================================")
  cat("\nPlotting Pedigree Networks...")
  cat("\n==========================================\n")
  
  # # # # # # # # # # # # # # # # # # # # # # # #
  # PLOT 1: VERIFIED PENDING PEDIGREE ONLY   ####
  # # # # # # # # # # # # # # # # # # # # # # # #
  
  families  <- read_csv(file.path(base_dir,"Pedigree", paste0("Verified_", species_code, "_Families_Import.csv")), show_col_types = FALSE)
  genotypes <- read_csv(file.path(base_dir,"Pedigree", paste0("Verified_", species_code, "_Genotypes_Import.csv")), show_col_types = FALSE)
  groups    <- read_csv(file.path(base_dir,"Pedigree", paste0("Verified_", species_code, "_Groups_Import.csv")), show_col_types = FALSE)
  
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
    labs(title = paste("Verified", species_code, "Pending Pedigree Network"), subtitle = "Ready for Upload")
  
  print(p1)
  
  # # # # # # # # # # # # # # # # # # # # # # # #
  # PLOT 2: ENTIRE PEDIGREE (Pending + DB)   ####
  # # # # # # # # # # # # # # # # # # # # # # # #
  
  db_fams_char <- db_fams %>% mutate(across(everything(), as.character))
  all_families <- bind_rows(
    pending_tables$families %>% mutate(across(everything(), as.character)), 
    db_fams_char
  ) %>% distinct(Family_name, .keep_all = TRUE)
  
  db_genos_char <- db_genos %>% mutate(across(everything(), as.character))
  all_genotypes <- bind_rows(
    pending_tables$genotypes %>% mutate(across(everything(), as.character)), 
    db_genos_char
  ) %>% distinct(Genotype_name, .keep_all = TRUE)
  
  db_group_col <- grep("(?i)name", names(db_groups), value = TRUE)[1]
  
  # Protect against db_groups being empty by explicitly defining the column
  if (is.na(db_group_col) || nrow(db_groups) == 0) {
    db_groups_clean <- tibble(Group_name = character())
  } else {
    db_groups_clean <- db_groups %>% 
      rename(Group_name = !!sym(db_group_col)) %>% 
      mutate(across(everything(), as.character))
  }
  
  all_groups <- bind_rows(
    pending_tables$groups %>% mutate(across(everything(), as.character)), 
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
    labs(title = paste("Complete", species_code, "Pedigree Network"), subtitle = "Existing Database + Target Extract")
  
  print(p2)
}