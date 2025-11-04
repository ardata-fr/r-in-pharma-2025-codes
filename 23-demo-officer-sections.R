source(here("_init.R"))

LIST_TBLS <- readRDS("data/LIST_TBLS.RDS")
LABELS <- readRDS("data/LABELS.RDS")


path_to_template <- here("template", "template-02.docx")
sub_path <- here("output", "patient-profile", "default-section")
dir.create(sub_path, showWarnings = FALSE, recursive = TRUE)


# Portrait section with page numbering footer
sect_properties_1 <- prop_section(
  page_margins = page_mar(top = 1, bottom = 1, left = 1, right = 1),
  type = "oddPage",
  footer_default = block_list(
    fpar(
      run_word_field(field = "PAGE \\* MERGEFORMAT"),
      " on ",
      run_word_field(field = "NumPages \\* MERGEFORMAT"),
      " pages"
    )
  )
)

# Landscape section for wide content
sect_properties_2 <- prop_section(
  page_size = page_size(orient = "landscape"),
  page_margins = page_mar(top = 1, bottom = 1, left = 1, right = 1),
  type = "oddPage"
)


for (i in seq_len(nrow(LIST_TBLS))) {
  # Create output filename for this patient
  USUBJID_STR <- LIST_TBLS[["USUBJID"]][i]
  file_basename <- paste0(USUBJID_STR, ".docx")
  fileout <- file.path(sub_path, file_basename)

  # Initialize document with table of contents
  doc <- read_docx(path = path_to_template)
  doc <- body_add_toc(doc)
  doc <- body_add_break(doc)

  # Add subject-level analysis table
  SL_TBL <- LIST_TBLS[["SL"]][[i]]
  tab1 <- as_flextable(SL_TBL, show_coltype = FALSE) |>
    labelizor(
      labels = pharmaverseadam::adsl |>
        labelled::get_variable_labels() |>
        unlist()
    ) |>
    padding(padding = 6)

  doc <- body_add_par(doc, value = "Subject Level Analysis", style = "heading 1")
  doc <- body_add_flextable(doc, tab1, align = "center")

  # Prepare metabolic data and create visualization
  METABOLIC_TBL <- LIST_TBLS[["METABOLIC"]][[i]] |>
    mutate(ADY = as.Date(ADY))

  gg <- ggplot(METABOLIC_TBL, aes(ADY, AVAL)) +
    geom_path() +
    scale_x_date(date_breaks = "4 weeks", date_labels = "%W") +
    facet_wrap(~PARAMCD, scales = "free_y") +
    labs(x = "Week", y = "Value")

  doc <- body_add_par(doc, value = "Vital Signs Analysis for Metabolic", style = "heading 1")
  doc <- body_add_par(doc, value = "Graphic", style = "heading 2")
  doc <- body_add_gg(doc, gg, style = "Graphic")

  # Create summary table for metabolic data
  mb_ft <- METABOLIC_TBL |>
    pivot_wider(
      id_cols = AVISIT,
      names_from = PARAMCD,
      values_from = AVAL
    ) |>
    flextable() |>
    theme_vanilla() |>
    colformat_double(digits = 2) |>
    labelizor(labels = LABELS)

  # End portrait section and start landscape section for table
  doc <- body_end_block_section(doc, value = block_section(property = sect_properties_1))
  doc <- body_add_par(doc, value = "Table", style = "heading 2")
  doc <- body_add_flextable(doc, mb_ft, align = "center")
  doc <- body_end_block_section(doc, value = block_section(property = sect_properties_2))

  # Set default section and document settings
  doc <- body_set_default_section(doc, sect_properties_1)
  doc <- docx_set_settings(doc, even_and_odd_headers = FALSE)

  # Save the document
  print(doc, target = fileout)
}


path_to_template <- here("template", "template-02.docx")
sub_path <- file.path("output", "patient-profile", "even-odd-section")
dir.create(sub_path, showWarnings = FALSE, recursive = TRUE)


# Section with different footers for even and odd pages
sect_properties_1 <- prop_section(
  page_margins = page_mar(top = 1, bottom = 1, left = 1, right = 1),
  type = "oddPage",
  footer_default = block_list(
    fpar(
      run_word_field(field = "PAGE \\* MERGEFORMAT"),
      " on ",
      run_word_field(field = "NumPages \\* MERGEFORMAT"),
      " pages"
    )
  ),
  footer_even = block_list(
    fpar(
      "page ",
      run_word_field(field = "PAGE \\* MERGEFORMAT")
    )
  )
)

# Landscape section (reuses definition from previous example)
sect_properties_2 <- prop_section(
  page_size = page_size(orient = "landscape"),
  page_margins = page_mar(top = 1, bottom = 1, left = 1, right = 1),
  type = "oddPage"
)


for (i in seq_len(nrow(LIST_TBLS))) {
  # Create output filename for this patient
  USUBJID_STR <- LIST_TBLS[["USUBJID"]][i]
  file_basename <- paste0(USUBJID_STR, ".docx")
  fileout <- file.path(sub_path, file_basename)

  # Initialize document with table of contents
  doc <- read_docx(path = path_to_template)
  doc <- body_add_toc(doc)
  doc <- body_add_break(doc)

  # Add subject-level analysis table
  SL_TBL <- LIST_TBLS[["SL"]][[i]]
  tab1 <- as_flextable(SL_TBL, show_coltype = FALSE) |>
    labelizor(
      labels = pharmaverseadam::adsl |>
        labelled::get_variable_labels() |>
        unlist()
    ) |>
    padding(padding = 6)

  doc <- body_add_par(doc, value = "Subject Level Analysis", style = "heading 1")
  doc <- body_add_flextable(doc, tab1, align = "center")

  # Prepare metabolic data and create visualization
  METABOLIC_TBL <- LIST_TBLS[["METABOLIC"]][[i]] |>
    mutate(ADY = as.Date(ADY))

  gg <- ggplot(METABOLIC_TBL, aes(ADY, AVAL)) +
    geom_path() +
    scale_x_date(date_breaks = "4 weeks", date_labels = "%W") +
    facet_wrap(~PARAMCD, scales = "free_y") +
    labs(x = "Week", y = "Value")

  doc <- body_add_par(doc, value = "Vital Signs Analysis for Metabolic", style = "heading 1")
  doc <- body_add_par(doc, value = "Graphic", style = "heading 2")
  doc <- body_add_gg(doc, gg, style = "Graphic")

  # Create summary table for metabolic data
  mb_ft <- METABOLIC_TBL |>
    pivot_wider(
      id_cols = AVISIT,
      names_from = PARAMCD,
      values_from = AVAL
    ) |>
    flextable() |>
    theme_vanilla() |>
    colformat_double(digits = 2) |>
    labelizor(labels = LABELS)

  # End portrait section and start landscape section for table
  doc <- body_end_block_section(doc, value = block_section(property = sect_properties_1))
  doc <- body_add_par(doc, value = "Table", style = "heading 2")
  doc <- body_add_flextable(doc, mb_ft, align = "center")
  doc <- body_end_block_section(doc, value = block_section(property = sect_properties_2))

  # Set default section and enable even/odd footers
  doc <- body_set_default_section(doc, sect_properties_1)
  doc <- docx_set_settings(doc, even_and_odd_headers = TRUE)

  # Save the document
  print(doc, target = fileout)
}

