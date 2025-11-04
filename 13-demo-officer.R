source(here("_init.R"))

path_to_template <- here("template", "template-01.docx")

doc <- read_docx(path = path_to_template)

table_styles <- styles_info(doc)

# Create "Table Caption" style for table titles
doc <- docx_set_paragraph_style(
  doc,
  base_on = "Normal",
  style_id = "TableCaption",
  style_name = "Table Caption",
  fp_p = fp_par(text.align = "center", padding.top = 12, padding.bottom = 3),
  fp_t = fp_text_lite(
    font.family = "Arial",
    italic = TRUE,
    font.size = 11,
    color = "#333333"
  )
)

# Create "Image Caption" style for figure titles
doc <- docx_set_paragraph_style(
  doc,
  base_on = "Normal",
  style_id = "ImageCaption",
  style_name = "Image Caption",
  fp_p = fp_par(text.align = "center", padding.top = 3, padding.bottom = 12),
  fp_t = fp_text_lite(
    font.family = "Arial",
    italic = TRUE,
    font.size = 11,
    color = "#333333"
  )
)

# Create "Graphic" style for centering graphics
doc <- docx_set_paragraph_style(
  doc,
  base_on = "Normal",
  style_id = "graphic",
  style_name = "Graphic",
  fp_p = fp_par(text.align = "center", padding.top = 3, padding.bottom = 3)
)

# Save the enriched template
print(doc, target = here("template", "template-02.docx"))


# doc <- read_docx(path = here("template", "template-02.docx"))
#
# # Test built-in heading styles
# doc <- body_add_par(doc, "Heading 1 Example", style = "heading 1")
# doc <- body_add_par(doc, "Heading 2 Example", style = "heading 2")
# doc <- body_add_par(doc, "Heading 3 Example", style = "heading 3")
# doc <- body_add_par(doc, "Heading 4 Example", style = "heading 4")
# doc <- body_add_par(doc, "Heading 5 Example", style = "heading 5")
#
# # Test our custom styles
# doc <- body_add_par(doc, "Table 1: Example Table Caption", style = "Table Caption")
# doc <- body_add_par(doc, "Figure 1: Example Image Caption", style = "Image Caption")
#
# # Preview in Word
# print(doc, preview = TRUE)


path_to_template <- here("template", "template-02.docx")

# Set flextable defaults
set_flextable_defaults(
  table.layout = "autofit", # Let Word optimize column widths
  font.family = "Arial", # Match our organization's standard
  font.size = 11, # Standard body text size
  digits = 2 # Two decimal places for numbers
)

# Set ggplot2 theme
theme_set(
  theme_minimal(
    base_family = "Arial" # Match font across tables and plots
  )
)


# Subject-level data
SL_TBLS <- pharmaverseadam::adsl |>
  select(USUBJID, RFSTDTC, RFENDTC, AGE, SEX) |>
  nest(.by = USUBJID, .key = "SL")

# Metabolic vital signs data
METABOLIC_TBLS <- pharmaverseadam::advs_metabolic |>
  dplyr::filter(
    !PARAMCD %in% c("HEIGHT", "WEIGHT", "BMI") # Exclude non-metabolic parameters
  ) |>
  select(USUBJID, AVISIT, ADY, AVAL, PARAMCD) |>
  drop_na() |>
  # Average multiple measurements per visit
  summarise(AVAL = mean(AVAL, na.rm = TRUE), .by = c(USUBJID, AVISIT, ADY, PARAMCD)) |>
  arrange(USUBJID, PARAMCD) |>
  nest(.by = USUBJID, .key = "METABOLIC")

# Combine both datasets
LIST_TBLS <- inner_join(METABOLIC_TBLS, SL_TBLS, by = "USUBJID")

# Define parameter labels for better readability
labels <- c(
  BMI = "Body Mass Index (kg/m2)",
  DIABP = "Diastolic Blood Pressure (mmHg)",
  HEIGHT = "Height (cm)",
  HIPCIR = "Hip Circumference (cm)",
  PULSE = "Pulse Rate (beats/min)",
  SYSBP = "Systolic Blood Pressure (mmHg)",
  TEMP = "Temperature (C)",
  WAISTHIP = "Waist to Hip Ratio",
  WEIGHT = "Weight (kg)",
  WSTCIR = "Waist Circumference (cm)"
)


saveRDS(LIST_TBLS, here("data/LIST_TBLS.RDS"))
saveRDS(labels, here("data/LABELS.RDS"))


sub_path <- here("output", "patient-profile", "simple")

# Create output directory
dir.create(sub_path, showWarnings = FALSE, recursive = TRUE)

# Loop through each patient
for (i in seq_len(nrow(LIST_TBLS))) {
  # Get patient ID and create filename
  USUBJID_STR <- LIST_TBLS[["USUBJID"]][i]

  # result file path
  file_basename <- paste0(USUBJID_STR, ".docx")
  fileout <- file.path(sub_path, file_basename)

  # Start with the template
  doc <- read_docx(path = path_to_template)

  # Add table of contents
  doc <- body_add_toc(doc)
  doc <- body_add_break(doc)

  ## Section 1: Subject-Level Analysis ----

  # Extract subject-level data for this patient
  SL_TBL <- LIST_TBLS[["SL"]][[i]]

  # Create a formatted table with proper labels
  tab1 <- as_flextable(SL_TBL, show_coltype = FALSE) |>
    labelizor(
      labels = pharmaverseadam::adsl |>
        labelled::get_variable_labels() |>
        unlist()
    ) |>
    padding(padding = 6)

  # Add section heading and table
  doc <- body_add_par(doc, value = "Subject Level Analysis", style = "heading 1")
  doc <- body_add_flextable(doc, tab1, align = "center")

  ## Section 2: Vital Signs Analysis ----

  # Extract metabolic data for this patient
  METABOLIC_TBL <- LIST_TBLS[["METABOLIC"]][[i]] |>
    mutate(ADY = as.Date(ADY))

  # Create time series plot
  gg <- ggplot(METABOLIC_TBL, aes(ADY, AVAL)) +
    geom_path() +
    scale_x_date(date_breaks = "4 weeks", date_labels = "%W") +
    facet_wrap(~PARAMCD, scales = "free_y") +
    labs(
      x = "Study Week",
      y = "Measurement Value",
      title = NULL
    )

  # Add section heading, subsection, and plot
  doc <- body_add_par(doc, value = "Vital Signs Analysis for Metabolic", style = "heading 1")
  doc <- body_add_par(doc, value = "Graphic", style = "heading 2")
  doc <- body_add_gg(doc, gg, style = "Graphic")

  # Create summary table (wide format)
  ft <- METABOLIC_TBL |>
    pivot_wider(
      id_cols = c(AVISIT),
      names_from = PARAMCD,
      values_from = AVAL
    ) |>
    flextable() |>
    theme_vanilla() |>
    colformat_double(digits = 2) |>
    labelizor(labels = labels)

  # Add subsection and table
  doc <- body_add_par(doc, value = "Table", style = "heading 2")
  doc <- body_add_flextable(doc, ft, align = "center")

  # Save the patient's report
  print(doc, target = fileout)
}


sub_path <- here("output", "patient-profile", "with-captions")

# Create output directory
dir.create(sub_path, showWarnings = FALSE, recursive = TRUE)

# Loop through each patient
for (i in seq_len(nrow(LIST_TBLS))) {
  # Get patient ID and create filename
  USUBJID_STR <- LIST_TBLS[["USUBJID"]][i]
  file_basename <- paste0(USUBJID_STR, ".docx")
  fileout <- file.path(sub_path, file_basename)

  # Start with the template
  doc <- read_docx(path = path_to_template)

  # Add table of contents for headings
  doc <- body_add_toc(doc)
  doc <- body_add_break(doc)

  # Add table of tables
  doc <- body_add_par(doc, "List of Tables", style = "heading 1")
  doc <- body_add_toc(doc, style = "Table Caption")
  doc <- body_add_break(doc)

  # Add table of figures
  doc <- body_add_par(doc, "List of Figures", style = "heading 1")
  doc <- body_add_toc(doc, style = "Image Caption")
  doc <- body_add_break(doc)

  ## Section 1: Subject-Level Analysis ----

  # Extract subject-level data for this patient
  SL_TBL <- LIST_TBLS[["SL"]][[i]]

  # Create a formatted table with proper labels
  tab1 <- as_flextable(SL_TBL, show_coltype = FALSE) |>
    labelizor(
      labels = pharmaverseadam::adsl |>
        labelled::get_variable_labels() |>
        unlist()
    ) |>
    padding(padding = 6)

  # Add section heading
  doc <- body_add_par(doc, value = "Subject Level Analysis", style = "heading 1")

  # Add table caption with automatic numbering
  doc <- body_add_fpar(
    doc,
    fpar(
      run_autonum(seq_id = "table", pre_label = "Table "),
      "Subject Demographics and Baseline Characteristics",
      fp_p = fp_par_lite(
        word_style = "Table Caption",
        keep_with_next = TRUE
      )
    )
  )

  # Add the table
  doc <- body_add_flextable(doc, tab1, align = "center")

  ## Section 2: Vital Signs Analysis ----

  # Extract metabolic data for this patient
  METABOLIC_TBL <- LIST_TBLS[["METABOLIC"]][[i]] |>
    mutate(ADY = as.Date(ADY))

  # Create time series plot
  gg <- ggplot(METABOLIC_TBL, aes(ADY, AVAL)) +
    geom_path() +
    scale_x_date(date_breaks = "4 weeks", date_labels = "%W") +
    facet_wrap(~PARAMCD, scales = "free_y") +
    labs(
      x = "Study Week",
      y = "Measurement Value",
      title = NULL
    )

  # Add section heading
  doc <- body_add_par(doc, value = "Vital Signs Analysis for Metabolic", style = "heading 1")

  # Add subsection
  doc <- body_add_par(doc, value = "Longitudinal Trends", style = "heading 2")

  # Add the plot
  doc <- body_add_gg(doc, gg, style = "Graphic")
  doc <- body_add_fpar(
    doc,
    fpar(
      run_autonum(seq_id = "fig", pre_label = "Figure "),
      "Metabolic Vital Signs Over Time by Parameter",
      fp_p = fp_par_lite(
        word_style = "Image Caption",
        keep_with_next = FALSE
      )
    )
  )

  # Create summary table (wide format)
  ft <- METABOLIC_TBL |>
    pivot_wider(
      id_cols = AVISIT,
      names_from = PARAMCD,
      values_from = AVAL
    ) |>
    flextable() |>
    theme_vanilla() |>
    colformat_double(digits = 2) |>
    labelizor(labels = labels)

  # Add subsection
  doc <- body_add_par(doc, value = "Summary by Visit", style = "heading 2")

  # Add table caption with automatic numbering
  doc <- body_add_fpar(
    doc,
    fpar(
      run_autonum(seq_id = "table", pre_label = "Table "),
      "Metabolic Vital Signs Summary by Study Visit",
      fp_p = fp_par_lite(
        word_style = "Table Caption",
        keep_with_next = TRUE
      )
    )
  )

  # Add the table
  doc <- body_add_flextable(doc, ft, align = "center")

  # Save the patient's report
  print(doc, target = fileout)
}

