source(here("_init.R"))

AE_TABLE <- arrow::read_parquet(here("data", "AE-TABLE.parquet"))
CT_ARM <- arrow::read_parquet(here("data", "CT-ARM.parquet"))


path_to_template <- here("template", "template-02.docx")
sub_path <- file.path("output", "ae-table", "flextable-paginate")
dir.create(sub_path, showWarnings = FALSE, recursive = TRUE)
out_path <- file.path(sub_path, "AE-TABLE.docx")


# Prepare data in wide format
pivot_tab <- select(AE_TABLE, LABEL, AESOC, AEDECOD, ARM, agg_level, stat_str) |>
  tidyr::pivot_wider(
    names_from = "ARM",
    values_from = "stat_str",
    values_fill = "0 (0%)"
  )

# Create flextable with pagination
ft <- pivot_tab |>
    flextable(col_keys = c("LABEL", CT_ARM$ARM)) |>
    # Add superscript notation for percentages
    append_chunks(
      i = 1, j = -1, part = "header",
      as_sup(" (1)")
    ) |>
    # Add denominators to header
    append_chunks(
      i = 1, j = -1, part = "header",
      as_chunk(fmt_header_n(CT_ARM$denom, newline = TRUE))
    ) |>
    # Indent sub-level items
    prepend_chunks(
      j = "LABEL",
      i = ~ agg_level == 2,
      as_chunk("\t")
    ) |>
    # Apply variable labels
    labelizor(
      labels = c(LABEL = "", unlist(labelled::var_label(adae))),
      j = "LABEL"
    ) |>
    labelizor(
      labels = stringr::str_to_sentence,
      j = "LABEL"
    ) |>
    # Create compound header for first column
    mk_par(
      i = 1, j = 1, part = "header",
      as_paragraph(
        labelled::get_variable_labels(adae)$AESOC,
        "\n\t",
        labelled::get_variable_labels(adae)$AEDECOD
      )
    ) |>
    add_header_lines("Table 15.3: 1 AE by SOC/PT") |>
    add_footer_lines(as_paragraph(as_sup("(1)"), " n (%)")) |>
    align(j = -1, align = "right", part = "all") |>
    width(width = 1) |>
    width(width = 2, j = 1) |>
    # Apply pagination
    paginate(
      init = TRUE,
      hdr_ftr = TRUE,
      group = "AESOC",
      group_def = "rle"
    )


# Define section properties with header and footer
main_sect_prop <- prop_section(
  page_margins = page_mar(top = 1, bottom = 1, left = 1, right = 1),
  type = "nextPage",
  footer_default = block_list(
    fpar(
      run_word_field(field = "PAGE \\* MERGEFORMAT"),
      " on ",
      run_word_field(field = "NumPages \\* MERGEFORMAT"),
      " pages"
    )
  ),
  header_default = block_list(
    fpar("Adverse Events Table")
  )
)

# Create document and add paginated table
doc <- read_docx(path = path_to_template)
doc <- body_add_flextable(doc, ft, align = "center")
doc <- body_set_default_section(doc, main_sect_prop)
doc <- docx_set_settings(doc, even_and_odd_headers = FALSE)

# Save the document
print(doc, target = out_path)


sub_path <- file.path("output", "ae-table", "split-pages")
dir.create(sub_path, showWarnings = FALSE, recursive = TRUE)
out_path <- file.path(sub_path, "AE-TABLE.docx")


# Prepare data in wide format
pivot_tab <- select(AE_TABLE, LABEL, AESOC, AEDECOD, ARM, agg_level, stat_str) |>
  tidyr::pivot_wider(
    names_from = "ARM",
    values_from = "stat_str",
    values_fill = "0 (0%)"
  )

# Split data by System Organ Class
lpivot_tab <- split(pivot_tab, pivot_tab$AESOC)

# Create one flextable per SOC
lft <- lapply(lpivot_tab, function(x) {
  x |>
    flextable(col_keys = c("LABEL", CT_ARM$ARM)) |>
    append_chunks(
      i = 1, j = -1, part = "header",
      as_sup(" (1)")
    ) |>
    append_chunks(
      i = 1, j = -1, part = "header",
      as_chunk(fmt_header_n(CT_ARM$denom, newline = TRUE))
    ) |>
    padding(
      j = "LABEL",
      i = ~ agg_level == 2,
      padding.left = 12
    ) |>
    labelizor(
      labels = c(LABEL = "", unlist(labelled::var_label(adae))),
      j = "LABEL"
    ) |>
    labelizor(
      labels = stringr::str_to_sentence,
      j = "LABEL"
    ) |>
    mk_par(
      i = 1, j = 1, part = "header",
      as_paragraph(
        labelled::get_variable_labels(adae)$AESOC,
        "\n\t",
        labelled::get_variable_labels(adae)$AEDECOD
      )
    ) |>
    add_header_lines("Table 15.3: 1 AE by SOC/PT") |>
    add_footer_lines(as_paragraph(as_sup("(1)"), " n (%)")) |>
    align(j = -1, align = "right", part = "all") |>
    width(width = 1) |>
    width(width = 2, j = 1) |>
    # Apply pagination within each table
    paginate(
      init = TRUE,
      hdr_ftr = TRUE,
      group = "AESOC",
      group_def = "rle"
    )
})


# Define default section properties
main_sect_prop <- prop_section(
  page_margins = page_mar(top = 1, bottom = 1, left = 1, right = 1),
  type = "nextPage",
  footer_default = block_list(
    fpar(
      run_word_field(field = "PAGE \\* MERGEFORMAT"),
      " on ",
      run_word_field(field = "NumPages \\* MERGEFORMAT"),
      " pages"
    )
  )
)

# Initialize document
doc <- read_docx(path = path_to_template)

# Add each table in its own section
for (ft_name in names(lft)) {
  # Create section properties with SOC name in header
  sect_prop_example <- prop_section(
    page_margins = page_mar(top = 1, bottom = 1, left = 1, right = 1),
    type = "nextPage",
    header_default = block_list(
      fpar(ft_name)
    )
  )

  # Add table and close section
  doc <- body_add_flextable(doc, lft[[ft_name]], align = "center")
  doc <- body_end_block_section(
    doc,
    value = block_section(property = sect_prop_example)
  )
}

# Add final marker and set default section
doc <- body_add_par(doc, "END")
doc <- body_set_default_section(doc, main_sect_prop)
doc <- docx_set_settings(doc, even_and_odd_headers = FALSE)

# Save the document
print(doc, target = out_path)

