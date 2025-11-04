source(here("_init.R"))

# Load example datasets
adsl <- pharmaverseadam::adsl
adae <- pharmaverseadam::adae

adae <- adae |>
  dplyr::filter(
    # safety population
    SAFFL == "Y"
  )
adae


CT_ARM <- adsl |>
  semi_join(adae, by = "ARM") |>
  count(ARM, name = "denom")
CT_ARM


AE_TABLE_LEV2 <- adae |>
  group_by(ARM, AESOC, AEDECOD) |>
  summarise(
    n = n_distinct(USUBJID),
    .groups = "drop"
  ) |>
  arrange(AESOC, AEDECOD, ARM) |>
  tibble::add_column(agg_level = 2L)
AE_TABLE_LEV2


AE_TABLE_LEV1 <- adae |>
  group_by(ARM, AESOC) |>
  summarise(
    n = n_distinct(USUBJID),
    .groups = "drop"
  ) |>
  arrange(AESOC, ARM) |>
  tibble::add_column(agg_level = 1L)
AE_TABLE_LEV1


AE_TABLE_LEV0 <- adae |>
  group_by(ARM) |>
  summarise(
    n = n_distinct(USUBJID)
  ) |>
  arrange(ARM) |>
  tibble::add_column(agg_level = 0L, AESOC = "ANY ADVERSE EVENTS")
AE_TABLE_LEV0


AE_TABLE <- bind_rows(AE_TABLE_LEV0, AE_TABLE_LEV1, AE_TABLE_LEV2) |>
  left_join(CT_ARM, by = "ARM") |>
  mutate(
    pct = n / denom,
    denom = NULL,
    first_level = is.na(AEDECOD)
  ) |>
  arrange(AESOC, agg_level, first_level, AEDECOD, ARM) |>
  mutate(
    LABEL = coalesce(AEDECOD, AESOC),
    LABEL = factor(LABEL, levels = unique(LABEL)),
    stat_str = fmt_n_percent(n, pct, digit = 1)
  )
AE_TABLE


arrow::write_parquet(AE_TABLE, here("data", "AE-TABLE.parquet"))
arrow::write_parquet(CT_ARM, here("data", "CT-ARM.parquet"))


AE_TABLE_SAMPLE <- AE_TABLE |>
  filter(
    AESOC %in% c(
      "ANY ADVERSE EVENTS",
      "CARDIAC DISORDERS",
      "GASTROINTESTINAL DISORDERS"
    )
  )

x <- select(AE_TABLE_SAMPLE, LABEL, AESOC, AEDECOD, ARM, agg_level, stat_str) |>
  tidyr::pivot_wider(
    names_from = "ARM",
    values_from = c("stat_str"),
    values_fill = "0 (0%)"
  )
x


set_flextable_defaults(
  font.family = "Arial",
  font.size = 11,
  padding = 5,
  table.layout = "autofit"
)


ft <- x |>
  flextable(col_keys = c("LABEL", CT_ARM$ARM)) |>
  # Indent sub-level items
    prepend_chunks(
      j = "LABEL",
      i = ~ agg_level == 2,
      as_chunk("\t")
    ) |>
  add_header_lines("Table 15.3: 1 AE by SOC/PT") |>
  add_footer_lines(as_paragraph(as_sup("(1)"), " n (%)")) |>
  set_table_properties(layout = "fixed") |>
  width(width = 1) |>
  width(width = 2, j = 1)
ft


ft <- ft |>
  append_chunks(
    i = 1, j = -1, part = "header",
    as_sup(" (1)")
  ) |>
  append_chunks(
    i = 1, j = -1, part = "header",
    as_chunk(fmt_header_n(CT_ARM$denom, newline = TRUE))
  )
ft


ft <- ft |>
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
  align(j = -1, align = "right", part = "all")
ft


sub_path <- file.path("output", "ae-table", "flextable-ae-first")
dir.create(sub_path, showWarnings = FALSE, recursive = TRUE)
out_path <- file.path(sub_path, "AE-TABLE.docx")


save_as_docx(ft, path = out_path)

