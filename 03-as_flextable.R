source(here("_init.R"))

ex_adsl <- formatters::ex_adsl

set_flextable_defaults(
  border.color = "#AAAAAA",
  font.family = "Arial",
  font.size = 10,
  padding = 3,
  line_spacing = 1.4
)

# Select relevant variables for demographics
adsl <- select(ex_adsl, AGE, SEX, COUNTRY, ARM)

# Extract variable labels from attributes
col_labels <- map_chr(adsl, function(x) attr(x, "label"))

# Create summary statistics by treatment arm
ft <- summarizor(adsl, by = "ARM") |>
  as_flextable(
    sep_w = 0,
    separate_with = "variable",
    spread_first_col = TRUE
  ) |>
  align(i = ~ !is.na(variable), align = "left") |>
  prepend_chunks(i = ~ is.na(variable), j = "stat", as_chunk("\t")) |>
  labelizor(
    j = "stat",
    labels = col_labels,
    part = "all"
  ) |>
  autofit() |>
  add_header_lines(
    c(
      "x.x: Study Subject Data",
      "x.x.x: Demographic Characteristics",
      "Table x.x.x.x: Demographic Characteristics - Full Analysis Set"
    )
  ) |>
  add_footer_lines("Source: ADSL DDMMYYYY hh:mm; Listing x.xx; SDTM package: DDMMYYYY")

ft


# Select specific statistics to display
summary_custom <- summarizor(
  adsl,
  by = "ARM",
  num_stats = c("range", "median_iqr")
)

as_flextable(summary_custom, spread_first_col = TRUE) |>
  autofit() |>
  labelizor(
    j = "stat",
    labels = c(
      AGE = "Age (years)",
      COUNTRY = "Country",
      SEX = "Sex"
    ),
    part = "all"
  )


# Compare different layouts
summary_data <- summarizor(adsl, by = "ARM")

# Layout 1: Spread groups across columns
ft1 <- as_flextable(
  summary_data,
  spread_first_col = TRUE,
  sep_w = 0
) |>
  autofit()

ft1


# Layout 2: Groups as rows
ft2 <- as_flextable(
  summary_data,
  spread_first_col = FALSE
) |>
  autofit() |>
  add_header_lines("Layout 2: Groups as rows")

ft2


# Create a simple frequency table
sex_table <- table(ex_adsl$SEX)

# Convert to flextable
as_flextable(sex_table) |>
  set_header_labels(value = "Sex", stat = "Count") |>
  autofit() |>
  add_header_lines("Distribution by Sex")


# Create a two-way contingency table
sex_arm_table <- table(
  Sex = ex_adsl$SEX,
  Treatment = ex_adsl$ARM
)

# Convert to flextable
as_flextable(sex_arm_table) |>
  autofit() |>
  add_header_lines("Sex Distribution by Treatment Arm") |>
  align(j = 1, align = "left", part = "body") |>
  align(j = -1, align = "center", part = "all")

