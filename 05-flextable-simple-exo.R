source(here("_init.R"))

library(dplyr)
cancers <- arrow::read_parquet("data/cancers-2021.parquet") |>
  arrange(desc(effectif))
cancers


set_flextable_defaults(
  font.family = "Arial",
  big.mark = " ",
  decimal.mark = ",",
  table.layout = "fixed",
  post_process_all = function(z) {
    autofit(z)
  }
)

flextable(cancers)


ft <- flextable(cancers, col_keys = c("name", "prevalence", "effectif"))
ft


ft <- ft |>
  colformat_double(digits = 0, j = "effectif") |>
  colformat_double(digits = 2, j = "prevalence", suffix = " %")
ft


ft <- ft |>
  set_header_labels(name = "", prevalence = "Prevalence", effectif = "Number of cases") |>
  add_header_lines(c("Cancers", "Count | in France | all ages | all genders | 2021"),
                   top = TRUE) |>
  add_footer_lines("The counts represent the number of patients treated for each pathology (or chronic treatment or episode of care) in the group.")
ft


ft <- ft |>
  theme_vanilla() |>
  italic(italic = TRUE, part = "footer") |>
  autofit()
ft

