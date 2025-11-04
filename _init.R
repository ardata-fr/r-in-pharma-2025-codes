library(tidyverse)
library(arrow)
library(conflicted)
library(knitr)
library(doconv)
library(here)
library(knitr)

library(gdtools)
library(promises)
library(gfonts)
library(ggplot2)
library(systemfonts)
library(ragg)

library(officer)
library(flextable)

library(labelled)
library(formatters)
library(palmerpenguins)
library(pharmaverseadam)

gdtools::register_gfont("Open Sans")

if (!font_family_exists("Arial")) {
  systemfonts::register_font(
    "Arial",
    plain = "static/fonts/Arial.ttf",
    bold = "static/fonts/Arial Bold.ttf",
    italic = "static/fonts/Arial Italic.ttf",
    bolditalic = "static/fonts/Arial Bold Italic.ttf"
  )
}

set_flextable_defaults(
  font.family = "Arial",
  font.size = 11,
  padding = 3,
  table.layout = "autofit",
  digits = 2
)

theme_set(
  theme_minimal(
    base_family = "Arial"
  )
)

conflicts_prefer(dplyr::select)
conflicts_prefer(dplyr::filter)
conflicts_prefer(dplyr::lag)
conflicts_prefer(purrr::compose)
conflicts_prefer(palmerpenguins::penguins)
