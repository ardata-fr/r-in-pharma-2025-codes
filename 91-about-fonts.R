library(tidyverse)
system_fonts_available <- systemfonts::system_fonts() |>
  dplyr::select(-path, -index) |>
  dplyr::arrange(family)

head(system_fonts_available, 20)


dat <- mtcars
dat$carname <- row.names(dat)

gg <- ggplot(dat, aes(drat, carname)) +
  geom_point() +
  theme_minimal(base_family = "Arial")
gg


gdtools::register_gfont("Lemon")


gg_lemon <- ggplot(dat, aes(drat, carname)) +
  geom_point() +
  theme_minimal(base_family = "Lemon")
gg_lemon

