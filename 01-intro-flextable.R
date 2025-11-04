library(flextable)
library(systemfonts)
library(gfonts)
library(gdtools)
library(dplyr)

source(here("_init.R"))


ft <- summarizor(cars) |>
  as_flextable(sep_w = 0) |>
  color(
    i = ~ stat == "range",
    color = "pink"
  ) |>
  bold(j = 1) |>
  italic(j = 2, italic = TRUE)
ft


head(airquality) |>
  flextable() |>
  autofit()


ft <- with(palmerpenguins::penguins, table(species, island)) |>
  as_flextable()
ft


ft <- add_header_lines(ft, "Size measurements for adult foraging penguins near Palmer Station, Antarctica") |>
  italic(part = "header", i = 1) |>
  color(color = "#0099FC", part = "footer")
ft


# save_as_docx(ft, path = "output/ft.docx")
# save_as_pptx(ft, path = "output/ft.pptx")
# save_as_image(ft, path = "output/ft.png")
# print(ft, preview = "docx")


set_flextable_defaults(
  font.color = "#0099FC",
  border.color = "red",
  theme_fun = "theme_box"
)

dat <- data.frame(
  wool = c("A", "B"),
  L = c(44.56, 28.22),
  M = c(24, 28.77),
  H = c(24.56, 18.78)
)
flextable(dat)


set_flextable_defaults(
  font.size = 12, font.family = "Open Sans",
  font.color = "#333333",
  table.layout = "fixed",
  border.color = "gray",
  theme_fun = theme_booktabs,
  padding.top = 3, padding.bottom = 3,
  padding.left = 4, padding.right = 4
)


flextable(dat)
flextable(dat) |> autofit()
flextable(dat) |> set_table_properties(layout = "autofit")
set_flextable_defaults(decimal.mark = ",", digits = 3, big.mark = " ")

flextable(head(ggplot2::diamonds)) |>
  colformat_double() |>
  colformat_int(j = "price", suffix = "$") |>
  autofit()

flextable(head(cars)) |>
  colformat_double(digits = 0) |>
  autofit()

ft <- flextable(head(airquality))
ft <- colformat_int(
  x = ft,
  na_str = "N/A"
)
autofit(ft)


adsl <- dplyr::select(formatters::ex_adsl, AGE, SEX, COUNTRY, ARM)

ft <- summarizor(adsl, by = "ARM") |>
  as_flextable(
    sep_w = 0, separate_with = "variable",
    spread_first_col = TRUE
  ) |>
  align(i = ~ !is.na(variable), align = "left")
ft


prepend_chunks(ft, i = ~ is.na(variable), j = "stat", as_chunk("\t"))


library(palmerpenguins)

dat <- penguins |>
  select(species, island, ends_with("mm")) |>
  group_by(species, island) |>
  summarise(
    across(
      where(is.numeric),
      .fns = list(
        avg = ~ mean(.x, na.rm = TRUE),
        sd = ~ sd(.x, na.rm = TRUE)
      )
    ),
    .groups = "drop"
  ) |>
  rename_with(~ tolower(gsub("_mm_", "_", .x, fixed = TRUE)))

ft_pen <- flextable(dat) |>
  colformat_double() |>
  separate_header() |>
  theme_vanilla() |>
  align(align = "center", part = "all") |>
  valign(valign = "center", part = "header") |>
  autofit()
ft_pen


ft_pen <- labelizor(
  x = ft_pen,
  part = "header",
  labels = c("avg" = "Mean", "sd" = "Standard Deviation")
)
ft_pen


ft_pen <- labelizor(
  x = ft_pen,
  part = "header",
  labels = stringr::str_to_title
)
ft_pen


dat <- data.frame(
  wool = c("A", "B"),
  L = c(44.56, 28.22),
  M = c(24, 28.77),
  H = c(24.56, 18.78)
)


flextable(dat) |>
  fontsize(i = ~ wool %in% "A", size = 10) |>
  font(part = "all", fontname = "Inconsolata") |>
  color(part = "header", color = "#e22323", j = c("L", "M", "H")) |>
  bold(part = "header", j = c("L", "M")) |>
  italic(part = "all", j = "wool") |>
  highlight(i = ~ L < 30, color = "wheat", j = c("M", "H"))


ft <- flextable(dat) |>
  align(align = "center", part = "all") |>
  line_spacing(space = 2, part = "all") |>
  padding(padding = 6, part = "header")
ft


ft |>
  bg(bg = "black", part = "all") |>
  color(color = "white", part = "all") |>
  merge_at(i = 1:2, j = 1) |>
  valign(i = 1, valign = "bottom")


myft <- as.data.frame(matrix(runif(5 * 5), ncol = 5)) |>
  flextable() |>
  colformat_double() |>
  autofit() |>
  align(align = "center", part = "all") |>
  bg(bg = "black", part = "header") |>
  color(color = "white", part = "all") |>
  bg(bg = scales::col_numeric(palette = "viridis", domain = c(0, 1)))
myft


myft <- myft |>
  rotate(rotation = "tbrl", part = "header", align = "center") |>
  height(height = 1, unit = "cm", part = "header") |>
  hrule(rule = "exact", part = "header") |>
  align(align = "right", part = "header")
myft


library(officer)
big_border <- fp_border(color = "red", width = 2)
small_border <- fp_border(color = "gray", width = 1)

myft <- flextable(head(airquality))
myft <- border_remove(x = myft)
myft <- border_outer(myft, part = "all", border = big_border)
myft <- border_inner_h(myft, part = "all", border = small_border)
myft <- border_inner_v(myft, part = "all", border = small_border)
myft


myft2 <- border_remove(myft)

myft2 <- vline(myft2, border = small_border, part = "all")
myft2 <- vline_left(myft2, border = big_border, part = "all")
myft2 <- vline_right(myft2, border = big_border, part = "all")
myft2 <- hline(myft2, border = small_border)
myft2 <- hline_bottom(myft2, border = big_border)
myft2 <- hline_top(myft2, border = big_border, part = "all")
myft2


ft <- flextable(head(airquality))
ft <- add_header_row(ft,
  top = TRUE,
  values = c("measurements", "time"),
  colwidths = c(4, 2)
)
ft <- align(ft, i = 1, align = "center", part = "header")
ft <- width(ft, width = .75)


theme_booktabs(ft)
theme_alafoli(ft)
theme_vader(ft)
theme_box(ft)
theme_vanilla(ft)

my_theme <- function(x, ...) {
  x <- colformat_double(x, big.mark = "'", decimal.mark = ",", digits = 1)
  x <- set_table_properties(x, layout = "fixed")
  x <- border_remove(x)
  std_border <- fp_border(width = 1, color = "orange")
  x <- border_outer(x, part = "all", border = std_border)
  x <- border_inner_h(x, border = std_border, part = "all")
  x <- border_inner_v(x, border = std_border, part = "all")
  autofit(x)
}
my_theme(ft)

ft <- flextable(head(iris))
separate_header(ft)


library(palmerpenguins)

dat <- penguins |>
  select(species, island, ends_with("mm")) |>
  group_by(species, island) |>
  summarise(
    across(
      where(is.numeric),
      .fns = list(
        avg = ~ mean(.x, na.rm = TRUE),
        sd = ~ sd(.x, na.rm = TRUE)
      )
    ),
    .groups = "drop"
  )
dat

ft_pen <- flextable(dat) |>
  separate_header() |>
  align(align = "center", part = "all") |>
  theme_box() |>
  colformat_double(digits = 2) |>
  autofit()
ft_pen


ft <- flextable(head(airquality))
ft <- set_header_labels(ft,
  Solar.R = "Solar R (lang)",
  Temp = "Temperature (degrees F)", Wind = "Wind (mph)",
  Ozone = "Ozone (ppb)"
)
ft <- set_table_properties(ft, layout = "autofit", width = .8)
ft


ft <- add_header_row(
  x = ft, values = c("air quality measurements", "time"),
  colwidths = c(4, 2)
)
ft <- theme_box(ft)
ft


ft <- add_header_lines(ft,
  values = c(
    "this is a first line",
    "this is a second line"
  )
)
theme_box(ft)

