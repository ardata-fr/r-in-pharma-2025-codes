source(here("_init.R"))

sub_path <- file.path("output", "officer-examples")
dir.create(sub_path, showWarnings = FALSE, recursive = TRUE)


read_docx() |>
  body_add_par(value = "Hello World!", style = "Normal") |>
  body_add_par(value = "Salut Bretons!", style = "centered") |>
  print(target = file.path(sub_path, "example_par.docx"))


read_docx() |>
  body_add_par(value = "This is a title 1", style = "heading 1") |>
  body_add_par(value = "This is a title 2", style = "heading 2") |>
  body_add_par(value = "This is a title 3", style = "heading 3") |>
  print(target = file.path(sub_path, "example_titles.docx"))


doc_toc <- read_docx() |>
  body_add_par("Table of Contents", style = "heading 1") |>
  body_add_toc(level = 2) |>
  body_add_par("Table of figures", style = "heading 1") |>
  body_add_toc(style = "Image Caption") |>
  body_add_par("Table of tables", style = "heading 1") |>
  body_add_toc(style = "Table Caption")

print(doc_toc, target = file.path(sub_path, "example_toc.docx"))


library(ggplot2)

gg <- ggplot(data = iris, aes(Sepal.Length, Petal.Length)) +
  geom_point()

doc_gg <- read_docx()
doc_gg <- body_add_gg(x = doc_gg, value = gg, style = "centered")


word_size <- docx_dim(doc_gg)
word_size
width <- word_size$page['width'] - word_size$margins['left'] - word_size$margins['right']
height <- word_size$page['height'] - word_size$margins['top'] - word_size$margins['bottom']

doc_gg <- body_add_gg(x = doc_gg, value = gg,
                      width = width, height = height,
                      style = "centered")

print(doc_gg, target = file.path(sub_path, "example_gg.docx"))


library(ggplot2)
library(flextable)

gg <- ggplot(data = iris, aes(Sepal.Length, Petal.Length)) +
  geom_point()

ft <- flextable(head(iris, n = 10))
ft <- set_table_properties(ft, layout = "autofit")

read_docx() |>
  body_add_par(value = "dataset iris", style = "heading 2") |>
  body_add_flextable(value = ft ) |>

  body_add_break() |>

  body_add_par(value = "plot examples", style = "heading 2") |>
  body_add_gg(value = gg, style = "centered") |>

  print(target = file.path(sub_path, "example_break.docx"))


doc <- read_docx()
styles <- styles_info(doc)
head(styles)


# Find all paragraph styles
para_styles <- styles |>
  filter(style_type == "paragraph") |>
  select(style_name, is_default)
head(para_styles, 10)


doc <- read_docx() |>
  body_add_par("This is a Heading 1", style = "heading 1") |>
  body_add_par("This is normal body text.", style = "Normal") |>
  body_add_par("This is a Heading 2", style = "heading 2") |>
  body_add_par("More body text here.", style = "Normal")

print(doc, target = file.path(sub_path, "example_styles.docx"))


# Read a template
doc <- read_docx()

# Create a custom "Table Caption" style
doc <- docx_set_paragraph_style(
  doc,
  base_on = "Normal",              # Base on existing style
  style_id = "TableCaption",       # Internal ID
  style_name = "Table Caption",    # Name used in R
  fp_p = fp_par(                   # Paragraph properties
    text.align = "center",
    padding.top = 12,
    padding.bottom = 3
  ),
  fp_t = fp_text_lite(             # Text properties
    font.family = "Arial",
    italic = TRUE,
    font.size = 11,
    color = "#333333"
  )
)


doc <- read_docx()

# Style for table captions
doc <- docx_set_paragraph_style(
  doc,
  base_on = "Normal",
  style_id = "TableCaption",
  style_name = "Table Caption",
  fp_p = fp_par(text.align = "center", padding.top = 12, padding.bottom = 3),
  fp_t = fp_text_lite(font.family = "Arial", italic = TRUE,
                      font.size = 11, color = "#333333")
)

# Style for figure captions
doc <- docx_set_paragraph_style(
  doc,
  base_on = "Normal",
  style_id = "ImageCaption",
  style_name = "Image Caption",
  fp_p = fp_par(text.align = "center", padding.top = 3, padding.bottom = 12),
  fp_t = fp_text_lite(font.family = "Arial", italic = TRUE,
                      font.size = 11, color = "#333333")
)

# Style for centered graphics
doc <- docx_set_paragraph_style(
  doc,
  base_on = "Normal",
  style_id = "Graphic",
  style_name = "Graphic",
  fp_p = fp_par(text.align = "center", padding.top = 3, padding.bottom = 3)
)

# Verify the styles were created
styles_info(doc) |>
  filter(style_name %in% c("Table Caption", "Image Caption", "Graphic")) |>
  select(style_name, style_type, is_custom)

