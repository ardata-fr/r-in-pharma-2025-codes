source(here("_init.R"))

sub_path <- file.path("output", "officer-sections")
dir.create(sub_path, showWarnings = FALSE, recursive = TRUE)


library(ggplot2)

gg <- ggplot(data = iris, aes(Sepal.Length, Petal.Length)) +
  geom_point()

# Define consistent margins
std_margins <- page_mar(top = 1, bottom = 1, left = 1, right = 1)

# Start document and add graphic
doc_section_1 <- read_docx()
doc_section_1 <- body_add_gg(
  x = doc_section_1, value = gg,
  width = 9, height = 6,
  style = "centered"
)

# Add section break to make previous content landscape
ps <- prop_section(
  page_size = page_size(orient = "landscape"),
  page_margins = std_margins,  # Use consistent margins
  type = "continuous"
)

doc_section_1 <- body_end_block_section(
  x = doc_section_1,
  value = block_section(property = ps)
)


doc_section_1 <- body_add_gg(
  x = doc_section_1, value = gg,
  width = 6.29, height = 9.72,
  style = "centered"
)

print(doc_section_1, target = file.path(sub_path, "example_landscape_gg.docx"))


# Define consistent properties
std_margins <- page_mar(top = 1, bottom = 1, left = 1, right = 1)

portrait_props <- prop_section(
  page_size = page_size(orient = "portrait"),
  page_margins = std_margins,
  type = "continuous"
)

landscape_props <- prop_section(
  page_size = page_size(orient = "landscape"),
  page_margins = std_margins,  # Same margins!
  type = "continuous"
)

doc_section_2 <- read_docx() |>
  # 1. Add portrait content
  body_add_par("This is a dummy text. It is in the default portrait section.") |>

  # 2. Close default section (starts special section)
  body_end_block_section(block_section(portrait_props)) |>

  # 3. Add landscape content
  body_add_gg(value = gg, width = 9, height = 6, style = "centered") |>

  # 4. Close landscape section (returns to default)
  body_end_block_section(block_section(landscape_props))

print(doc_section_2, target = file.path(sub_path, "example_landscape_gg2.docx"))


# Consistent margins
std_margins <- page_mar(top = 1, bottom = 1, left = 1, right = 1)

# Landscape, single column
landscape_one_column <- block_section(
  prop_section(
    page_size = page_size(orient = "landscape"),
    page_margins = std_margins,
    type = "continuous"
  )
)

# Landscape, two columns
landscape_two_columns <- block_section(
  prop_section(
    page_size = page_size(orient = "landscape"),
    page_margins = std_margins,  # Same margins!
    type = "continuous",
    section_columns = section_columns(widths = c(4, 4))
  )
)

doc_section_3 <- read_docx() |>
  body_add_table(value = head(mtcars), style = "table_template") |>
  body_end_block_section(value = landscape_one_column) |>
  body_add_par(value = paste(rep(letters, 60), collapse = " ")) |>
  body_end_block_section(value = landscape_two_columns)

print(doc_section_3, target = file.path(sub_path, "example_complex_section.docx"))

