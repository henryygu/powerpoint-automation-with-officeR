# PowerPoint Automation with officer Package

Simple R tool for creating PowerPoint presentations with text, plots, and images.

## Features

- **Two positioning modes**: Use placeholders or precise coordinates
- **Multiple plot types**: ggplot2 and mschart support
- **Content types**: Text, images, and plots

## Installation

```r
install.packages(c("officer", "magrittr", "ggplot2"))
```

## Basic Usage

```r
source("pptx_generator.R")
library(officer)

ppt <- read_pptx()
p <- ggplot(mtcars, aes(mpg, wt)) + geom_point()

ppt <- add_custom_slide(
  ppt,
  contents = list(
    list(type = "text", value = "My Title", placeholder = "title"),
    list(type = "plot", plot = p, placeholder = "body")
  ),
  layout = "Title and Content"
)

print(ppt, target = "presentation.pptx")
```

## Positioning Options

**Placeholder positioning** (recommended):
```r
list(type = "text", value = "Title", placeholder = "title")
```

**Absolute positioning** (precise control):
```r
list(type = "text", value = "Text", left = 1, top = 2, width = 5, height = 1)
```

**Common placeholders**: `"title"`, `"body"`, `"ctrTitle"`, `"subTitle"`

## Content Types

**Text**: `list(type = "text", value = "Hello")`
**Images**: `list(type = "image", path = "chart.png")`
**Plots**: `list(type = "plot", plot = my_ggplot)`

## Complete Example

Here's how to create a full presentation with multiple slide types:

```r
create_demo_presentation <- function() {
  ppt <- read_pptx()
  
  # Load libraries
  library(ggplot2)
  
  # Slide 1: Title slide
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      create_text_placeholder("PowerPoint Automation Demo", "ctrTitle"),
      create_text_placeholder("Enhanced officer Package Integration", "subTitle")
    ),
    layout = "Title Slide"
  )
  
  # Slide 2: Data visualization
  gg_plot <- ggplot(mtcars, aes(mpg, wt, color = factor(cyl))) + 
    geom_point(size = 3) +
    theme_minimal() +
    labs(title = "MPG vs Weight by Cylinders")
  
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      create_text_placeholder("Data Visualization Examples", "title"),
      create_text_absolute("Custom positioned note", 0.5, 0.5, 3, 0.5),
      create_plot_placeholder(gg_plot, "body")
    ),
    layout = "Title and Content"
  )
  
  # Slide 3: Absolute positioning demo
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      create_text_absolute("Absolute Positioning Demo", 1, 0.5, 8, 1),
      create_text_absolute("Top Left", 0.5, 1.5, 2, 0.5),
      create_text_absolute("Top Right", 7.5, 1.5, 2, 0.5),
      create_text_absolute("Bottom Center", 4, 6, 2, 0.5)
    ),
    layout = "Blank"
  )
  
  return(ppt)
}

# Create and save the presentation
ppt <- create_demo_presentation()
print(ppt, target = "demo.pptx")
```

## More Examples

```r
source("examples.R")  # Run all example functions
```

## Main Function

`add_custom_slide(ppt, contents, layout = "Title and Content")`

- `ppt`: PowerPoint object
- `contents`: List of content items
- `layout`: Slide layout name

## Common Layouts

- `"Title Slide"` - For presentation titles
- `"Title and Content"` - Standard slide
- `"Blank"` - Custom positioning only