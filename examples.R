# Load the main script
source("pptx_generator.R")

# Required libraries
library(officer)
library(magrittr)
library(ggplot2) 
# library(mschart)

# ============================================================================
# EXAMPLE 1: Basic Usage
# ============================================================================

example_basic_usage <- function() {
  cat("Running Example 1: Basic Usage (Enhanced)\n")
  cat("=========================================\n")
  
  # Create new presentation
  ppt <- read_pptx()
  
  # Create sample plot (your original example)
  p <- ggplot(mtcars, aes(mpg, wt)) + 
    geom_point(aes(color = factor(cyl)), size = 3) +
    theme_minimal() +
    labs(title = "MPG vs Weight", color = "Cylinders")
  
  # Slide 1: Absolute positioning (your original approach)
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      list(type = "text", value = "Hello, this is a title",
           left = 1, top = 1, width = 5, height = 1),
      list(type = "plot", plot = p,
           left = 1, top = 2, width = 6, height = 4)
    ),
    layout = "Blank"
  )
  
  # Slide 2: Mixed positioning - placeholder + absolute
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      # Use layout's title placeholder
      list(type = "text", value = "Mixed Positioning Example", placeholder = "title"),
      # Custom positioned subtitle
      list(type = "text", value = "Subtitle with absolute positioning",
           left = 1, top = 1.5, width = 8, height = 0.5),
      # Plot in main content area
      list(type = "plot", plot = p, placeholder = "body")
    ),
    layout = "Title and Content"
  )
  
  # Save the file
  print(ppt, target = "example_basic_enhanced.pptx")
  cat("Saved: example_basic_enhanced.pptx\n\n")
  
  return(ppt)
}

# ============================================================================
# EXAMPLE 2: Advanced Plot Integration (ggplot2 + mschart)
# ============================================================================

example_plot_integration <- function() {
  cat("Running Example 2: Advanced Plot Integration\n")
  cat("===========================================\n")
  
  ppt <- read_pptx()
  
  # Create various ggplot2 charts
  # Scatter plot
  scatter_plot <- ggplot(mtcars, aes(mpg, wt)) + 
    geom_point(aes(color = factor(cyl)), size = 3, alpha = 0.7) +
    geom_smooth(method = "lm", se = FALSE, color = "darkblue") +
    theme_minimal() +
    labs(title = "Vehicle Efficiency Analysis",
         subtitle = "Miles per gallon vs weight by cylinder count",
         x = "Miles per Gallon", y = "Weight (1000 lbs)",
         color = "Cylinders")
  
  # Bar chart
  mpg_summary <- aggregate(mpg ~ cyl, data = mtcars, mean)
  bar_plot <- ggplot(mpg_summary, aes(factor(cyl), mpg)) +
    geom_col(fill = "steelblue", alpha = 0.8) +
    geom_text(aes(label = round(mpg, 1)), vjust = -0.5) +
    theme_minimal() +
    labs(title = "Average MPG by Cylinder Count",
         x = "Number of Cylinders", y = "Average MPG")
  
  # Histogram
  hist_plot <- ggplot(mtcars, aes(mpg)) +
    geom_histogram(bins = 10, fill = "lightblue", color = "darkblue", alpha = 0.7) +
    theme_minimal() +
    labs(title = "Distribution of MPG", x = "Miles per Gallon", y = "Count")
  
  # Slide 1: Single large plot with placeholder
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      create_text_placeholder("Statistical Analysis Dashboard", "title"),
      create_plot_placeholder(scatter_plot, "body")
    ),
    layout = "Title and Content"
  )
  
  # Slide 2: Multiple plots with absolute positioning
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      create_text_absolute("Multiple Chart Comparison", 1, 0.2, 8, 0.8),
      create_plot_absolute(bar_plot, 0.5, 1.2, 4.5, 3),
      create_plot_absolute(hist_plot, 5.5, 1.2, 4, 3),
      create_text_absolute("Bar Chart: Average by Group", 0.5, 4.5, 4.5, 0.5),
      create_text_absolute("Histogram: Distribution", 5.5, 4.5, 4, 0.5)
    ),
    layout = "Blank"
  )
  
  # mschart example (if available)
  if (requireNamespace("mschart", quietly = TRUE)) {
    library(mschart)
    
    # Create mschart
    chart_data <- data.frame(
      Month = c("Jan", "Feb", "Mar", "Apr", "May", "Jun"),
      Sales = c(120, 150, 180, 160, 200, 250),
      Target = c(100, 140, 170, 180, 190, 220)
    )
    
    ms_chart <- ms_linechart(data = chart_data, x = "Month", y = c("Sales", "Target")) %>%
      chart_settings(style = "Office") %>%
      chart_labels(title = "Sales Performance vs Target", 
                   xlab = "Month", ylab = "Amount")
    
    # Slide 3: mschart integration
    ppt <- add_custom_slide(
      ppt,
      contents = list(
        create_text_placeholder("Business Performance Chart", "title"),
        create_plot_placeholder(ms_chart, "body")
      ),
      layout = "Title and Content"
    )
    
    cat("Added mschart example\n")
  } else {
    cat("mschart not available - skipping mschart examples\n")
  }
  
  print(ppt, target = "example_plot_integration.pptx")
  cat("Saved: example_plot_integration.pptx\n\n")
  
  return(ppt)
}

# ============================================================================
# EXAMPLE 3: Image Integration and Layouts
# ============================================================================

example_image_integration <- function() {
  cat("Running Example 3: Image Integration (Placeholder)\n")
  cat("=================================================\n")
  
  ppt <- read_pptx()
  
  # Create a sample plot to save as image
  sample_plot <- ggplot(mtcars, aes(factor(cyl), mpg)) +
    geom_boxplot(fill = "lightblue", alpha = 0.7) +
    theme_minimal() +
    labs(title = "MPG Distribution by Cylinders",
         x = "Number of Cylinders", y = "Miles per Gallon")
  
  # Save plot as image file
  ggsave("temp_plot.png", sample_plot, width = 6, height = 4, dpi = 300)
  
  # Slide 1: Image with placeholder positioning
  if (file.exists("temp_plot.png")) {
    ppt <- add_custom_slide(
      ppt,
      contents = list(
        create_text_placeholder("Image Integration Example", "title"),
        create_image_placeholder("temp_plot.png", "body")
      ),
      layout = "Title and Content"
    )
    
    # Slide 2: Multiple images with absolute positioning
    ppt <- add_custom_slide(
      ppt,
      contents = list(
        create_text_absolute("Multiple Image Layout", 1, 0.2, 8, 0.8),
        create_image_absolute("temp_plot.png", 0.5, 1.2, 4, 3),
        create_image_absolute("temp_plot.png", 5, 1.2, 4, 3),
        create_text_absolute("Chart Copy 1", 0.5, 4.5, 4, 0.5),
        create_text_absolute("Chart Copy 2", 5, 4.5, 4, 0.5)
      ),
      layout = "Blank"
    )
    
    # Clean up temporary file
    if (file.exists("temp_plot.png")) {
      file.remove("temp_plot.png")
      cat("Cleaned up temporary image file\n")
    }
  } else {
    cat("Could not create temporary image - skipping image examples\n")
  }
  
  print(ppt, target = "example_image_integration.pptx")
  cat("Saved: example_image_integration.pptx\n\n")
  
  return(ppt)
}

# ============================================================================
# EXAMPLE 4: Multi slide
# ============================================================================

create_demo_presentation <- function() {
  
  # Initialize presentation
  ppt <- read_pptx()
  
  message("Creating demo presentation with multiple content types and positioning strategies...")
  
  # Load required libraries for demo
  if (!requireNamespace("ggplot2", quietly = TRUE)) {
    message("ggplot2 not available - skipping ggplot examples")
    create_ggplot <- FALSE
  } else {
    library(ggplot2)
    create_ggplot <- TRUE
  }
  
  if (!requireNamespace("mschart", quietly = TRUE)) {
    message("mschart not available - skipping mschart examples")
    create_mschart <- FALSE
  } else {
    library(mschart)
    create_mschart <- TRUE
  }
  
  # Slide 1: Title slide with placeholder positioning
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      create_text_placeholder("PowerPoint Automation Demo", "ctrTitle"),
      create_text_placeholder("Enhanced officer Package Integration", "subTitle")
    ),
    layout = "Title Slide"
  )
  
  # Slide 2: Mixed positioning with text and plots
  if (create_ggplot) {
    # Create sample ggplot
    gg_plot <- ggplot(mtcars, aes(mpg, wt, color = factor(cyl))) + 
      geom_point(size = 3) +
      theme_minimal() +
      labs(title = "MPG vs Weight by Cylinders", 
           color = "Cylinders",
           x = "Miles per Gallon",
           y = "Weight (1000 lbs)")
    
    ppt <- add_custom_slide(
      ppt,
      contents = list(
        create_text_placeholder("Data Visualization Examples", "title"),
        create_text_absolute("Custom positioned note", 0.5, 0.5, 3, 0.5),
        create_plot_placeholder(gg_plot, "body")
      ),
      layout = "Title and Content"
    )
  }
  
  # Slide 3: mschart example (if available)
  if (create_mschart) {
    # Create sample mschart
    chart_data <- data.frame(
      category = c("A", "B", "C", "D"),
      values = c(10, 25, 15, 30)
    )
    
    ms_chart <- ms_barchart(data = chart_data, x = "category", y = "values") %>%
      chart_settings(style = "Office")
    
    ppt <- add_custom_slide(
      ppt,
      contents = list(
        create_text_placeholder("Business Chart Example", "title"),
        create_plot_absolute(ms_chart, 1, 2, 8, 4)
      ),
      layout = "Title and Content"
    )
  }
  
  # Slide 4: Absolute positioning demonstration
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
  
  message("Demo presentation created successfully!")
  return(ppt)
}


# ============================================================================
# EXAMPLE 5: Layout Exploration and Template Usage
# ============================================================================

example_layout_exploration <- function() {
  cat("Running Example 5: Layout Exploration\n")
  cat("====================================\n")
  
  ppt <- read_pptx()
  
  # Demonstrate different layouts
  sample_plot <- ggplot(mtcars, aes(mpg, wt)) + 
    geom_point(size = 2) + 
    theme_minimal()
  
  # Title Slide layout
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      create_text_placeholder("PowerPoint Automation", "ctrTitle"),
      create_text_placeholder("Layout Demonstration", "subTitle")
    ),
    layout = "Title Slide"
  )
  
  # Title and Content layout
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      create_text_placeholder("Main Content Slide", "title"),
      create_plot_placeholder(sample_plot, "body")
    ),
    layout = "Title and Content"
  )
  
  # Two Content layout (if available)
  tryCatch({
    ppt <- add_custom_slide(
      ppt,
      contents = list(
        create_text_placeholder("Two Content Layout", "title"),
        create_text_absolute("Left content area", 0.5, 2, 4, 3),
        create_text_absolute("Right content area", 5.5, 2, 4, 3)
      ),
      layout = "Two Content"
    )
    cat("Added Two Content layout slide\n")
  }, error = function(e) {
    cat("Two Content layout not available in template\n")
  })
  
  # Blank layout with full absolute positioning
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      create_text_absolute("Blank Layout - Full Control", 1, 0.5, 8, 1),
      create_plot_absolute(sample_plot, 1, 1.5, 4, 3),
      create_text_absolute("Custom positioned elements", 6, 2, 3, 1),
      create_text_absolute("Complete flexibility", 6, 3, 3, 1)
    ),
    layout = "Blank"
  )
  
  print(ppt, target = "example_layout_exploration.pptx")
  cat("Saved: example_layout_exploration.pptx\n\n")
  
  return(ppt)
}

# ============================================================================
# EXAMPLE 6: Helper Functions Demonstration
# ============================================================================

example_helper_functions <- function() {
  cat("Running Example 6: Helper Functions\n")
  cat("===================================\n")
  
  ppt <- read_pptx()
  
  # Using helper functions for cleaner code
  sample_plot <- ggplot(iris, aes(Sepal.Length, Sepal.Width, color = Species)) +
    geom_point(size = 2) +
    theme_minimal() +
    labs(title = "Iris Dataset Analysis")
  
  contents <- list(
    create_text_placeholder("Helper Functions Demo", "title"),
    create_text_absolute("Created with helper functions", 1, 1.5, 6, 0.5),
    create_plot_placeholder(sample_plot, "body")
  )
  
  validate_contents(contents)
  
  ppt <- add_custom_slide(ppt, contents = contents, layout = "Title and Content")
  
  print(ppt, target = "example_helper_functions.pptx")
  cat("Saved: example_helper_functions.pptx\n\n")
  
  return(ppt)
}



example_basic_usage()
example_plot_integration()
example_image_integration()
example_layout_exploration()
example_helper_functions()
create_demo_presentation()