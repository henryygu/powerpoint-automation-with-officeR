library(officer)
library(magrittr)

add_custom_slide <- function(ppt, 
                             contents = list(),
                             layout = "Blank", 
                             master = "Office Theme") {
  
  if (!inherits(ppt, "rpptx")) {
    stop("ppt must be a PowerPoint object created with read_pptx()")
  }
  
  if (!is.list(contents)) {
    stop("contents must be a list of content items")
  }
  
  ppt <- add_slide(ppt, layout = layout, master = master)
  
  for (i in seq_along(contents)) {
    item <- contents[[i]]
    
    tryCatch({
      ppt <- process_content_item(ppt, item, i)
    }, error = function(e) {
      warning(sprintf("Failed to process content item %d: %s", i, e$message))
    })
  }
  
  return(ppt)
}

process_content_item <- function(ppt, item, index = 1) {
  
  if (is.null(item$type) || !item$type %in% c("text", "image", "plot")) {
    stop(sprintf("Item %d: Invalid or missing 'type'. Must be 'text', 'image', or 'plot'", index))
  }
  
  location <- determine_location(item, index)
  
  switch(item$type,
    "text" = process_text_content(ppt, item, location, index),
    "image" = process_image_content(ppt, item, location, index),
    "plot" = process_plot_content(ppt, item, location, index),
    stop(sprintf("Item %d: Unsupported content type: %s", index, item$type))
  )
}

determine_location <- function(item, index) {
  
  has_absolute <- all(c("left", "top", "width", "height") %in% names(item))
  
  has_placeholder <- "placeholder" %in% names(item)
  
  if (has_absolute) {
    coords <- c(item$left, item$top, item$width, item$height)
    if (any(is.na(coords)) || any(coords < 0)) {
      stop(sprintf("Item %d: Invalid coordinates. All values must be positive numbers", index))
    }
    
    return(ph_location(
      left = item$left, 
      top = item$top,
      width = item$width, 
      height = item$height
    ))
    
  } else if (has_placeholder) {
    if (is.null(item$placeholder) || item$placeholder == "") {
      stop(sprintf("Item %d: Empty placeholder name", index))
    }
    
    return(ph_location_type(type = item$placeholder))
    
  } else {
    stop(sprintf("Item %d: No positioning specified. Provide either coordinates (left, top, width, height) or placeholder", index))
  }
}

process_text_content <- function(ppt, item, location, index) {
  
  if (is.null(item$value) || item$value == "") {
    stop(sprintf("Item %d: Missing or empty text value", index))
  }
  
  ppt %>% 
    ph_with(
      value = item$value,
      location = location
    )
}

process_image_content <- function(ppt, item, location, index) {
  
  if (is.null(item$path) || item$path == "") {
    stop(sprintf("Item %d: Missing image path", index))
  }
  
  if (!file.exists(item$path)) {
    stop(sprintf("Item %d: Image file not found: %s", index, item$path))
  }
  
  if (inherits(location, "location_str")) {
    img_obj <- external_img(item$path)
  } else {
    img_obj <- external_img(item$path, width = location$width, height = location$height)
  }
  
  ppt %>% 
    ph_with(
      value = img_obj,
      location = location
    )
}

process_plot_content <- function(ppt, item, location, index) {
  
  if (is.null(item$plot)) {
    stop(sprintf("Item %d: Missing plot object", index))
  }
  
  plot_obj <- item$plot
  is_ggplot <- inherits(plot_obj, "ggplot")
  is_mschart <- inherits(plot_obj, c("ms_chart", "ms_barchart", "ms_linechart", "ms_scatterchart", "ms_areachart"))
  
  if (!is_ggplot && !is_mschart) {
    stop(sprintf("Item %d: Plot must be a ggplot2 or mschart object. Found: %s", 
                index, paste(class(plot_obj), collapse = ", ")))
  }
  
  ppt %>% 
    ph_with(
      value = plot_obj,
      location = location
    )
}

validate_contents <- function(contents) {
  
  if (length(contents) == 0) {
    message("Warning: Empty contents list - slide will be blank")
    return(TRUE)
  }
  
  for (i in seq_along(contents)) {
    item <- contents[[i]]
    
    if (!is.list(item)) {
      stop(sprintf("Item %d: Must be a list", i))
    }
    
    if (is.null(item$type)) {
      stop(sprintf("Item %d: Missing 'type' field", i))
    }
  }
  
  return(TRUE)
}

create_text_absolute <- function(value, left, top, width, height) {
  list(
    type = "text",
    value = value,
    left = left,
    top = top,
    width = width,
    height = height
  )
}

create_text_placeholder <- function(value, placeholder) {
  list(
    type = "text",
    value = value,
    placeholder = placeholder
  )
}

create_image_absolute <- function(path, left, top, width, height) {
  list(
    type = "image",
    path = path,
    left = left,
    top = top,
    width = width,
    height = height
  )
}

create_image_placeholder <- function(path, placeholder) {
  list(
    type = "image",
    path = path,
    placeholder = placeholder
  )
}

create_plot_absolute <- function(plot, left, top, width, height) {
  list(
    type = "plot",
    plot = plot,
    left = left,
    top = top,
    width = width,
    height = height
  )
}

create_plot_placeholder <- function(plot, placeholder) {
  list(
    type = "plot",
    plot = plot,
    placeholder = placeholder
  )
}

get_available_layouts <- function(ppt) {
  if (is.character(ppt)) {
    ppt <- read_pptx(ppt)
  }
  
  layout_summary(ppt)$name
}

get_layout_placeholders <- function(ppt, layout_name) {
  if (is.character(ppt)) {
    ppt <- read_pptx(ppt)
  }
  
  temp_ppt <- add_slide(ppt, layout = layout_name)
  
  slide_summary(temp_ppt, index = length(temp_ppt))
}

create_demo_presentation <- function() {
  
  ppt <- read_pptx()
  
  message("Creating demo presentation with multiple content types and positioning strategies...")
  
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
  
  ppt <- add_custom_slide(
    ppt,
    contents = list(
      create_text_placeholder("PowerPoint Automation Demo", "ctrTitle"),
      create_text_placeholder("Enhanced officer Package Integration", "subTitle")
    ),
    layout = "Title Slide"
  )
  
  if (create_ggplot) {
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
  
  if (create_mschart) {
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

