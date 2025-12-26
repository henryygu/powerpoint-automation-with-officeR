# Load required libraries
library(officer)
library(dplyr)
library(ggplot2)
library(readr)
library(mschart)
library(stringr)

# Function to create mscharts
create_mschart <- function(data_subset, title, geom_type, x_var, y_var, fill_var) {
  
  if (nrow(data_subset) == 0) return(NULL)
  
  # Initialize chart
  if (grepl("bar", geom_type)) {
    # Bar chart
    # For mschart, we typically want data in a format where categorical x and numeric y are clear
    # We might need to handle grouping
    if (!is.na(fill_var) && fill_var != "") {
      # Grouped bar chart
      p <- ms_barchart(data_subset, x = x_var, y = y_var, group = fill_var)
    } else {
      p <- ms_barchart(data_subset, x = x_var, y = y_var)
    }
  } else if (grepl("line", geom_type)) {
    # Line chart
    if (!is.na(fill_var) && fill_var != "") {
      p <- ms_linechart(data_subset, x = x_var, y = y_var, group = fill_var)
    } else {
      p <- ms_linechart(data_subset, x = x_var, y = y_var)
    }
  } else if (grepl("scatter", geom_type) || grepl("point", geom_type)) {
    # Scatter plot
    if (!is.na(fill_var) && fill_var != "") {
      p <- ms_scatterchart(data_subset, x = x_var, y = y_var, group = fill_var)
    } else {
      p <- ms_scatterchart(data_subset, x = x_var, y = y_var)
    }
  } else {
    # Default to bar
    p <- ms_barchart(data_subset, x = x_var, y = y_var)
  }
  
  # Basic formatting
  p <- chart_settings(p, grouped = (!is.na(fill_var) && fill_var != ""))
  p <- chart_labs(p, title = title, xlab = x_var, ylab = y_var)
  
  return(p)
}

# Function to prepare data for tables
prepare_table_data <- function(data_subset) {
  # Reshape data for better table presentation
  # Simplistic approach: just show Org and Value if those exist, or whole subset
  if ("Organization" %in% names(data_subset) && "Value" %in% names(data_subset)) {
      table_data <- data_subset %>%
        select(Organization, Value) %>%
        mutate(Value = round(Value, 2))
  } else {
      table_data <- data_subset
  }
  return(table_data)
}

# Main function to generate presentation
generate_presentation <- function(content_metadata, data, external_objects = list(), template_path = NULL) {
  
  # Load template or default
  if (!is.null(template_path) && file.exists(template_path)) {
    presentation <- read_pptx(template_path)
  } else {
    presentation <- read_pptx()
  }
  
  # Get unique slide numbers and iterate
  # We assume the CSV is sorted or we handle it by slide_number
  slide_numbers <- unique(content_metadata$slide_number)
  
  for (slide_num in slide_numbers) {
    # Get all content for this slide
    slide_content_rows <- content_metadata[content_metadata$slide_number == slide_num, ]
    
    # Get slide definition from the first row of this slide
    # Fallback to defaults if missing
    s_master <- if(!is.na(slide_content_rows$slide_master[1]) && slide_content_rows$slide_master[1] != "") slide_content_rows$slide_master[1] else "Office Theme"
    s_layout <- if(!is.na(slide_content_rows$slide_layout[1]) && slide_content_rows$slide_layout[1] != "") slide_content_rows$slide_layout[1] else "Title and Content"
    
    # Add slide
    # Check if layout exists, if not, try to error gracefully or fallback (officer will error if invalid)
    tryCatch({
      presentation <- add_slide(presentation, layout = s_layout, master = s_master)
    }, error = function(e) {
      warning(paste("Could not add slide", slide_num, "with layout", s_layout, "and master", s_master, "- using defaults."))
      presentation <- add_slide(presentation) 
    })
    
    # Iterate through content items
    for (i in 1:nrow(slide_content_rows)) {
      item <- slide_content_rows[i, ]
      
      # Determine content object to place
      content_obj <- NULL
      
      if (item$content_type == "text") {
        # Handle text and bullets
        raw_text <- as.character(item$metric)
        
        # Check for delimiters for bullets (e.g. pipe |)
        if (grepl("\\|", raw_text)) {
          # Split into chunks
          chunks <- str_split(raw_text, "\\|")[[1]]
          chunks <- chunks[chunks != ""] # Remove empty
          
          # Create block list of fpars
          fpar_list <- lapply(chunks, function(txt) {
            # Style text
            fp_t <- fp_text(font.size = if(!is.na(item$font_size)) as.numeric(item$font_size) else 18,
                            color = if(!is.na(item$font_color)) item$font_color else "black")
            fpar(ftext(txt, fp_t))
          })
          content_obj <- do.call(block_list, fpar_list)
          
        } else {
          # Single paragraph
           fp_t <- fp_text(font.size = if(!is.na(item$font_size)) as.numeric(item$font_size) else 18,
                            color = if(!is.na(item$font_color)) item$font_color else "black")
           content_obj <- fpar(ftext(raw_text, fp_t))
        }
        
      } else if (item$content_type == "table") {
        table_data <- data[data$Metric == item$metric, ]
        if (nrow(table_data) > 0) {
          content_obj <- prepare_table_data(table_data)
        }
        
      } else if (grepl("mschart", item$content_type)) {
        chart_data <- data[data$Metric == item$metric, ]
        if (nrow(chart_data) > 0) {
          content_obj <- create_mschart(chart_data, item$metric, item$geom_type, item$x_var, item$y_var, item$fill_var)
          
          # Apply background processing if needed (mschart themes)
          # if (!is.na(item$bg)) ... 
        }
        
      } else if (item$content_type == "custom") {
        # Look up in external_objects
        obj_name <- item$metric
        if (!is.null(external_objects[[obj_name]])) {
          content_obj <- external_objects[[obj_name]]
        }
      }
      
      # Determine Location and Place Logic
      if (!is.null(content_obj)) {
        # Check if coordinates are provided
        if (item$position_type == "coordinates" && 
            !is.na(item$left) && !is.na(item$top) && 
            !is.na(item$width) && !is.na(item$height)) {
          
          loc <- ph_location(left = as.numeric(item$left), 
                             top = as.numeric(item$top), 
                             width = as.numeric(item$width), 
                             height = as.numeric(item$height),
                             rotation = if(!is.na(item$rotation)) as.numeric(item$rotation) else 0)
          
          presentation <- ph_with(presentation, value = content_obj, location = loc)
          
        } else if (item$position_type == "placeholder_name" && !is.na(item$placeholder_name)) {
          # Named placeholder
          presentation <- ph_with(presentation, value = content_obj, location = ph_location_label(item$placeholder_name))
          
        } else if (item$position_type == "placeholder_type" && !is.na(item$placeholder_name)) {
             # Type-based placeholder (using placeholder_name col as the type key)
             presentation <- ph_with(presentation, value = content_obj, location = ph_location_type(type = item$placeholder_name))
             
        } else {
             # Default body
             presentation <- ph_with(presentation, value = content_obj, location = ph_location_type(type = "body"))
        }
      }
    }
  }
  
  return(presentation)
}