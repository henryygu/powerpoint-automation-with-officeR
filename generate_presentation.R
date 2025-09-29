# Load required libraries
library(officer)
library(dplyr)
library(ggplot2)
library(readr)

# Function to create graphs from data
create_graph <- function(data_subset, title, geom_type, x_var, y_var, fill_var) {
  # Create a ggplot based on the specified geom type and variables
  if (geom_type == "geom_line") {
    # For line charts, we need to ensure proper grouping
    p <- ggplot(data_subset, aes_string(x = x_var, y = y_var, group = fill_var))
  } else {
    # For other chart types
    p <- ggplot(data_subset, aes_string(x = x_var, y = y_var))
  }
  
  # Add the appropriate geom based on geom_type
  if (geom_type == "geom_bar") {
    if (!is.na(fill_var) && fill_var != "") {
      p <- p + geom_bar(aes_string(fill = fill_var), stat = "identity")
    } else {
      p <- p + geom_bar(stat = "identity")
    }
  } else if (geom_type == "geom_line") {
    if (!is.na(fill_var) && fill_var != "") {
      p <- p + geom_line(aes_string(color = fill_var))
    } else {
      p <- p + geom_line()
    }
  } else if (geom_type == "geom_point") {
    if (!is.na(fill_var) && fill_var != "") {
      p <- p + geom_point(aes_string(color = fill_var))
    } else {
      p <- p + geom_point()
    }
  } else {
    # Default to bar chart if geom_type is not specified or unrecognized
    if (!is.na(fill_var) && fill_var != "") {
      p <- p + geom_bar(aes_string(fill = fill_var), stat = "identity")
    } else {
      p <- p + geom_bar(stat = "identity")
    }
  }
  
  # Add theme and labels
  p <- p + theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.title = element_text(hjust = 0.5)) +
    labs(title = title)
  
  return(p)
}

# Function to prepare data for tables
prepare_table_data <- function(data_subset) {
  # Reshape data for better table presentation
  table_data <- data_subset %>%
    select(Organization, Value) %>%
    mutate(Value = round(Value, 2))
  return(table_data)
}

# Main function to generate presentation
generate_presentation <- function(slide_metadata, content_metadata, data) {
  # Create a new PowerPoint presentation
  presentation <- read_pptx()
  
  # Get unique slide numbers
  slide_numbers <- unique(slide_metadata$slide_number)
  
  # Process each slide
  for (slide_num in slide_numbers) {
    # Get slide layout and master
    slide_info <- slide_metadata[slide_metadata$slide_number == slide_num, ]
    
    # Add a new slide with the specified layout and master
    presentation <- add_slide(presentation, 
                             layout = slide_info$layout, 
                             master = slide_info$master)
    
    # Get content for this slide
    slide_content <- content_metadata[content_metadata$slide_number == slide_num, ]
    
    # Process each content item on the slide
    for (i in 1:nrow(slide_content)) {
      content_item <- slide_content[i, ]
      
      if (content_item$content_type == "text") {
        # Add text content using the specified placeholder
        text_value <- as.character(content_item$metric)
        
        # Handle different position types
        if (content_item$position_type == "placeholder_type") {
          # Use ph_location_type for placeholder types like "ctrTitle", "subTitle", "title"
          presentation <- ph_with(presentation, 
                                 value = text_value,
                                 location = ph_location_type(type = content_item$placeholder_name))
        } else if (content_item$position_type == "placeholder_name") {
          # Use ph_location_label for named placeholders like "Content Placeholder 1"
          presentation <- ph_with(presentation, 
                                 value = text_value,
                                 location = ph_location_label(content_item$placeholder_name))
        } else {
          # Default to "title" placeholder if position_type is not specified or unrecognized
          presentation <- ph_with(presentation, 
                                 value = text_value,
                                 location = ph_location_type(type = "title"))
        }
      } else if (content_item$content_type == "graph") {
        # Filter data for this graph
        graph_data <- data[data$Metric == content_item$metric, ]
        
        # Check if we have data to plot
        if (nrow(graph_data) > 0) {
          # Create the graph with specified parameters
          graph <- create_graph(
            data_subset = graph_data, 
            title = content_item$metric,
            geom_type = content_item$geom_type,
            x_var = content_item$x_var,
            y_var = content_item$y_var,
            fill_var = content_item$fill_var
          )
          
          # Add graph to slide
          if (content_item$position_type == "placeholder_name" && 
              !is.na(content_item$placeholder_name) && 
              content_item$placeholder_name != "") {
            # Use the specified placeholder name
            presentation <- ph_with(presentation,
                                   value = graph,
                                   location = ph_location_label(content_item$placeholder_name))
          } else if (content_item$position_type == "coordinates" && 
                     !is.na(content_item$left) && !is.na(content_item$top)) {
            # Use coordinate-based positioning
            presentation <- ph_with(presentation,
                                   value = graph,
                                   location = ph_location(
                                     left = as.numeric(content_item$left),
                                     top = as.numeric(content_item$top),
                                     width = as.numeric(content_item$width),
                                     height = as.numeric(content_item$height)
                                   ))
          } else {
            # Default positioning if no specific positioning is provided
            presentation <- ph_with(presentation,
                                   value = graph,
                                   location = ph_location_type(type = "body"))
          }
        }
      } else if (content_item$content_type == "table") {
        # Filter data for this table
        table_data <- data[data$Metric == content_item$metric, ]
        
        # Check if we have data for the table
        if (nrow(table_data) > 0) {
          # Prepare table data
          table_df <- prepare_table_data(table_data)
          
          # Add table to slide
          if (content_item$position_type == "placeholder_name" && 
              !is.na(content_item$placeholder_name) && 
              content_item$placeholder_name != "") {
            # Use the specified placeholder name
            presentation <- ph_with(presentation,
                                   value = table_df,
                                   location = ph_location_label(content_item$placeholder_name))
          } else if (content_item$position_type == "coordinates" && 
                     !is.na(content_item$left) && !is.na(content_item$top)) {
            # Use coordinate-based positioning
            presentation <- ph_with(presentation,
                                   value = table_df,
                                   location = ph_location(
                                     left = as.numeric(content_item$left),
                                     top = as.numeric(content_item$top),
                                     width = as.numeric(content_item$width),
                                     height = as.numeric(content_item$height)
                                   ))
          } else {
            # Default positioning if no specific positioning is provided
            presentation <- ph_with(presentation,
                                   value = table_df,
                                   location = ph_location_type(type = "body"))
          }
        }
      }
    }
  }
  
  return(presentation)
}