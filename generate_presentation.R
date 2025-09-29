# Load required libraries
library(officer)
library(dplyr)
library(ggplot2)
library(readr)

# Function to create PowerPoint presentation from metadata and data files
create_presentation <- function(slide_metadata_file, content_metadata_file, data_file, output_file) {
  # Read metadata and data files
  slide_metadata <- read.csv(slide_metadata_file, stringsAsFactors = FALSE)
  content_metadata <- read.csv(content_metadata_file, stringsAsFactors = FALSE)
  data <- read.csv(data_file, stringsAsFactors = FALSE)
  
  # Create sample graph objects for demonstration
  # In practice, these would be defined in a separate script
  graph_objects <- create_sample_graphs(data)
  
  # Create a new PowerPoint presentation
  ppt <- read_pptx()
  
  # Get unique slide numbers
  slide_numbers <- unique(slide_metadata$slide_number)
  
  # Process each slide
  for (slide_num in slide_numbers) {
    # Get slide layout and master
    slide_info <- slide_metadata[slide_metadata$slide_number == slide_num, ]
    layout <- slide_info$layout[1]
    master <- slide_info$master[1]
    
    # Add a new slide with specified layout and master
    ppt <- add_slide(ppt, layout = layout, master = master)
    
    # Get content for this slide
    slide_content <- content_metadata[content_metadata$slide_number == slide_num, ]
    
    # Process each content item on the slide
    for (i in 1:nrow(slide_content)) {
      content_item <- slide_content[i, ]
      
      # Handle different content types
      if (content_item$content_type == "text") {
        # Add text content
        if (content_item$position_type == "object_name") {
          # Position by object name
          ppt <- ph_with(ppt, 
                         value = content_item$metric,
                         location = ph_location_label(content_item$object_name))
        } else if (content_item$position_type == "coordinates") {
          # Position by coordinates
          ppt <- ph_with(ppt,
                         value = content_item$metric,
                         location = ph_location(
                           left = content_item$left,
                           top = content_item$top,
                           width = content_item$width,
                           height = content_item$height,
                           bg = content_item$bg
                         ))
        }
      } else if (content_item$content_type == "graph") {
        # Create and add graph
        if (content_item$position_type == "object_name") {
          # Position by object name
          # Check if there's a predefined graph object
          if (exists(content_item$metric, where = graph_objects)) {
            p <- get(content_item$metric, pos = graph_objects)
          } else {
            # For demo purposes, we'll create a simple plot
            p <- create_sample_plot(data, content_item$metric)
          }
          ppt <- ph_with(ppt, 
                         value = p,
                         location = ph_location_label(content_item$object_name))
        } else if (content_item$position_type == "coordinates") {
          # Position by coordinates
          # Check if there's a predefined graph object
          if (exists(content_item$metric, where = graph_objects)) {
            p <- get(content_item$metric, pos = graph_objects)
          } else {
            # For demo purposes, we'll create a simple plot
            p <- create_sample_plot(data, content_item$metric)
          }
          ppt <- ph_with(ppt,
                         value = p,
                         location = ph_location(
                           left = content_item$left,
                           top = content_item$top,
                           width = content_item$width,
                           height = content_item$height,
                           bg = content_item$bg
                         ))
        }
      } else if (content_item$content_type == "table") {
        # Add table content
        if (content_item$position_type == "object_name") {
          # Position by object name
          table_data <- prepare_table_data(data, content_item$metric)
          ppt <- ph_with(ppt, 
                         value = table_data,
                         location = ph_location_label(content_item$object_name))
        } else if (content_item$position_type == "coordinates") {
          # Position by coordinates
          table_data <- prepare_table_data(data, content_item$metric)
          ppt <- ph_with(ppt,
                         value = table_data,
                         location = ph_location(
                           left = content_item$left,
                           top = content_item$top,
                           width = content_item$width,
                           height = content_item$height,
                           bg = content_item$bg
                         ))
        }
      }
    }
  }
  
  # Save the presentation
  print(ppt, target = output_file)
  cat("Presentation saved to", output_file, "\n")
}

# Helper function to create sample graph objects
create_sample_graphs <- function(data) {
  # Create a new environment to store graph objects
  graph_env <- new.env()
  
  # Create sample graphs as shown in example.R
  
  # Graph 1: Revenue trend graph
  graph_env$graph1 <- data %>%
    filter(Metric == "Revenue") %>%
    ggplot(aes(x = Date, y = Value, color = Organization, group = Organization)) +
    geom_line(size = 1.2) +
    geom_point(size = 3) +
    theme_minimal() +
    labs(
      title = "Revenue Trends by Organization",
      x = "Date",
      y = "Revenue (USD)",
      color = "Organization"
    ) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
  
  # Graph 2: Profit margin comparison chart
  graph_env$graph2 <- data %>%
    filter(Metric == "Profit Margin") %>%
    ggplot(aes(x = Organization, y = Value, fill = Organization)) +
    geom_bar(stat = "identity") +
    theme_minimal() +
    labs(
      title = "Profit Margin Comparison",
      x = "Organization",
      y = "Profit Margin",
      fill = "Organization"
    ) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
  
  # Graph 3: Customer satisfaction heatmap
  graph_env$graph3 <- data %>%
    filter(Metric == "Customer Satisfaction") %>%
    ggplot(aes(x = Date, y = Organization, fill = Value)) +
    geom_tile() +
    theme_minimal() +
    labs(
      title = "Customer Satisfaction Heatmap",
      x = "Date",
      y = "Organization",
      fill = "Satisfaction Score"
    ) +
    scale_fill_gradient(low = "red", high = "green")
  
  # Graph 4: Market share pie chart
  latest_date <- max(data$Date)
  graph_env$graph4 <- data %>%
    filter(Metric == "Market Share" & Date == latest_date) %>%
    ggplot(aes(x = "", y = Value, fill = Organization)) +
    geom_bar(stat = "identity", width = 1) +
    coord_polar("y", start = 0) +
    theme_minimal() +
    labs(
      title = "Market Share Distribution",
      fill = "Organization"
    ) +
    theme(axis.title.x = element_blank(),
          axis.title.y = element_blank(),
          axis.text.x = element_blank(),
          axis.ticks = element_blank())
  
  # Return the environment containing all graph objects
  return(graph_env)
}

# Helper function to create sample plots
create_sample_plot <- function(data, metric_name) {
  # Filter data for the specific metric
  filtered_data <- data[data$Metric == metric_name | grepl(metric_name, data$Metric), ]
  
  # If no specific metric found, use the first few rows as sample
  if (nrow(filtered_data) == 0) {
    filtered_data <- head(data, 10)
  }
  
  # Create a simple bar plot
  p <- ggplot(filtered_data, aes(x = Organization, y = Value, fill = Organization)) +
    geom_bar(stat = "identity") +
    theme_minimal() +
    labs(title = paste("Sample Plot for", metric_name),
         x = "Organization",
         y = "Value") +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
  
  return(p)
}

# Helper function to prepare table data
prepare_table_data <- function(data, metric_name) {
  # Filter data for the specific metric
  filtered_data <- data[data$Metric == metric_name | grepl(metric_name, data$Metric), ]
  
  # If no specific metric found, use the first few rows as sample
  if (nrow(filtered_data) == 0) {
    filtered_data <- head(data, 10)
  }
  
  # Select relevant columns for display
  table_data <- filtered_data[, c("Organization", "Value", "Date")]
  
  return(table_data)
}

# Main execution
main <- function() {
  # Define file paths
  slide_metadata_file <- "slide_metadata.csv"
  content_metadata_file <- "content_metadata.csv"
  data_file <- "data.csv"
  output_file <- "generated_presentation.pptx"
  
  # Check if required files exist
  if (!file.exists(slide_metadata_file)) {
    stop(paste("Slide metadata file not found:", slide_metadata_file))
  }
  
  if (!file.exists(content_metadata_file)) {
    stop(paste("Content metadata file not found:", content_metadata_file))
  }
  
  if (!file.exists(data_file)) {
    stop(paste("Data file not found:", data_file))
  }
  
  # Generate the presentation
  cat("Generating PowerPoint presentation...\n")
  create_presentation(slide_metadata_file, content_metadata_file, data_file, output_file)
  cat("Done!\n")
}

# Run the main function
main()