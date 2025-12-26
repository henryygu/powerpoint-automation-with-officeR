# Script to validate content_metadata.csv
# Run this script to check if your metadata file is correctly formatted before generating the presentation.

library(readr)
library(dplyr)
library(stringr)

validate_metadata <- function(file_path = "content_metadata.csv") {
  
  cat(sprintf("Validating %s...\n", file_path))
  
  if (!file.exists(file_path)) {
    stop("Error: File not found at ", file_path)
  }
  
  # Read with all cols as character initially to safely check types, or stick to auto
  data <- read_csv(file_path, show_col_types = FALSE)
  
  # -------------------------------------------------------------------------
  # 1. Check Required Columns
  # -------------------------------------------------------------------------
  required_cols <- c("slide_number", "slide_master", "slide_layout", 
                     "position_type", "content_type", "placeholder_name", 
                     "metric", "left", "top", "width", "height", 
                     "rotation", "bg", "font_size", "font_color")
  
  missing_cols <- setdiff(required_cols, names(data))
  if (length(missing_cols) > 0) {
    stop("Error: Missing required columns: ", paste(missing_cols, collapse = ", "))
  }
  
  errors <- c()
  warnings <- c()
  
  # -------------------------------------------------------------------------
  # 2. Check Data Types & Values
  # -------------------------------------------------------------------------
  
  # Slide Number
  if (!is.numeric(data$slide_number)) {
    errors <- c(errors, "Column 'slide_number' must be numeric.")
  }
  
  # Position Type
  valid_pos_types <- c("placeholder_type", "placeholder_name", "coordinates")
  # Filter out NAs or empty strings if allowed, assuming strict checking for now
  invalid_pos <- data %>% 
    filter(!is.na(position_type) & position_type != "") %>%
    filter(!position_type %in% valid_pos_types) %>% 
    pull(position_type) %>% unique()
    
  if (length(invalid_pos) > 0) {
    errors <- c(errors, paste("Invalid 'position_type' values found:", paste(invalid_pos, collapse = ", ")))
  }
  
  # Content Type
  valid_content_types <- c("text", "table", "custom")
  # We also allow mschart_*
  
  invalid_content <- data %>% 
    filter(!is.na(content_type) & content_type != "") %>%
    filter(!grepl("mschart", content_type) & !content_type %in% valid_content_types) %>%
    pull(content_type) %>% unique()
    
  if (length(invalid_content) > 0) {
    errors <- c(errors, paste("Invalid 'content_type' values found (expected text, table, custom, mschart_*):", paste(invalid_content, collapse = ", ")))
  }
  
  # -------------------------------------------------------------------------
  # 3. Logic Checks
  # -------------------------------------------------------------------------
  
  # Coordinates check
  # If position_type is 'coordinates', then left/top/width/height must be numeric and not NA
  coord_rows <- which(data$position_type == "coordinates")
  if (length(coord_rows) > 0) {
    coord_data <- data[coord_rows, ]
    if (any(is.na(coord_data$left) | is.na(coord_data$top) | is.na(coord_data$width) | is.na(coord_data$height))) {
       errors <- c(errors, paste("Rows with position_type='coordinates' must have valid numeric 'left', 'top', 'width', and 'height'. Check rows:", paste(coord_rows, collapse = ", ")))
    }
  }
  
  # Placeholder check
  # If position_type is 'placeholder_name' or 'placeholder_type', placeholder_name must be present
  ph_rows <- which(data$position_type %in% c("placeholder_name", "placeholder_type"))
  if (length(ph_rows) > 0) {
    if (any(is.na(data$placeholder_name[ph_rows]) | data$placeholder_name[ph_rows] == "")) {
      errors <- c(errors, paste("Rows with position_type='placeholder_*' must have 'placeholder_name' defined. Check rows:", paste(ph_rows[is.na(data$placeholder_name[ph_rows]) | data$placeholder_name[ph_rows] == ""], collapse = ", ")))
    }
  }

  # Geom type check for mscharts
  chart_rows <- which(grepl("mschart", data$content_type))
  if (length(chart_rows) > 0) {
    if (any(is.na(data$geom_type[chart_rows]) | data$geom_type[chart_rows] == "")) {
       warnings <- c(warnings, paste("Some mschart rows have missing 'geom_type'. They may default to bar charts. Rows:", paste(chart_rows, collapse = ", ")))
    }
  }

  # -------------------------------------------------------------------------
  # 4. Report
  # -------------------------------------------------------------------------
  if (length(errors) > 0) {
    cat("\n[FAIL] Validation found critical errors:\n")
    cat(paste("-", errors, collapse = "\n"))
    return(invisible(FALSE))
  } else {
    cat("\n[PASS] content_metadata.csv looks valid.\n")
    if (length(warnings) > 0) {
      cat("\n[WARN] Warnings:\n")
      cat(paste("-", warnings, collapse = "\n"))
    }
    return(invisible(TRUE))
  }
}


