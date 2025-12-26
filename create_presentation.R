# Script to create PowerPoint presentation from CSV data and metadata

# Source the main presentation generation function
# Source the main presentation generation function
source("generate_presentation.R")
source("validate_metadata.R")

# -------------------------------------------------------------------------
# Validation
# -------------------------------------------------------------------------
if (!validate_metadata("content_metadata.csv")) {
  stop("Validation failed. Please fix content_metadata.csv before proceeding.")
}

# Read the CSV files
content_metadata <- read.csv("content_metadata.csv", stringsAsFactors = FALSE)
data <- read.csv("data.csv", stringsAsFactors = FALSE)

# Convert Date column to Date type
data$Date <- as.Date(data$Date)

# -------------------------------------------------------------------------
# Define Custom Objects (External ggplots or logic)
# -------------------------------------------------------------------------
# This section simulates an external R script that creates custom plots.
# Users can source their own files here: source("my_custom_plots.R")

custom_plots_list <- list()

# Example: Custom ggplot for slide 9
custom_plots_list[["graph1"]] <- ggplot(data[data$Metric == "Revenue", ], aes(x = Date, y = Value, color = Organization)) +
  geom_line() +
  geom_point() +
  theme_minimal() +
  labs(title = "Custom Revenue Trend Analysis (ggplot)", x = "Date", y = "Revenue ($)")

# -------------------------------------------------------------------------
# Generate Presentation
# -------------------------------------------------------------------------
# Pass the metadata, data, and the list of custom objects
# You can also pass 'template_path' if you have a specific .pptx template
presentation <- generate_presentation(content_metadata, data, external_objects = custom_plots_list)

# -------------------------------------------------------------------------
# Manual Fallback / customization
# -------------------------------------------------------------------------
# If the CSV driven approach isn't enough, you can manually add slides here
# presentation <- add_slide(presentation, layout = "Title Only", master = "Office Theme")
# presentation <- ph_with(presentation, value = "Manual Fallback Slide", location = ph_location_type(type = "title"))

# -------------------------------------------------------------------------
# Save
# -------------------------------------------------------------------------
print(presentation, target = "quarterly_report.pptx")
cat("PowerPoint presentation generated successfully as 'quarterly_report.pptx'\n")