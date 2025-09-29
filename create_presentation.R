# Script to create PowerPoint presentation from CSV data and metadata
# This script reads the CSV files and generates the presentation

# Source the main presentation generation function
source("generate_presentation.R")

# Read the CSV files
slide_metadata <- read.csv("slide_metadata.csv", stringsAsFactors = FALSE)
content_metadata <- read.csv("content_metadata.csv", stringsAsFactors = FALSE)
data <- read.csv("data.csv", stringsAsFactors = FALSE)

# Convert Date column to Date type
data$Date <- as.Date(data$Date)

# Generate the presentation
presentation <- generate_presentation(slide_metadata, content_metadata, data)

# Save the presentation
print(presentation, target = "quarterly_report.pptx")

cat("PowerPoint presentation generated successfully as 'quarterly_report.pptx'\n")