# Example script showing how to create custom graphs for use with the PowerPoint automation system

# Load required libraries
library(ggplot2)
library(dplyr)
library(readr)
library(officer)

# Read the data
data <- read.csv("data.csv", stringsAsFactors = FALSE)

# Example 1: Create a revenue trend graph
graph1 <- data %>%
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

# Example 2: Create a profit margin comparison chart
graph2 <- data %>%
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

# Example 3: Create a customer satisfaction heatmap
graph3 <- data %>%
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

# Example 4: Create a market share pie chart
graph4 <- data %>%
  filter(Metric == "Market Share" & Date == max(Date)) %>%
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

# Example 5: Create a multi-metric dashboard
graph5 <- data %>%
  filter(Date == max(Date)) %>%
  ggplot(aes(x = Organization, y = Value, fill = Metric)) +
  geom_bar(stat = "identity", position = "dodge") +
  theme_minimal() +
  labs(
    title = "Latest Metrics by Organization",
    x = "Organization",
    y = "Value",
    fill = "Metric"
  ) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Example 6: Create a time series comparison
graph6 <- data %>%
  filter(Metric %in% c("Revenue", "Profit Margin")) %>%
  ggplot(aes(x = Date, y = Value, color = Metric, group = Metric)) +
  geom_line(size = 1.2) +
  facet_wrap(~Organization, scales = "free_y") +
  theme_minimal() +
  labs(
    title = "Revenue vs Profit Margin Over Time",
    x = "Date",
    y = "Value",
    color = "Metric"
  ) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Example 7: Create a comprehensive summary chart
graph7 <- data %>%
  filter(Date == max(Date)) %>%
  ggplot(aes(x = Organization, y = Value)) +
  geom_col(aes(fill = Metric), position = "dodge") +
  theme_minimal() +
  labs(
    title = "Q1 Performance Summary",
    x = "Organization",
    y = "Value",
    fill = "Metric"
  ) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Example of how to use these graphs in a presentation
# (This would typically be integrated into the main generate_presentation.R script)

create_sample_presentation <- function() {
  # Create a new presentation
  ppt <- read_pptx()
  
  # Add a title slide
  ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
  ppt <- ph_with(ppt, "Sample Presentation with Custom Graphs", location = ph_location_label("Title 1"))
  ppt <- ph_with(ppt, "This presentation demonstrates custom graphs", location = ph_location_label("Content 1"))
  
  # Add a slide with graph1
  ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
  ppt <- ph_with(ppt, "Revenue Trends", location = ph_location_label("Title 1"))
  ppt <- ph_with(ppt, value = graph1, location = ph_location_label("Content 1"))
  
  # Add a slide with graph2
  ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
  ppt <- ph_with(ppt, "Profit Margin Comparison", location = ph_location_label("Title 1"))
  ppt <- ph_with(ppt, value = graph2, location = ph_location_label("Content 1"))
  
  # Save the presentation
  print(ppt, target = "sample_presentation.pptx")
  cat("Sample presentation saved as sample_presentation.pptx\n")
}

# Uncomment the line below to run the example
# create_sample_presentation()