# Data Dictionary

This document describes the structure and meaning of the data and metadata files used in the PowerPoint automation system.

## Data File (data.csv)

Contains the actual data to be visualized in the presentation.

### Columns

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| Metric | String | The type of measurement | "Revenue", "Profit Margin" |
| Organization | String | The company or entity | "Acme Corp", "Globex Inc" |
| Value | Number | The measured value | 1250000, 0.24 |
| Date | Date | When the measurement was taken | "2025-01-31" |

## Slide Metadata (slide_metadata.csv)

Defines the overall structure of slides in the presentation.

### Columns

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| slide_number | Integer | Unique identifier for each slide | 1, 2, 3 |
| layout | String | PowerPoint slide layout | "Title and Content", "Two Content" |
| master | String | Slide master theme | "Office Theme" |

## Content Metadata (content_metadata.csv)

Specifies what content goes on each slide and where it should be positioned.

### Columns

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| slide_number | Integer | Reference to the slide | 1, 2, 3 |
| position_type | String | Positioning method | "coordinates", "object_name" |
| content_type | String | Type of content | "text", "graph", "table" |
| object_name | String | Placeholder name (when using object positioning) | "Title 1", "Content 2" |
| left | Number | Horizontal position in inches (when using coordinate positioning) | 2, 1.5 |
| top | Number | Vertical position in inches (when using coordinate positioning) | 1.5, 1 |
| width | Number | Width in inches (when using coordinate positioning) | 4.5, 5 |
| height | Number | Height in inches (when using coordinate positioning) | 4.5, 3 |
| bg | String | Background color (when using coordinate positioning) | "black", "white" |
| metric | String | Reference to data or text content | "graph1", "Quarterly Performance Report" |

## Content Types

### Text
Plain text content, typically used for titles and labels.

### Graph
Data visualizations created with ggplot2. The metric column references the graph object that should be created separately.

### Table
Tabular data displays. The metric column references the data subset to be displayed.

## Positioning Methods

### Object Name
Uses predefined placeholders in the PowerPoint template. Common object names include:
- "Title 1" - Main slide title
- "Content 1" - First content area
- "Content 2" - Second content area

### Coordinates
Explicit positioning using left, top, width, and height values in inches. This method provides precise control over element placement.

## Metric References

In the content metadata, the metric column can reference:
1. Text strings for text content types
2. Graph objects (e.g., "graph1", "graph2") for graph content types
3. Data subsets (e.g., "performance_data", "q1_data") for table content types

When creating custom presentations, ensure that any metric references have corresponding objects or data subsets defined in your R code.