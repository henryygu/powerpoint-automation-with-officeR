# Data Dictionary

## slide_metadata.csv

| Column | Description | Data Type | Example |
|--------|-------------|-----------|---------|
| slide_number | Sequential slide identifier | Integer | 1 |
| layout | PowerPoint layout name | String | "Title and Content" |
| master | Master theme to use | String | "Office Theme" |

## content_metadata.csv

| Column | Description | Data Type | Example |
|--------|-------------|-----------|---------|
| slide_number | Which slide this content belongs to | Integer | 1 |
| position_type | How to position content | String | "coordinates", "placeholder_type", "placeholder_name" |
| content_type | Type of content | String | "text", "graph", "table" |
| placeholder_name | Name of the placeholder to use (e.g., "Content Placeholder 1", "Title 1") | String | "Content Placeholder 1" |
| geom_type | Type of ggplot geom to use | String | "geom_bar", "geom_line", "geom_point" |
| x_var | Variable to use for x-axis | String | "Organization", "Date" |
| y_var | Variable to use for y-axis | String | "Value" |
| fill_var | Variable to use for fill/color grouping | String | "Organization" |
| left | Left position in inches (when position_type="coordinates") | Numeric | 2.0 |
| top | Top position in inches (when position_type="coordinates") | Numeric | 1.5 |
| width | Width in inches (when position_type="coordinates") | Numeric | 4.5 |
| height | Height in inches (when position_type="coordinates") | Numeric | 4.5 |
| bg | Background color (reserved for future use) | String | "black" |
| metric | Data metric to display (must match Metric column in data.csv) | String | "Revenue" |

## data.csv

| Column | Description | Data Type | Example |
|--------|-------------|-----------|---------|
| Metric | Name of the metric | String | "Revenue" |
| Organization | Category/organization name | String | "Acme Corp" |
| Value | Numeric value | Numeric | 1250000 |
| Date | Date of measurement | Date | "2025-01-31" |

## Layout Types and Placeholder Information

### Title Slide
- Available placeholder types: "ctrTitle" (main title), "subTitle" (subtitle), "dt" (date), "ftr" (footer), "sldNum" (slide number)

### Title and Content
- Available placeholder types: "title" (slide title), "body" (content area)
- Available placeholder names: "Content Placeholder 1", "Title 1"

### Two Content
- Available placeholder types: "title" (slide title), "body" (left content), "other" (right content)
- Available placeholder names: "Content Placeholder 1", "Content Placeholder 2", "Title 1"