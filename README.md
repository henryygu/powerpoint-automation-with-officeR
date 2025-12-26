# PowerPoint Automation with officeR

This repository demonstrates how to automate the generation of PowerPoint presentations using the R `officer` and `mschart` packages, driven by a configuration CSV file (`content_metadata.csv`).

## Overview

The automation allows you to:
1.  **Define Slides & Content**: Control slide creation, layout, and content placement entirely via `content_metadata.csv`.
2.  **Flexible Positioning**: Place objects using PowerPoint placeholders (`placeholder_name`, `placeholder_type`) or precise coordinates (`left`, `top`, `width`, `height`).
3.  **Chart Support**: Use native Office charts (`mschart`) by default, or fallback to custom `ggplot` objects.
4.  **Styling**: Control text size, color, background, and element rotation.

## Workflow

1.  **Template (Optional)**: Create a PowerPoint template (`.pptx`) with your desired Master and Layouts.
2.  **Configuration**: Edit `content_metadata.csv` to define your presentation structure.
3.  **Custom Code**: If you need complex custom plots (ggplots), define them in an R script (e.g. `Custom_Plots.R`) and list them.
4.  **Generation**: Run `create_presentation.R`.

## Configuration: `content_metadata.csv`

The CSV file drives the generation process. Key columns:

| Column | Description |
| :--- | :--- |
| `slide_number` | Integer. Slides are generated in this order. |
| `slide_master` | Name of the Slide Master (e.g., "Office Theme"). |
| `slide_layout` | Name of the Slide Layout (e.g., "Title and Content", "Two Content"). |
| `position_type` | `placeholder_type`, `placeholder_name`, or `coordinates`. |
| `content_type` | `text`, `table`, `mschart_bar`, `mschart_line`, `custom`. |
| `placeholder_name` | Label of the placeholder (e.g., "Content Placeholder 2") or type (e.g., "title"). |
| `metric` | **Text**: content string (use `|` for bullets). **Charts**: data filter key. **Custom**: name of object in R list. |
| `geom_type` | For `mschart`: `geom_bar`, `geom_line`, etc. |
| `x_var`, `y_var`, `fill_var` | Column names in `data.csv` to map to the chart. |
| `left`, `top`, `width`, `height` | **Coordinates** (in inches) for custom positioning. |
| `rotation` | Rotation angle in degrees. |
| `font_size`, `font_color` | Styling for text elements. |

## Customization

### Using Custom ggplots
To use a ggplot instead of an mschart:
1.  In `content_metadata.csv`, set `content_type` to `custom` and `metric` to a unique name (e.g., "my_plot").
2.  In `create_presentation.R`, define your plot in the `custom_plots_list`:
    ```r
    custom_plots_list[["my_plot"]] <- ggplot(...)
    ```

### Manual Fallback
If the CSV automation does not cover a specific edge case, you can manually add slides in `create_presentation.R` before the `print()` statement using standard `officer` functions:
```r
presentation <- add_slide(presentation, layout = "Title Only")
presentation <- ph_with(presentation, value = "Manual Slide", location = ...)
```

## Example: Adding a New Revenue Chart Slide

Here is a step-by-step example of how to add a new slide representing **Revenue by Organization** using the existing `data.csv`.

### 1. Inspect your Data
Ensure `data.csv` contains the metrics you want to plot.
```csv
Metric,Organization,Value,Date
Revenue,Acme Corp,1250000,2025-01-31
...
```

### 2. Update `content_metadata.csv`
Open `content_metadata.csv` and add new rows for the slide. We will add specific rows to:
1.  Define the Slide properties (Master/Layout).
2.  Add a Title.
3.  Add the Bar Chart using specific coordinates.

**Add these lines to `content_metadata.csv`:**
```csv
11,Office Theme,Title Only,placeholder_type,text,title,,,,Example Revenue Chart,,,,,,,,,,
11,Office Theme,Title Only,coordinates,mschart_bar,,geom_bar,Organization,Value,Organization,Revenue,1,2,8,4.5,,,
```

**Explanation of the Chart Row:**
-   **`slide_number`**: `11` (New slide)
-   **`slide_layout`**: `Title Only` (Simple layout)
-   **`position_type`**: `coordinates` (We want precise control)
-   **`content_type`**: `mschart_bar` (Native Office Bar Chart)
-   **`x_var`**: `Organization` (X-axis category)
-   **`y_var`**: `Value` (Height of bars)
-   **`fill_var`**: `Organization` (Color bars by Org)
-   **`metric`**: `Revenue` (Filters `data.csv` to rows where Metric == 'Revenue')
-   **`left`, `top`, `width`, `height`**: `1`, `2`, `8`, `4.5` (Inches)

### 3. Run the Generation Script
Execute the main script in R:
```r
source("create_presentation.R")
```

The script will:
1.  **Validate** your CSV changes.
2.  **Generate** `quarterly_report.pptx`.
3.  **Place** the new Bar Chart at the exact coordinates specified.

## Requirements

- R
- `officer`
- `mschart`
- `dplyr`
- `ggplot2`
- `readr`
- `stringr`