# PowerPoint Automation with OfficeR

This project automates the generation of PowerPoint presentations using R and metadata files. It creates professional-looking presentations from structured data without manual copy-pasting.

## How It Works

The system uses three CSV files to control the presentation generation:

1. **slide_metadata.csv** - Defines the structure of slides (layout, master theme)
2. **content_metadata.csv** - Controls what content goes on each slide and where it's positioned
3. **data.csv** - Contains the actual data used in charts and tables

## Files

- `slide_metadata.csv`: Slide structure definitions
- `content_metadata.csv`: Content placement instructions
- `data.csv`: Source data for charts and tables
- `generate_presentation.R`: Core function library for presentation generation
- `create_presentation.R`: Main script to read CSV files and generate presentation
- `DATA_DICTIONARY.md`: Documentation of data structures and formats

## Requirements

- R (version 3.5.0 or higher)
- officer package
- dplyr package
- ggplot2 package
- readr package

## Installation

Install the required packages:

```r
install.packages(c("officer", "dplyr", "ggplot2", "readr"))
```

## Usage

Run the main script to generate the presentation:

```r
source("create_presentation.R")
```

This will create a file named `quarterly_report.pptx` in the project directory.

## Customization

To create your own presentations:

1. Modify `data.csv` with your own data
2. Update `slide_metadata.csv` to change slide layouts
3. Adjust `content_metadata.csv` to control content placement
4. Run the script to generate your customized presentation

## Project Structure

- `generate_presentation.R`: Contains the core `generate_presentation()` function that creates the PowerPoint presentation from data and metadata
- `create_presentation.R`: Main script that reads CSV files and calls the `generate_presentation()` function
- This separation allows for easier testing and reuse of the core presentation generation logic

## Metadata File Formats

### slide_metadata.csv
- `slide_number`: Sequential slide identifier
- `layout`: PowerPoint layout name (e.g., "Title and Content")
- `master`: Master theme to use

### content_metadata.csv
- `slide_number`: Which slide this content belongs to
- `position_type`: How to position content ("coordinates", "placeholder_type", "placeholder_name")
- `content_type`: Type of content ("text", "graph", "table")
- `placeholder_name`: Name of the placeholder to use (e.g., "Content Placeholder 1", "Title 1")
- `geom_type`: Type of ggplot geom to use ("geom_bar", "geom_line", "geom_point")
- `x_var`: Variable to use for x-axis (e.g., "Organization", "Date")
- `y_var`: Variable to use for y-axis (e.g., "Value")
- `fill_var`: Variable to use for fill/color grouping (e.g., "Organization")
- `left`, `top`, `width`, `height`: Position coordinates in inches (when position_type="coordinates")
- `metric`: Data metric to display (must match Metric column in data.csv)

### data.csv
- `Metric`: Name of the metric
- `Organization`: Category/organization name
- `Value`: Numeric value
- `Date`: Date of measurement

## Troubleshooting

### PowerPoint says the file needs repair

If PowerPoint indicates that the generated file needs repair, try these solutions:

1. Make sure all required R packages are up to date:
   ```r
   install.packages(c("officer", "dplyr", "ggplot2", "readr"))
   ```

2. Check that your CSV files are properly formatted with no missing values in required columns.

3. Ensure you're using a recent version of PowerPoint that supports the .pptx format.

4. Try opening the file in a different version of PowerPoint or on a different computer.

5. If issues persist, check the R console for any error messages during script execution.

### Error: "Found no placeholder of type..." or "Found no placeholder with label..."

This error occurs when the script tries to place content in a placeholder that doesn't exist in the slide layout. Make sure:

1. The layout names in `slide_metadata.csv` match exactly with PowerPoint layout names.
2. The placeholder names used in the `content_metadata.csv` file match what's available in each layout.
3. Refer to `DATA_DICTIONARY.md` for information about available placeholder names for each layout.