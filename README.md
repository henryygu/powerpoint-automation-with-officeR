# PowerPoint Automation with R

This project automates the creation of PowerPoint presentations using R and metadata files. It allows users to generate professional presentations from structured data without manually creating each slide.

## Overview

The system uses three main components:
1. **Data File** (`data.csv`) - Contains the actual data to be visualized
2. **Slide Metadata** (`slide_metadata.csv`) - Defines slide layouts and masters
3. **Content Metadata** (`content_metadata.csv`) - Specifies what content goes on each slide and where

## Files

- `data.csv` - Sample organizational data with metrics, organizations, values, and dates
- `slide_metadata.csv` - Defines slide layouts and masters
- `content_metadata.csv` - Specifies content placement on each slide
- `generate_presentation.R` - R script that creates the PowerPoint presentation

## Requirements

- R (version 3.5.0 or higher)
- R packages:
  - `officer`
  - `dplyr`
  - `ggplot2`
  - `readr`

## Installation

1. Install R from [CRAN](https://cran.r-project.org/)
2. Install required packages:
```r
install.packages(c("officer", "dplyr", "ggplot2", "readr"))
```

## Usage

1. Modify the metadata files to define your presentation structure:
   - Update `slide_metadata.csv` to specify slide layouts
   - Update `content_metadata.csv` to define content placement
   - Update `data.csv` with your own data

2. Run the R script:
```r
source("generate_presentation.R")
```

3. The generated presentation will be saved as `generated_presentation.pptx`

## Customization

For more advanced usage, you can create your own custom graphs and reference them in the metadata:

1. Create your graph objects in a separate R script (see `example.R` for examples)
2. Modify the `create_sample_graphs` function in `generate_presentation.R` to include your custom graphs
3. Reference your graph objects by name in the `metric` column of `content_metadata.csv`

## Metadata File Structure

### Slide Metadata (`slide_metadata.csv`)
- `slide_number` - Unique identifier for each slide
- `layout` - PowerPoint slide layout (e.g., "Title and Content", "Two Content")
- `master` - Slide master theme (e.g., "Office Theme")

### Content Metadata (`content_metadata.csv`)
- `slide_number` - Reference to the slide
- `position_type` - Either "coordinates" or "object_name"
- `content_type` - Type of content ("text", "graph", or "table")
- `object_name` - Placeholder name when using object positioning
- `left`, `top`, `width`, `height` - Position coordinates (inches) when using coordinate positioning
- `bg` - Background color
- `metric` - Reference to data or text content

## Customization

### Adding New Slides
1. Add new rows to `slide_metadata.csv` with the slide number, layout, and master
2. Add corresponding content rows to `content_metadata.csv`

### Changing Layouts
Modify the `layout` column in `slide_metadata.csv` to use different PowerPoint layouts

### Adding Content
Add new rows to `content_metadata.csv` specifying:
- Slide number
- Positioning method (coordinates or object name)
- Content type
- Position details
- Content reference

## Examples

### Text Content
```
slide_number,position_type,content_type,object_name,left,top,width,height,bg,metric
1,object_name,text,Title 1,,,,,,Quarterly Performance Report
```

### Graph Content with Coordinate Positioning
```
slide_number,position_type,content_type,object_name,left,top,width,height,bg,metric
1,coordinates,graph,,2,1.5,4.5,4.5,black,graph1
```

### Table Content with Coordinate Positioning
```
slide_number,position_type,content_type,object_name,left,top,width,height,bg,metric
2,coordinates,table,,1,1,5,3,white,performance_data
```

## Troubleshooting

### Package Not Found
If you get errors about missing packages, ensure you've installed all required packages:
```r
install.packages(c("officer", "dplyr", "ggplot2", "readr"))
```

### File Not Found
Ensure all metadata and data files are in the same directory as the R script.

### Layout Issues
Make sure the layout names in `slide_metadata.csv` match those available in your PowerPoint installation.

## License

This project is open source and available under the MIT License.