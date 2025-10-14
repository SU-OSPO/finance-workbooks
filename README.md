# Repository Structure

```
.
├── input (includes all award workbooks)
├── ouput (where updated workbooks will be saved)
├── full_data (where the DataInsights workbooks go)
├── workbook_update.R (the R script)
└── workbook_update.bat (double-click this on Windows to run the R script)
```

# Dependencies

The following must be installed manually in order to run the script:
- [R](https://cloud.r-project.org/)
- (Optional) [RStudio](https://posit.co/download/rstudio-desktop/)

The following will be installed by the script:
- The [`pak` R package](https://pak.r-lib.org/) for updating dependencies
- The [`openxlsx2` R package](https://janmarvin.github.io/openxlsx2/) for reading and writing .xlsx files
- The [`dplyr` R package](https://dplyr.tidyverse.org/) for data filtering and handling
- The [`optparse` R package](https://trevorldavis.com/R/optparse/) for handling optional command-line arguments

# Instructions

1. Place old/prepped workbooks in the `input` folder (__should be saved as .xlsx files__)
2. (Optional) Clean out old workbooks from the `output` folder
3. Download updated raw data workbooks from DataInsights and place them in the `full_data` folder:
   - The budget data workbook should be named `BUD.xlsm`
   - The expense data workbook should be named `EXP.xlsm`
   - The commitments data workbook should be named `COM.xlsm`
4. Run the script (using one of these options):
   - Open the `workbook_update.R` script in RStudio, select all of the text (e.g., Ctrl + A), then click Run.
   - Open a terminal window and run `Rscript workbook_update.R`
   - Double click `workbook_update.bat` (only on Windows)

## Optional arguments


# What the script does
