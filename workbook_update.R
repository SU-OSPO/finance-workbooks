cat("Checking dependencies...\n")
if (!requireNamespace("pak", quietly = TRUE)) {
  install.packages("pak", repos = "http://cran.rstudio.com/")
}
pak::pak(c("openxlsx2", "dplyr", "optparse"))

suppressMessages(library(openxlsx2))
suppressMessages(library(dplyr))
suppressMessages(library(optparse))

# clear existing data in the environment if running this interactively
rm(list = ls())

# Parse command line options ####
parser <- OptionParser()
parser <- add_option(parser, c("-i", "--input"), default = "input",
                     help = paste("Relative path to the folder that contains",
                                  "the input files [default: %default]"))
parser <- add_option(parser, c("-o", "--output"), default = "output",
                     help = paste("Relative path to the folder that will",
                                  "contain the output files [default: %default]"))
parser <- add_option(parser, c("-r", "--ref"), default = "full_data",
                     help = paste("Relative path to the folder that contains",
                                  "the reference financial spreadsheets",
                                  "[default: %default]"))
args <- parse_args(parser)

# Read in full data ####
# these are the full data files downloaded from DataInsights
# make sure project_id is character to preserve leading zeros
# make sure account is numeric to avoid issues with workbook formulae
# need to check names because the last column doesn't have a name
options(warn = -1)

# check for files
if (!file.exists(file.path(args$ref, "BUD.xlsm"))) {
  stop("Reference file BUD.xlsm not found in ", args$ref)
}
if (!file.exists(file.path(args$ref, "COM.xlsm"))) {
  stop("Reference file COM.xlsm not found in ", args$ref)
}
if (!file.exists(file.path(args$ref, "EXP.xlsm"))) {
  stop("Reference file EXP.xlsm not found in ", args$ref)
}
bud_wb <- wb_load(file.path(args$ref, "BUD.xlsm"))
bud_df <- wb_to_df(bud_wb, skip_empty_rows = TRUE,
                   types = c(PROJECT_ID = "character", ACCOUNT = "numeric"),
                   check_names = TRUE)
com_wb <- wb_load(file.path(args$ref, "COM.xlsm"))
com_df <- wb_to_df(com_wb, skip_empty_rows = TRUE,
                   types = c(PROJECT_ID = "character", ACCOUNT = "numeric"),
                   check_names = TRUE)
exp_wb <- wb_load(file.path(args$ref, "EXP.xlsm"))
exp_df <- wb_to_df(exp_wb, skip_empty_rows = TRUE,
                   types = c(PROJECT_ID = "character", ACCOUNT = "numeric"),
                   check_names = TRUE)
options(warn = 0)

# Process Workbooks ####
if (!dir.exists(args$input)) {
  stop("The input folder '", args$input, "' does not exist")
}
# loop over all workbooks in a directory
to_process <- list.files(args$input, pattern = "\\.xlsx$")
if (length(to_process) == 0) {
  stop("No .xlsx files found in ", args$input)
}

# create output directory if it doesn't exist
if (!dir.exists(args$output)) {
  dir.create(args$output, recursive = TRUE)
}

for (wb_file in to_process) {
  cat(paste("Processing workbook:", wb_file, "\n"))

  ## Load workbook ####
  # NB: the xlsm file didn't have any macros, which caused issues with openxlsx2
  #     converted to xlsx for now
  grant_wb <- wb_load(file.path(args$input, wb_file))

  # get the project ids from the summary sheet
  award_info <- wb_to_df(grant_wb, sheet = "Award Info", start_row = 2,
                         skip_empty_rows = TRUE, skip_hidden_cols = TRUE,
                         check_names = TRUE)
  project_ids <- award_info$Project.ID

  ## Subset financial data ####
  # need to make sure the number of rows matches the existing workbook
  n_bud <- wb_to_df(grant_wb, sheet = "Budget Data", skip_empty_rows = FALSE) %>%
    nrow()
  bud_sub <- bud_df %>%
    filter(PROJECT_ID %in% project_ids) %>%
    add_row(PROJECT_ID = rep(NA, n_bud - nrow(.)))

  n_com <- wb_to_df(grant_wb, sheet = "Commitments Data", skip_empty_rows = FALSE) %>%
    nrow()
  com_sub <- com_df %>%
    filter(PROJECT_ID %in% project_ids) %>%
    add_row(PROJECT_ID = rep(NA, n_com - nrow(.)))

  exp_sub <- exp_df %>%
    filter(PROJECT_ID %in% project_ids)

  ## Update data ####
  # write to the Budget Data sheet in the workbook
  grant_wb$add_data(sheet = "Budget Data", x = bud_sub, start_row = 2,
                    col_names = FALSE, na.strings = "")

  # write to the Commitments Data sheet in the workbook
  grant_wb$add_data(sheet = "Commitments Data", x = com_sub, start_row = 2,
                    col_names = FALSE, na.strings = "")

  # write to the Expense Data sheet in the workbook
  grant_wb$add_data(sheet = "Expense Data", x = exp_sub, start_row = 2,
                    col_names = FALSE, na.strings = "")

  ## Write workbook ####
  wb_file_write <- gsub("\\d{1,2}\\.\\d{1,2}\\.\\d{2}",
                        format(Sys.Date(), "%m.%d.%y"),
                        wb_file)
  grant_wb$save(file.path(args$output, wb_file_write),
                overwrite = TRUE)
}
cat("All done!\n")
