TROUBLESHOOT <- "N"

# Author:bob.badgett@gmail.com
# Permissions:
#* Code GNU GPLv3 https://choosealicense.com/licenses/gpl-3.0/
#* Images CC BY-NC-SA 4.0 https://creativecommons.org/licenses/by-nc-sa/4.0/
# Optimized for coding with R Studio document outline view
# Last edited 2025-12-102025-12-15

#== Startup ======
library(tcltk) # For interactions and troubleshooting, part of base package so no install needed.
library(crayon)

#* Cleanup ======
#* # REMOVE ALL:
# rm(list = ls())

#* Set working directory -----
if (Sys.getenv("RSTUDIO") != "1"){
  args <- commandArgs(trailingOnly = FALSE)
  script_path <- sub("--file=", "", args[grep("--file=", args)])  
  script_path <- dirname(script_path)
  setwd(script_path)
}else{
  setwd(dirname(rstudioapi::getSourceEditorContext()$path))
}
getwd()

#* Troubleshooting options -----
options(error = NULL)   # Default

# Functions -----
normalize_text <- function(x) {
  x <- iconv(x, to = "UTF-8")        # ensure UTF-8
  x <- gsub("\u00A0", " ", x, fixed = TRUE)   # replace NBSP with normal space
  x <- gsub("\\s+", " ", x, perl = TRUE)      # collapse multiple spaces
  trimws(x)
}

## function_get_question_values ----
source("function_get_question_values-2025-12-13.R")

# Packages/libraries -----
library(openxlsx)
library(data.table)
library(dplyr)
library(tidyr)
library(stringr)
# Analysis
library(meta)

# Cleanup ======
#* Remove all but functions from the environment ------
rm(list = ls(envir = .GlobalEnv)[!sapply(ls(envir = .GlobalEnv), function(x) is.function(get(x, envir = .GlobalEnv)))])

# __________________________________------
# Data creation ===================================
## FEVS -----
###* FEVS Data frame specifications -----
#** https://www.opm.gov/fevs/public-data-file/
# OR adjust year in the URL below
#* https://www.opm.gov/fevs/reports/data-reports/data-reports/report-by-agency/2022/2022-agency-report.pdf
#* As an example, search for "(70)"
# "how satisfied": (70) Considering everything, how satisfied are you with your job?

### 1. data_FEVS_questions_included ------
# rm(data_FEVS_questions_included)
data_FEVS_questions_included <- read.xlsx(xlsxFile = "FEVS questions included - 2025-12-10.xlsx")

### 2. data_FEVS_temp_imported2 (temporary after feeds to data_FEVS) ------
data_FEVS_temp_imported2 <- data.frame()

### 3. all_responses_row (temporary after feeds to data_FEVS) ------
all_responses_row <- data.frame()

### 4. data_FEVS ------
###* Start populating the Year and xlsx filename for each year ------

Short_Names <- data_FEVS_questions_included$Short_Name
Short_Names <- Short_Names[!sapply(Short_Names, is.null) & !is.na(Short_Names) & Short_Names != ""]

# 1. List all xlsx files in the FEVS subdirectory
filenames <- NULL
filenames <- list.files(
  path = "FEVS",             # Place files in FEVS subdirectory
  pattern = "\\.xlsx$",
  full.names = TRUE
)
filenames <- filenames[!grepl("~", basename(filenames))]

years <- as.numeric(substr(basename(filenames), 1, 4))

# Create / populate data_FEVS
rm(data_FEVS)
data_FEVS <- data.frame(
  Year          = years,
  FileName_xlsx = filenames,
  stringsAsFactors = FALSE
)

###* Add columns from the FEVS questions listed in data_FEVS_questions_included -----
for (sn in Short_Names) {
  
  cols_to_add <- c(
    paste0(sn, "_qno"),
    paste0(sn, "_agencies"),
    paste0(sn, "_topbox_rate"),
    paste0(sn, "_rate"),
    paste0(sn, "_respondents")
  )
  
  for (col in cols_to_add) {
    if (!col %in% names(data_FEVS)) {
      data_FEVS[[col]] <- character(nrow(data_FEVS))
    }
  }
}

data_FEVS[ , -(1:2)] <- lapply(data_FEVS[ , -(1:2)], as.numeric)

###* FEVS data grab -----
#data_FEVS_row_number <- 0
#match_row_index <- NULL

# Troubleshooting selected questions
if (1==2){ # Troubleshooting (21) Employees in my work unit produce high-quality work."  should have gotten text but got nothing in 2022 and 2023
  data_FEVS_current_Year <- 2022
  data_FEVS_row_number <- 2
  data_FEVS_questions_included_current_qno <- 7
}

if (1==2){ # Troubleshooting examples
  data_FEVS_current_Year <- 2000
  value_qno_temp <- 8 # 2020
}

####* Cycle thru years start-----
for (i in seq_along(data_FEVS$FileName_xlsx)) {
  current_file <- data_FEVS$FileName_xlsx[i]
  #current_file <- data_FEVS$FileName_xlsx[5] # TROUBLESHOOTIN
  cat(crayon::black$bold("\n==========================================\nCycle through year: ", 
                         current_file,"\n", sep = ""))
  #####* TROUBLESHOOTING thru years -----
    if (current_file == data_FEVS$FileName_xlsx[2] && TROUBLESHOOT == "Y") {
      stop("\ First file completed, lets stop.")
    }
    
  #i <- 3 # TROUBLESHOOTIN -----
  data_FEVS_row_number   <- i 
  #data_FEVS_row_number <- 5 # TROUBLESHOOTIN -----
  if (i==1){
    data_FEVS_current_Year <- as.integer(substr(basename(data_FEVS$FileName_xlsx), 1, 4))
    data_FEVS_current_Year <- min(data_FEVS_current_Year, na.rm = TRUE)
    }
  # Below moved to the end of the loop
  #data_FEVS_current_Year <- data_FEVS_current_Year + 1
  current_file <- filenames[data_FEVS_row_number]
  
  cat(crayon::black$bold("\n__________________________________________\nNow opening data from: \"", 
                          current_file,"\"\n", sep = ""))
 # data_FEVS_questions_included_current_qno <- "Q15_2" # TROUBLESHOOT ------
  
  # Open the Excel file and overwrite data_FEVS_File_Index_Current completely
  data_FEVS_File_Index_Current <- read.xlsx(
    xlsxFile = current_file,
    sheet = "File_Index",
    rows = 1:20,          # Adjust the '20' if your header might be further down
    colNames = FALSE, # Use the content of the startRow as headers
    skipEmptyRows = FALSE
  )
  
  header_row_index <- which(data_FEVS_File_Index_Current[, 1] == "Item")
  
  data_FEVS_File_Index_Current <- read.xlsx(
    xlsxFile = current_file,
    sheet = "File_Index",
    startRow = header_row_index,
    colNames = TRUE # Use the content of the startRow as headers
  )
  
  ####* Cycle thru questions for current year -----
  for (data_FEVS_questions_included_current_qno in 1:nrow(data_FEVS_questions_included)) {
    #####* TROUBLESHOOTING thru questions -----
    if (data_FEVS_questions_included_current_qno == 3 && TROUBLESHOOT == "Y") {
      stop("\ First file completed, lets stop.")
    }
  
  # Open the Excel file and overwrite data_FEVS_File_Index_Current completely
    # Moved out of the loop

  #data_FEVS_questions_included_current_qno <- 10 # TROUBLESHOOT

    if (data_FEVS_questions_included_current_qno == 10){
      # Poor performers with different response structure and no period after question
      #target_text_to_match <- paste0(target_text_to_match,".")
    }
    #####* Look for target_text_to_match in the yearly file and Grab FEV question number if found -----
    target_text_to_match <- data_FEVS_questions_included$Item_Text[data_FEVS_questions_included_current_qno]
    target_text_to_match <-  normalize_text(target_text_to_match)
    # Below is 17 for poor performers
    match_row_index <- which(normalize_text(data_FEVS_File_Index_Current$Item.Text) == target_text_to_match)
    cat(crayon::black$black("\nmatch_row_index: for question ", 
                            data_FEVS_questions_included_current_qno, " (", target_text_to_match,"): ", match_row_index,"\n", sep = ""))
    
    if (length(match_row_index) > 0) {
      # If a match is found, extract and write the corresponding values
      value_qno_temp        <- data_FEVS_File_Index_Current$Item[match_row_index[1]]
      #####* TROUBLESHOOTING thru years -----
      if (current_file == data_FEVS$FileName_xlsx[2] && TROUBLESHOOT == "Y") {
        #stop("\ First file completed, lets stop.")
      }
      if (TROUBLESHOOT == "Y"){
        cat(crayon::magenta$bold(
          "\nvalue_qno_temp for ", value_Short_Name_temp, " (",") for ", data_FEVS_File_Index_Current$Item.Text, ":\n",
          value_qno_temp,
          sep = ""))
      }
      value_Short_Name_temp <- data_FEVS_questions_included$Short_Name[data_FEVS_questions_included_current_qno]
      data_FEVS[data_FEVS$FileName_xlsx == current_file, paste0(value_Short_Name_temp,"_qno")] <- value_qno_temp
      cat(crayon::green$bold("\nFound matching Item code: ",
                           value_qno_temp," for question \"", data_FEVS_questions_included_current_qno, " : ", 
                           target_text_to_match, "\"\n", sep = ""))
      function_get_question_values (value_qno_temp, value_Short_Name_temp) # First try just to get Short_Name 's rate
    } else {
      # Handle cases where no matching text is found
      value_qno_temp <- NA 
      cat(crayon::red$bold("\nNo matching Item Text found in  ",
                           current_file,"\nfor question ", data_FEVS_questions_included_current_qno  ,"\n", sep = ""))
      cat(crayon::black$bold("\nNOT starting function_get_question_values\n\n", sep = ""))
    }
  }
  data_FEVS_current_Year <- data_FEVS_current_Year + 1
}

# __________________________________------
# Write files -----
## csv ----
out_file <- sprintf("data_FEVS - %s.csv", Sys.Date())
write.csv(data_FEVS,out_file, row.names = FALSE)

## xlsx ----
out_file <- sprintf("data_FEVS - %s.xlsx", Sys.Date())

## Start workbook & worksheet ----
wb  <- createWorkbook()
addWorksheet(wb, "data_FEVS")

## 3. Write data (header in first row) ----

writeData(
  wb,
  sheet      = "data_FEVS",
  x          = data_FEVS,
  startRow   = 1,
  startCol   = 1,
  headerStyle = NULL  # we’ll style the header explicitly below
)

n_rows <- nrow(data_FEVS)
n_cols <- ncol(data_FEVS)

## 4. Define styles ----

headerStyle <- createStyle(
  fgFill         = "#D9D9D9",
  fontColour     = "black",
  textDecoration = "bold",
  halign         = "center",
  valign         = "center",
  border         = "TopBottomLeftRight",
  borderColour   = "lightgray",
  borderStyle    = "thin"
)

mediumGreenStyle <- createStyle(
  fgFill       = "#A9D18E",
  fontColour   = "black",
  border       = "TopBottomLeftRight",
  borderColour = "lightgray",
  borderStyle  = "thin"
)

lightGreenStyle <- createStyle(
  fgFill       = "#C6EFCE",
  fontColour   = "black",
  border       = "TopBottomLeftRight",
  borderColour = "lightgray",
  borderStyle  = "thin"
)

addStyle(
  wb,
  sheet      = "data_FEVS",
  style      = headerStyle,
  rows       = 1,
  cols       = 1:n_cols,
  gridExpand = TRUE
)

## 6. Determine column groups based on names ----
col_names <- names(data_FEVS)

# First two columns are Year and FileName_xlsx and are left uncolored
fixed_cols <- 1:2

# Columns whose names start with "Employees"
emp_cols <- which(grepl("^Employees", col_names))

# Columns before first "Employees" (beyond first two)
if (length(emp_cols) > 0) {
  first_emp <- min(emp_cols)
  last_emp  <- max(emp_cols)
  
  before_emp_cols <- setdiff(3:(first_emp - 1), fixed_cols)   # 3..(first_emp-1), skipping 1–2
  after_emp_cols  <- setdiff((last_emp + 1):n_cols, fixed_cols)
} else {
  # No Employees* columns: treat all columns from 3 onward as "before/after"
  before_emp_cols <- if (n_cols > 2) 3:n_cols else integer(0)
  after_emp_cols  <- integer(0)
}

## 7. Apply styles to data cells (rows 2..n_rows+1) ----

data_rows <- if (n_rows > 0) 2:(n_rows + 1) else integer(0)

# Medium green: Employees* columns
if (length(emp_cols) > 0 && length(data_rows) > 0) {
  addStyle(
    wb,
    sheet      = "data_FEVS",
    style      = mediumGreenStyle,
    rows       = data_rows,
    cols       = emp_cols,
    gridExpand = TRUE
  )
}

# Light green: before Employees columns (excluding first two columns)
if (length(before_emp_cols) > 0 && length(data_rows) > 0) {
  addStyle(
    wb,
    sheet      = "data_FEVS",
    style      = lightGreenStyle,
    rows       = data_rows,
    cols       = before_emp_cols,
    gridExpand = TRUE
  )
}

# Light green: after Employees columns
if (length(after_emp_cols) > 0 && length(data_rows) > 0) {
  addStyle(
    wb,
    sheet      = "data_FEVS",
    style      = lightGreenStyle,
    rows       = data_rows,
    cols       = after_emp_cols,
    gridExpand = TRUE
  )
}

## 8. Save xlsx ----
saveWorkbook(wb, out_file, overwrite = TRUE)

