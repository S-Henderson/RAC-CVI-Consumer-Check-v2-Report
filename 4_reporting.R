library(tidyverse)
library(readxl)
library(dplyr)
library(janitor)
library(readr)
library(stringr)
library(openxlsx)

# By: Scott Henderson
# Last Updated: Apr 1, 2020

#--------------- CREATE EXCEPTIONS FILE ---------------

# Build exceptions file to send to app support
create_exceptions <- function(df) {
  df <- df %>%
    filter(
      `Raction` == "TAG"
    ) %>%
    select(
      `Transaction Number`
    ) %>%
    mutate(
      `Exception` = "TRUE",
      `Exception Reason` = "existing wearer",
      `Client` = "CVI"
    )
}

exceptions <- create_exceptions(df)

# Renames Transaction Number header to be used for app support tool
rename_exceptions_file <- function(df) {
  df <- df %>%
    rename(
      Transaction = `Transaction Number`
    )
}

exceptions <- rename_exceptions_file(exceptions)

#--------------- EXPORT EXCEPTIONS FILE ---------------

# Exceptions file to send to App Support
create_exceptions_file <- function() {
  # Create workbook
  wb <- createWorkbook()
  # Add sheets
  addWorksheet(wb, 
               sheetName = "Sheet1"
               )
  # Write exceptions to worksheet
  writeData(wb, 
            sheet = "Sheet1", 
            x = exceptions
            )
  # Exceptions filename
  exceptions_filename <- paste0("CVI Exceptions ", 
                                     format(Sys.Date(), 
                                     "%m-%d-%Y"), ".xlsx")
  # Save workbook
  saveWorkbook(wb, 
               exceptions_filename
               )
}

create_exceptions_file()

print("Exceptions File Exported")

#--------------- EXPORT BUILT FILE ---------------

# Get row and column index
last_row <- nrow(df)+1
all_cols <- 1:ncol(df)

# Create report file in excel with conditional formatting rules
create_report_workbook <- function() {
  # Create workbook
  wb <- createWorkbook()
  # Add sheets
  addWorksheet(wb, 
               sheetName = "Data"
               )
  # Write built df to worksheet
  writeData(wb, 
            sheet = "Data", 
            x = df
            )
  # Conditional Formatting Styles
  redStyle <- createStyle(fontColour = "#9C0006", 
                          bgFill = "#FFC7CE"
                          )
  yellowStyle <- createStyle(fontColour = "#9C6500", 
                             bgFill = "#FFEB9C"
                             )
  greenStyle <- createStyle(fontColour = "#006100", 
                            bgFill = "#C6EFCE"
                            )
  # Conditional Formatting Rules
  # Main rules
  conditionalFormatting(wb, 
                        sheet = "Data", 
                        cols = all_cols, 
                        rows = 1:last_row, 
                        type = "expression", 
                        rule = '$A1="TAG"', 
                        style = redStyle
                        )
  conditionalFormatting(wb, 
                        sheet = "Data", 
                        cols = all_cols, 
                        rows = 1:last_row, 
                        type = "expression", 
                        rule = '$A1="PREV TAG"', 
                        style = yellowStyle
                        )
  conditionalFormatting(wb, 
                        sheet = "Data", 
                        cols = all_cols, 
                        rows = 1:last_row, 
                        type = "expression", 
                        rule = '$A1="Diff Patient"', 
                        style = greenStyle
                        )
  # Misc rules
  conditionalFormatting(wb, sheet = "Data", 
                        cols = all_cols, 
                        rows = 1:last_row, 
                        type = "expression", 
                        rule = '$A1="BH TAG"', 
                        style = redStyle
                        )
  conditionalFormatting(wb, 
                        sheet = "Data", 
                        cols = all_cols, 
                        rows = 1:last_row, 
                        type = "expression",
                        rule = '$A1="IS"', 
                        style = greenStyle
                        )
  # Built report filename
  built_report_filename_xlsx <- paste0("Copy of RAC CVI Consumer Check v2 ", 
                                       format(Sys.Date(), "%m-%d-%Y"), 
                                       ".xlsx")
  # Save Workbook
  saveWorkbook(wb, 
               built_report_filename_xlsx
               )
}

create_report_workbook()

print("Built Report File Exported")

#--------------- TRACKER INFO ---------------

# Print tracker info
tracker_info <- function(df, df_exceptions) {
  # Counts total hits 
  print(paste0(nrow(df)," - Total Hits"))
  # Counts total actioned
  print(paste0(nrow(exceptions)," - Total Actioned"))
}

tracker_info(df, exceptions)

#--------------- OPEN APP SUPPORT LINK ---------------

# opens up service request portal website to send exceptions file to App Support -> opens in default browser
browseURL("https://360insights.atlassian.net/servicedesk/customer/portal/28/group/107/create/520")