library(tidyverse)
library(readxl)
library(dplyr)
library(janitor)
library(readr)
library(stringr)
library(openxlsx)

# By: Scott Henderson
# Last Updated: March 27, 2020

# Input: RAC CVI Consumer Check v2 Report
# Output: 
# Totals Hits & Total Actioned tracker info
# Built Report .xlsx file into a folder on your Desktop named RAC_CVI_Consumer_Check_v2_Exports
# If any TAG transactions, CVI Exceptions File to be sent to App Support

start_time <- format(Sys.time(), "%X")

print(paste0("Script Starting at ", start_time))

#--------------- SETUP ---------------

# directory paths
Desktop <- file.path(Sys.getenv("USERPROFILE"),"Desktop")
Export_Directory <- "RAC_CVI_Consumer_Check_v2_Exports"

# set directory to named export folder
set_directory_paths <- function(mainDir, subDir) {
  setwd(mainDir)
  ifelse(!dir.exists(subDir), dir.create(subDir), "Export Directory already exists")
  setwd(file.path(mainDir, subDir))
  print(paste0("Current Working Directory is ", getwd()))
}

set_directory_paths(Desktop, Export_Directory)

# reads excel file - opens file browser window
df <- read_excel(
        file.choose(), 
        sheet = "Sheet1",
        col_types = "text",
        guess_max = Inf
)

print("Raw Data File Imported")

#--------------- PREP REPORT ---------------

# removes duplicates by transaction_number
remove_duplicates <- function(df) {
  df <- df %>%
    distinct(`Transaction Number`, .keep_all = TRUE)
}

df <- remove_duplicates(df)

#--------------- CLEAN TRANSACTIONS ---------------

df$`Transaction Number` = as.numeric(as.character(df$`Transaction Number`))

#--------------- CLEAN PATIENT NAMES ---------------

df$`Patient First Name` = toupper(df$`Patient First Name`)

df$`Patient Last Name` = toupper(df$`Patient Last Name`)

df$`Previous Patient First Name` = toupper(df$`Previous Patient First Name`)

df$`Previous Patient Last Name` = toupper(df$`Previous Patient Last Name`)

#--------------- CLEAN DATE ---------------

#df$`Created Date` = 


# adds Raction column
create_Raction <- function(df) {
  df <- df %>%
    mutate(
      `Raction` = case_when(
        `Is Blackhawk` == "TRUE" ~ "BH TAG",
        `Previous Claim Status`  == "Invalid Submission" ~ "IS",
        # if exception_reason is not blank - accounts for mutiple exception reasons
        !is.na(`Exception Reason`) ~ "PREV TAG",
        # TAG name matches -> report pulls same last name so just check first names for match
        `Patient First Name` == `Previous Patient First Name` ~ "TAG",
        `Patient First Name` != `Previous Patient First Name` ~ "Diff Patient"
      ))
}

df <- create_Raction(df)

# adds patient first name match column -> mostly for audit check - following process doc instructions
patient_name_match <- function(df) {
  df <- df %>%
    mutate(
      `Patient First Name Match` = case_when(
        `Patient First Name` == `Previous Patient First Name` ~ "TRUE",
        `Patient First Name` != `Previous Patient First Name` ~ "FALSE"
      ))
}

df <- patient_name_match(df)

#--------------- ORDER OF COLUMNS ---------------

# re-orders columns -> puts Raction at start -> adds patient name match after patient names
reorder_df_columns <- function(df) {
  df <- df %>%
    select(`Raction`, 1:38, `Patient First Name Match`, everything())
}

df <- reorder_df_columns(df)

#--------------- BUILD EXCEPTIONS FILE ---------------

# build exceptions file to send to app support
build_exceptions <- function(df) {
  df <- df %>%
    filter(`Raction` == "TAG"
    ) %>%
    select(`Transaction Number`
    ) %>%
    mutate(
      Exception = "TRUE",
      `Exception Reason` = "existing wearer",
      Client = "CVI"
    )
}

df_exceptions <- build_exceptions(df)

# renames transaction_number header to be used for app support tool
rename_exceptions_file <- function(df) {
  df <- df %>%
    rename(Transaction = `Transaction Number`)
}

df_exceptions <- rename_exceptions_file(df_exceptions)

#--------------- EXPORT EXCEPTIONS FILE ---------------

# create exceptions excel file to send to app support
create_exceptions_file <- function() {
  #--------------- CREATE EXCEL FILE ---------------
  # create excel workbook object
  wb <- createWorkbook()
  # add sheet to workbook
  addWorksheet(wb, "Sheet1")
  # write df to worksheet
  writeData(wb, "Sheet1", x = df_exceptions)
  #--------------- SAVING EXCEL FILE ---------------
  # exceptions filename for xlsx -> adds current date to filename
  exceptions_filename_xlsx <- paste0("CVI Exceptions ", format(Sys.Date(), "%m-%d-%Y"), ".xlsx")
  # saves excel workbook
  saveWorkbook(wb, exceptions_filename_xlsx)
}

create_exceptions_file()

#--------------- OPTIONAL - EXPORT EXCEPTIONS FILE ---------------

### OPTIONAL SIMPLE XLSX EXPORT ###

# write exception dataframe to excel file
##write.xlsx(df_exceptions, exceptions_filename_xlsx, sheetName = "Sheet1", row.names = FALSE)

### OPTIONAL CSV EXPORT ###

# exceptions filename for csv -> adds current date to filename
##exceptions_filename_csv <- paste0("CVI Exceptions ", format(Sys.Date(), "%m-%d-%Y"), ".csv")

# write exception dataframe to csv
##write.csv(df_exceptions, exceptions_filename_csv, row.names = FALSE)

#--------------- EXPORT BUILT FILE TO EXCEL ---------------

# get row and column index
last_row <- nrow(df)+1
all_cols <- 1:ncol(df)

# create report file in excel with conditional formatting rules
create_report_workbook <- function() {
  #--------------- CREATE EXCEL FILE ---------------
  # create excel workbook object
  wb <- createWorkbook()
  # add sheet named Data to workbook
  addWorksheet(wb, "Data")
  # write df to Data worksheet
  writeData(wb, "Data", x = df)
  #--------------- CONDITIONAL FORMATTING STYLES ---------------
  # colour font & fill styles for conditional formatting rules -> find colour palette -> http://dmcritchie.mvps.org/excel/colors.htm
  redStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
  yellowStyle <- createStyle(fontColour = "#9C6500", bgFill = "#FFEB9C")
  greenStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")
  #--------------- CONDITIONAL FORMATTING RULES ---------------
  # conditional formatting rules to highlight excel rows based on Raction value -> limit to 100 rows -> issues doing dynamic range for row
  # main rules
  conditionalFormatting(wb, "Data", cols = all_cols, rows = 1:last_row, type = "expression", rule = '$A1="TAG"', style = redStyle)
  conditionalFormatting(wb, "Data", cols = all_cols, rows = 1:last_row, type = "expression", rule = '$A1="PREV TAG"', style = yellowStyle)
  conditionalFormatting(wb, "Data", cols = all_cols, rows = 1:last_row, type = "expression", rule = '$A1="Diff Patient"', style = greenStyle)
  # misc rules
  conditionalFormatting(wb, "Data", cols = all_cols, rows = 1:last_row, type = "expression", rule = '$A1="BH TAG"', style = redStyle)
  conditionalFormatting(wb, "Data", cols = all_cols, rows = 1:last_row, type = "expression", rule = '$A1="IS"', style = greenStyle)
  #--------------- SAVING EXCEL FILE ---------------
  # built report filename -> xlsx format to allow conditional formatting
  built_report_filename_xlsx <- paste0("Copy of RAC CVI Consumer Check v2 ", format(Sys.Date(), "%m-%d-%Y"), ".xlsx")
  # saves excel workbook
  saveWorkbook(wb, built_report_filename_xlsx)
}

create_report_workbook()

print("Built File Exported")

#--------------- TRACKER INFO ---------------

# print tracker info
tracker_info <- function(df, df_exceptions) {
  # counts total hits of built file for tracker
  print(paste0(nrow(df)," - Total Hits"))
  # counts total actioned of exception file for tracker
  print(paste0(nrow(df_exceptions)," - Total Actioned"))
}

tracker_info(df, df_exceptions)

# opens up service request portal website to send exceptions file to app support -> opens in default browser
browseURL("https://360insights.atlassian.net/servicedesk/customer/portal/28/group/107/create/520")

#--------------- SCRIPT COMPLETED ---------------

end_time <- format(Sys.time(), "%X")

print(paste0("Script Completed at ", end_time))
