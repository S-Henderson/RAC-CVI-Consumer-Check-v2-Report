library(readxl)
library(dplyr)
library(janitor)
library(readr)
library(stringr)
library(tidyverse)
library(openxlsx)

print("Script Starting")

#--------------- SETUP ---------------#

# directory paths
Desktop <- file.path(Sys.getenv("USERPROFILE"),"Desktop")
Export_Directory <- "RAC_CVI_Consumer_Check_v2_Exports"

#set directory to named export folder
set_directory_paths <- function(mainDir, subDir) {
  setwd(mainDir)
    ifelse(!dir.exists(subDir), dir.create(subDir), "Export directory already exists")
      setwd(file.path(mainDir, subDir))
        print(paste0("Current Working Directory is ", getwd()))
}

set_directory_paths(Desktop, Export_Directory)

# reads excel file - opens file browser window
df <- read_excel(file.choose())

# reads excel file - hardcode path - optional
#df <- read_excel("RAC CVI Consumer Check v2 02-21-2020.xlsx", sheet = "Sheet1")

# cleans up column names
clean_headers <- function(df) {
  make_clean_names(names(df))
}

names(df) <- clean_headers(df)

#--------------- PREP REPORT ---------------#

# removes duplicates by transaction_number
df <- df %>%
  distinct(transaction_number, .keep_all = TRUE)

# adds Raction column
Raction <- function(df) {
  df <- df %>%
    mutate(
      'Raction' = case_when(
        is_blackhawk == "TRUE" ~ "BH TAG",
        previous_claim_status == "Invalid Submission" ~ "IS",
        # if exception_reason is not blank - accounts for mutiple exception reasons
        !is.na(exception_reason) ~ "PREV TAG",
        # TAG name matches -> report pulls same last name so just check first names for match
        patient_first_name == previous_patient_first_name ~ "TAG",
        patient_first_name != previous_patient_first_name ~ "DIFFERENT PATIENT"
      ))
}

df <- Raction(df)

# adds patient first name match column -> mostly for audit check - following process doc instructions
patient_name_match <- function(df) {
  df <- df %>%
    mutate(
      patient_first_name_match = case_when(
        patient_first_name == previous_patient_first_name ~ "TRUE",
        patient_first_name != previous_patient_first_name ~ "FALSE"
      ))
}

df <- patient_name_match(df)

#--------------- ORDER OF COLUMNS ---------------#

# re-orders columns -> puts notes at start -> adds patient name match after patient names
df <- df %>%
  select('Raction', 1:38, patient_first_name_match, everything())

#--------------- BUILD EXCEPTIONS FILE ---------------#

# build exceptions file to send to app support
build_exceptions <- function(df) {
  df <- df %>%
    filter('Raction' == "TAG"
    ) %>%
    select(transaction_number
    ) %>%
    mutate(
      Exception = "TRUE",
      'Exception Reason' = "existing wearer",
      Client = "CVI"
    )
}

df_exceptions <- build_exceptions(df)

# renames transaction_number header to be used for app support tool
df_exceptions <- df_exceptions %>%
  rename(Transaction = transaction_number)

#--------------- EXPORT EXCEPTIONS FILE ---------------#

# exceptions filename for xlsx -> adds current date to filename
exceptions_filename_xlsx <- paste0("CVI Exceptions ", format(Sys.Date(), "%m-%d-%Y"), ".xlsx")

# write exception dataframe to csv
write.xlsx(df_exceptions, exceptions_filename_xlsx, sheetName = "Sheet1", row.names = FALSE)

### OPTIONAL CSV EXPORT ###

# exceptions filename for csv -> adds current date to filename
##exceptions_filename_csv <- paste0("CVI Exceptions ", format(Sys.Date(), "%m-%d-%Y"), ".csv")

# write exception dataframe to csv
##write.csv(df_exceptions, exceptions_filename_csv, row.names = FALSE)

#--------------- EXPORT BUILT FILE TO EXCEL ---------------#

# create excel workbook object
wb <- createWorkbook()

# add sheet named Data to workbook
addWorksheet(wb, "Data")

# colour font & fill styles for conditional formatting rules
## find colour palette -> http://dmcritchie.mvps.org/excel/colors.htm
redStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
yellowStyle <- createStyle(fontColour = "#9C5600", bgFill = "#FFEB9C")
greenStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")

# write df to Data worksheet
writeData(wb, "Data", x = df)

#--------------- CONDITIONAL FORMATTING RULES ---------------#

# conditional formatting rules to highlight excel rows based on notes value -> limit to 100 rows -> issues doing dynamic range for row
# main rules
conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:100, type = "expression", rule = '$A1="TAG"', style = redStyle)
conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:100, type = "expression", rule = '$A1="PREV TAG"', style = yellowStyle)
conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:100, type = "expression", rule = '$A1="DIFFERENT PATIENT"', style = greenStyle)
# misc rules
conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:100, type = "expression", rule = '$A1="IS"', style = greenStyle)

#--------------- SAVING EXCEL FILE ---------------#

# built report filename -> xlsx format to allow conditional formatting
built_report_filename_xlsx <- paste0("Copy of RAC CVI Consumer Check v2 ", format(Sys.Date(), "%m-%d-%Y"), ".xlsx")

# saves excel workbook
saveWorkbook(wb, built_report_filename_xlsx)

#--------------- TRACKER INFO ---------------#

# counts total hits of built file for tracker
print(paste(nrow(df),"- Total Hits"))

# counts total actioned of exception file for tracker
print(paste(nrow(df_exceptions),"- Total Actioned"))

# opens up service request portal website to send exceptions file -> opens in default browser
browseURL("https://360insights.atlassian.net/servicedesk/customer/portal/28/group/107/create/520")

print("Script Completed")
