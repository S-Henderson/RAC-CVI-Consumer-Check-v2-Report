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
mainDir <- file.path(Sys.getenv("USERPROFILE"),"Desktop")
subDir <- "RAC_CVI_Consumer_Check_v2_Exports"

# sets working directory to user desktop to add subfolder
setwd(mainDir)

#Check if the sub directory folder exists in the current directory, if not then creates it
ifelse(!dir.exists(subDir), dir.create(subDir), "Export directory already exists")

# sets working directory to sub path - this is where stuff is exported/saved
setwd(file.path(mainDir, subDir))

# reads excel file - opens file browser window
df <- read_excel(file.choose())

# reads excel file - hardcode path - optional
#df <- read_excel("RAC CVI Consumer Check v2 02-21-2020.xlsx", sheet = "Sheet1")

# cleans up column names
names(df) <- make_clean_names(names(df))

#--------------- PREP REPORT ---------------#

# removes duplicates by transaction_number
df <- df %>%
  distinct(transaction_number, .keep_all = TRUE)

# adds notes column -> calculates values based on case_when
df <- df %>%
  mutate(
    notes = case_when(
      previous_claim_status == "Invalid Submission" ~ "INVALID",
      is_blackhawk == "TRUE" ~ "BH TAG",
      # if exception_reason is not blank - accounts for mutiple exception reasons
      !is.na(exception_reason) ~ "PREV TAGGED",
      # TAG name matches -> report pulls same last name so just check first names for match
      patient_first_name == previous_patient_first_name ~ "TAG",
      patient_first_name != previous_patient_first_name ~ "DIFFERENT PATIENT"
    ))

# adds patient first name match column -> mostly for audit check - following process doc instructions
df <- df %>%
  mutate(
    patient_first_name_match = case_when(
      patient_first_name == previous_patient_first_name ~ "TRUE",
      patient_first_name != previous_patient_first_name ~ "FALSE"
    ))

# re-orders columns -> puts notes at start -> adds patient name match after patient names
df <- df %>%
  select(notes, 1:38, patient_first_name_match, everything())

# view dataframe in RStudio
View(df)

#--------------- BUILD EXCEPTIONS FILE ---------------#

# build exceptions file to send to app support
df_exceptions <- df %>%
  filter(notes == "TAG") %>%
  select(transaction_number) %>%
  mutate(
    Exception = "TRUE",
    'Exception Reason' = "existing wearer",
    Client = "CVI"
  )

# renames transaction_number to be used for app support tool
df_exceptions <- df_exceptions %>%
  rename(Transaction = transaction_number)

# view exception dataframe in RStudio
View(df_exceptions)

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

# create workbook
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

# conditional formatting rules to highlight excel rows based on notes value -> limit to 100 rows -> issues doing dynamic range for row
conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:100, type = "expression", rule = '$A1="TAG"', style = redStyle)
conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:100, type = "expression", rule = '$A1="PREV TAGGED"', style = yellowStyle)
conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:100, type = "expression", rule = '$A1="DIFFERENT PATIENT"', style = greenStyle)

# built report filename -> xlsx format to allow conditional formatting
built_report_filename_xlsx <- paste0("Copy of RAC CVI Consumer Check v2 ", format(Sys.Date(), "%m-%d-%Y"), ".xlsx")

# saves excel workbook
saveWorkbook(wb, built_report_filename_xlsx)

# counts total hits of built file for tracker
print(paste(nrow(df),"- Total Hits"))

# counts total actioned of exception file for tracker
print(paste(nrow(df_exceptions),"- Total Actioned"))

print("Script Completed")
