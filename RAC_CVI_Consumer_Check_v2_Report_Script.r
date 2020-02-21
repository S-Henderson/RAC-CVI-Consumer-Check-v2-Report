library(readxl)
library(dplyr)
library(janitor)
library(readr)
library(stringr)
library(tidyverse)
library(openxlsx)

#--------------- SETUP ---------------#

# directory paths
mainDir <- "C:/Users/shenderson/Desktop"
subDir <- "RAC_CVI_Consumer_Check_v2_Exports"

# sets working directory to sub path - this is where stuff is exported
setwd(mainDir)

#Check if the sub directory folder exists in the current directory, if not then creates it
ifelse(!dir.exists(subDir), dir.create(subDir), "Export directory already exists")

setwd(file.path(mainDir, subDir))

# reads excel file - opens file browser window
df <- read_excel(file.choose())

# reads excel file - hardcode path - optional
#df <- read_excel("RAC CVI Consumer Check v2 02-21-2020.xlsx", sheet = "Sheet1")

# cleans up column names
names(df) <- make_clean_names(names(df))

#--------------- PREP ---------------#

# removes duplicates by transaction_number
df <- df %>%
  distinct(transaction_number, .keep_all = TRUE)

# adds notes column -> calculates values based on case_when
df <- df %>%
  mutate(
    notes = case_when(
      previous_claim_status == "Invalid Submission" ~ "INVALID",
      is_blackhawk == "TRUE" ~ "BH TAG",
      exception_reason == "existing wearer" ~ "PREV TAGGED",
      patient_first_name == previous_patient_first_name ~ "TAG",
      patient_first_name != previous_patient_first_name ~ "DIFFERENT PATIENT"
    ))

# re-orders columns -> puts notes at start
df <- df %>%
  select(notes, everything())

# view dataframe in RStudio
View(df)

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

# write exception dataframe to csv
write.csv(df_exceptions, "CVI Exceptions MM-DD-YYYY.csv", row.names = FALSE)

#--------------- EXPORT TO EXCEL ---------------#

# create workbook
wb <- createWorkbook()

# add sheet named Data to workbook
addWorksheet(wb, "Data")

# colour font & fill styles
# find colour palette -> http://dmcritchie.mvps.org/excel/colors.htm
redStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
yellowStyle <- createStyle(bgFill = "#FFFF00")
greenStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")

# write df to Data worksheet
writeData(wb, "Data", x = df)

# conditional formatting rules to highlight excel rows based on notes value
conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:100, type = "expression", rule = '$A1="TAG"', style = redStyle)
conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:100, type = "expression", rule = '$A1="PREV TAGGED"', style = yellowStyle)
conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:100, type = "expression", rule = '$A1="DIFFERENT PATIENT"', style = greenStyle)

# saves excel workbook
saveWorkbook(wb, "RAC CVI Consumer Check v2 02-21-2020 - BUILT.xlsx")
