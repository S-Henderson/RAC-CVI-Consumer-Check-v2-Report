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
      #exception_reason == "existing wearer" ~ "PREV TAGGED",
      !is.na(exception_reason) ~ "PREV TAGGED",
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


# re-orders columns -> puts notes at start
df <- df %>%
  select(notes, everything())

# add patient name match after names

client_code	submission_email	program_type	created_date	transaction_number	session_number	program_code	model	invoice_number	status	customer_name	customer_address	customer_address_2	customer_city	customer_state	customer_zip_code	customer_phone_number	dealer	dealer_name	user_id	serial_number	on_hold_reason	sales_associate	comments	submission_type	is_blackhawk	previous_session_number	previous_claim_number	previous_claim_purchase_date	previous_claim_email	previous_claim_customer_name	previous_claim_model	previous_claim_status	purchase_sale_date	patient_first_name	patient_last_name	previous_patient_first_name	previous_patient_last_name


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
