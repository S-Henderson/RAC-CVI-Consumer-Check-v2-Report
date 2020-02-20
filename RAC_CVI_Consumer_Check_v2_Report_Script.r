library(readxl)
library(dplyr)
library(janitor)
library(readr)
library(stringr)
library(tidyverse)

library(openxlsx)


# reads excel file
df <- read_excel("RAC CVI Consumer Check v2 02-20-2020.xlsx", sheet = "Sheet1")

# cleans up column names
names(df) <- make_clean_names(names(df))

df <- df %>%
  # removes duplicates
  distinct(transaction_number, .keep_all = TRUE)

#View(df)

# adds notes column
df <- df %>%
  mutate(
    notes = case_when(
      previous_claim_status == "Invalid Submission" ~ "INVALID",
      is_blackhawk == "TRUE" ~ "BH TAG",
      exception_reason == "existing wearer" ~ "PREV TAGGED",
      patient_first_name == previous_patient_first_name ~ "TAG",
      patient_first_name != previous_patient_first_name ~ "DIFFERENT PATIENT"
))

View(df)

### export to excel
wb <- createWorkbook()

addWorksheet(wb, "Data")

negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
posStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")

## cells containing text
# df is object
writeData(wb, "Data", x = df)
conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:15, type = "expression", rule = "=$A1=""TAG""", style = posStyle)

conditionalFormatting(wb, "Data", cols = 1:52, rows = 1:15, type = "expression", rule = "=$AR1=0", style = negStyle)

#conditionalFormatting(wb, sheet, cols, rows, rule = NULL, style = NULL,
                      #type = "contains", ...)

#write.xlsx(wb, "test.xlsx")

saveWorkbook(wb, "conditionalFormattingExample.xlsx", TRUE)

