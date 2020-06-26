# By: Scott Henderson
# Last Updated: June 26, 2020

# Purpose: Identify existing wearers and track them, in case the number of existing wearers claiming for the new wearers bonus ever need to be quantified for the client. 

#--------------- LOAD LIBRARIES ---------------

if (!require("pacman")) install.packages("pacman"); library(pacman)

p_load("tidyverse", "readxl", "openxlsx") 

#--------------- SCRIPT STARTING ---------------

start_time <- format(Sys.time(), "%X")

print(paste0("Script Starting at ", start_time))

#--------------- IMPORT DATA ---------------

# File name is: "RAC CVI Consumer Check v2 %m-%d-%Y.xlsx"

# Select file -----

src_file_path <- paste0(Sys.getenv(c("USERPROFILE")),"\\Downloads")

src_file_pattern <- "^RAC CVI Consumer Check v2(.*)xlsx$"

src_file_name = list.files(path = src_file_path,
                           pattern = src_file_pattern,
                           full.names = TRUE)

print(paste0("File Selected is -> ", src_file_name))

# Import file -----

df <- read_excel(src_file_name,
                 sheet = "Sheet1",
                 guess_max = Inf)

#--------------- CLEAN DATA ---------------

# Transactions -----

df$`Transaction Number` <- as.numeric(df$`Transaction Number`)

# Patient Names -----

# Names to Uppercase to properly match/compare
df$`Patient First Name` <- toupper(df$`Patient First Name`)

df$`Patient Last Name` <- toupper(df$`Patient Last Name`)

df$`Previous Patient First Name` <- toupper(df$`Previous Patient First Name`)

df$`Previous Patient Last Name` <- toupper(df$`Previous Patient Last Name`)


# convert_to_uppercase <- function(df, column1) {
#   
#   column1 <- enquo(column1)
#   
#   df %>% 
#     !!column1 <- toupper(!!column1)
# }
# 
# df <- convert_to_uppercase(df, quo(`Patient Last Name`))

# Remove duplicates -----

# Query pulls multiple same transactions, only need to action it once
df <- df %>% 
  distinct(
    `Transaction Number`, 
    .keep_all = TRUE
  )

#--------------- ANALYSIS ---------------

# Raction -----

# Document actions to take
df <- df %>%
  mutate(
    `Raction` = case_when(
      
      `Is Blackhawk` %in% c("TRUE", "True")                 ~ "BH TAG",
      
      `Previous Claim Status`  == "Invalid Submission"      ~ "IS",
      
      !is.na(`Exception Reason`)                            ~ "PREV TAG",    # Accounts for multiple exception reasons
      
      `Patient First Name` == `Previous Patient First Name` ~ "TAG",         # Checks First Name as query pulls same Last Name matches
      `Patient First Name` != `Previous Patient First Name` ~ "Diff Patient" 
      
    )
  )

# Patient Name Match -----

# Patient name match column -> mostly for audit check guide
df <- df %>%
  mutate(
    `Patient First Name Match` = case_when(
      
      `Patient First Name` == `Previous Patient First Name` ~ "TRUE",
      `Patient First Name` != `Previous Patient First Name` ~ "FALSE"
    )
  )

# Re-order columns -----

# Put Raction at start to read more easily & put name match check after names
df <- df %>%     
  select(
    `Raction`, 
    1:40, 
    `Patient First Name Match`,
    everything()
  )

#--------------- REPORTING ---------------

# Setwd function -----

# To reset to current working directory throughout report
old_dir <- getwd()

# Save file to server drive (if connected)
connect_to_server <- function(server_path) {
  
  old_dir <- getwd()
  new_dir <- server_path
  
  if (dir.exists(new_dir)) {
    
    setwd(new_dir)
    print("Connected to Z-Drive Server - Saving File on Server")
    
  } else {
    
    print("Cannot connect to Z-Drive Server. Will Save File To Working Directory")
    print(paste0("Working Directory -> ", old_dir))
    
  }
}

# Build exceptions file -----

# Only send TAG checks
df_exceptions <- df %>%
  filter(
    `Raction` == "TAG"
    
  ) %>%
  
  select(
    `Transaction Number`,
    `Claim Amount`
    
  ) %>%
  
  mutate(
    `Exception` = "TRUE",
    `Exception Reason` = "existing wearer",
    `Client` = "CVI"
    
  ) %>% 
  
  rename(
    `Transaction` = `Transaction Number` # Standardize header template for App Support Tools
  )

# Worked claims sum -----

# For report tracking
df_exceptions %>%
  summarize(
    `Sum of TAG Transactions` = sum(`Claim Amount`)
  )

# Save exceptions file -----

print("Attempting to save CVI Exceptions File")

connect_to_server("\\\\360Corp-WShare\\zdrive\\RAC\\RAC Report Archive\\RAC CVI Consumer Check v2\\CVI Exceptions Folder\\")

# Add date to file name
exceptions_file_name <- paste0("CVI Exceptions ", 
                               format(Sys.Date(), "%m-%d-%Y"), 
                               ".xlsx")

write.xlsx(
  x = df_exceptions,
  file = exceptions_file_name,
  sheetName = "Sheet1"
)

# Reset working directory -----
setwd(old_dir)

# Data validation list reference -----

# List of values for data validation options in excel export
data_validation_list <- data.frame(
  
  `Raction Reasons` = c("BH TAG", 
                        "IS", 
                        "PREV TAG", 
                        "TAG", 
                        "Diff Patient"),
  
  # To allow spaces in header
  check.names = FALSE
  
)

# Build report file -----

print("Attempting to save Built Report File")

connect_to_server("\\\\360Corp-WShare\\zdrive\\RAC\\RAC Report Archive\\RAC CVI Consumer Check v2\\")

# Get row and column index
last_row <- nrow(df)+1
all_cols <- 1:ncol(df)

# Prep workbook

wb <- createWorkbook()

addWorksheet(
  wb, 
  sheetName = "Data"
)

addWorksheet(
  wb, 
  sheetName = "Data Validation"
)

writeData(
  wb, 
  sheet = "Data", 
  x = df,
  withFilter = TRUE
)

setColWidths(
  wb, 
  sheet = "Data", 
  cols = 1:1, # Raction column
  widths = 10
)

# Conditional formatting styles

red_style <- createStyle(
  fontColour = "#9C0006", 
  bgFill = "#FFC7CE"
)

yellow_style <- createStyle(
  fontColour = "#9C6500", 
  bgFill = "#FFEB9C"
)

green_style <- createStyle(
  fontColour = "#006100", 
  bgFill = "#C6EFCE"
)

# Conditional formatting rules 

# Main rules
conditionalFormatting(
  wb, 
  sheet = "Data", 
  cols = all_cols, 
  rows = 1:last_row, 
  type = "expression", 
  rule = '$A1="TAG"', 
  style = red_style
)

conditionalFormatting(
  wb, 
  sheet = "Data", 
  cols = all_cols, 
  rows = 1:last_row, 
  type = "expression", 
  rule = '$A1="PREV TAG"', 
  style = yellow_style
)

conditionalFormatting(
  wb, 
  sheet = "Data", 
  cols = all_cols, 
  rows = 1:last_row, 
  type = "expression", 
  rule = '$A1="Diff Patient"', 
  style = green_style
)

# Misc rules
conditionalFormatting(
  wb, 
  sheet = "Data", 
  cols = all_cols, 
  rows = 1:last_row, 
  type = "expression", 
  rule = '$A1="BH TAG"', 
  style = red_style
)

conditionalFormatting(
  wb, 
  sheet = "Data", 
  cols = all_cols, 
  rows = 1:last_row, 
  type = "expression",
  rule = '$A1="IS"', 
  style = green_style
)

# Data validation code

dataValidation(
  wb, 
  sheet = "Data", 
  col = 1, 
  rows = 2:last_row,
  type = "list", 
  value = "'Data Validation'!$A$2:$A$6" # Watch out for how many items on data validation list
)


built_report_filename <- paste0("Copy of RAC CVI Consumer Check v2 ", 
                                format(Sys.Date(), "%m-%d-%Y"), 
                                ".xlsx"
)

saveWorkbook(
  wb, 
  file = built_report_filename
)

# Tracking info -----

print(paste0(nrow(df)," - Total Hits"))                # Counts total hits of report
print(paste0(nrow(df_exceptions)," - Total Actioned")) # Counts total actioned (TAG transactions)

# Data check -----

# Check if any Raction reasons/tags are missing
na_Raction <- sum(is.na(df$`Raction`))

print(if_else(na_Raction > 0, 
              "Missing Raction Reasons - Please Check Data",    # Missing data
              "No Raction Reasons Missing - Continue Forward")) # NOT missing data

# Reset working directory -----
setwd(old_dir)

#--------------- SCRIPT COMPLETED ---------------

end_time <- format(Sys.time(), "%X")

print(paste0("Script Completed at ", end_time))