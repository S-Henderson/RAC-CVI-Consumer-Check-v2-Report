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

exceptions_df <- create_exceptions(df)

print("Creating Exceptions File")

#--------------- RENAME EXCEPTIONS FILE HEADERS ---------------

# Renames Transaction Number header to be used for app support tool
rename_exceptions_file <- function(df) {
  df <- df %>%
    rename(
      `Transaction` = `Transaction Number`
    )
}

exceptions_df <- rename_exceptions_file(exceptions_df)

print("Cleaning Exceptions File")

#--------------- EXPORT EXCEPTIONS FILE ---------------

# Network drive path to report Archive
exceptions_file_path <- "\\\\360Corp-WShare\\zdrive\\RAC\\RAC Report Archive\\RAC CVI Consumer Check v2\\CVI Exceptions Folder\\"

# Add date to file name
exceptions_file_name <- paste0("CVI Exceptions ", 
                               format(Sys.Date(), "%m-%d-%Y"), 
                               ".xlsx"
                               )

# Save exceptions file to Network Drive Archive
write.xlsx(
  x = exceptions_df,
  file = paste0(exceptions_file_path, exceptions_file_name),
  sheetName = "Sheet1"
)

print("Exceptions File Exported")

#--------------- CHANGE WORKING DIRECTORY ---------------

# To save built report workbook to network drive archive

# Save working directory to reset back to project directory later
old_dir <- getwd()

# Save to Network Drive Archive
setwd("\\\\360Corp-WShare\\zdrive\\RAC\\RAC Report Archive\\RAC CVI Consumer Check v2\\")

#--------------- EXPORT BUILT FILE ---------------

# Get row and column index
last_row <- nrow(df)+1
all_cols <- 1:ncol(df)

# Create report file in excel with conditional formatting rules
create_report_workbook <- function() {
  # Create workbook #
  wb <- createWorkbook()
  # Add sheets #
  addWorksheet(wb, 
               sheetName = "Data"
               )
  # Write data #
  writeData(wb, 
            sheet = "Data", 
            x = df,
            withFilter = TRUE
            )
  # Set column width #
  # Raction column
  setColWidths(wb, 
               sheet = "Data", 
               cols = 1:1, 
               widths = 10
               )
  # Conditional formatting styles #
  red_style <- createStyle(fontColour = "#9C0006", 
                          bgFill = "#FFC7CE"
                          )
  yellow_style <- createStyle(fontColour = "#9C6500", 
                             bgFill = "#FFEB9C"
                             )
  green_style <- createStyle(fontColour = "#006100", 
                            bgFill = "#C6EFCE"
                            )
  # Conditional formatting rules #
  # Main rules
  conditionalFormatting(wb, 
                        sheet = "Data", 
                        cols = all_cols, 
                        rows = 1:last_row, 
                        type = "expression", 
                        rule = '$A1="TAG"', 
                        style = red_style
                        )
  conditionalFormatting(wb, 
                        sheet = "Data", 
                        cols = all_cols, 
                        rows = 1:last_row, 
                        type = "expression", 
                        rule = '$A1="PREV TAG"', 
                        style = yellow_style
                        )
  conditionalFormatting(wb, 
                        sheet = "Data", 
                        cols = all_cols, 
                        rows = 1:last_row, 
                        type = "expression", 
                        rule = '$A1="Diff Patient"', 
                        style = green_style
                        )
  # Misc rules
  conditionalFormatting(wb, 
                        sheet = "Data", 
                        cols = all_cols, 
                        rows = 1:last_row, 
                        type = "expression", 
                        rule = '$A1="BH TAG"', 
                        style = red_style
                        )
  conditionalFormatting(wb, 
                        sheet = "Data", 
                        cols = all_cols, 
                        rows = 1:last_row, 
                        type = "expression",
                        rule = '$A1="IS"', 
                        style = green_style
                        )
  # Built report filename #
  built_report_filename <- paste0("Copy of RAC CVI Consumer Check v2 ", 
                                  format(Sys.Date(), "%m-%d-%Y"), 
                                  ".xlsx"
                                  )
  # Save workbook #
  saveWorkbook(wb, 
               file = built_report_filename
               )
}

create_report_workbook()

print("Built Report File Exported")

#--------------- RESET WORKING DIRECTORY ---------------

# Reset working directory back to project directory
setwd(old_dir)

#--------------- TRACKER INFO ---------------

# For tracking
tracker_info <- function(df, exceptions_df) {
  # Counts total hits 
  print(paste0(nrow(df)," - Total Hits"))
  # Counts total actioned (TAG)
  print(paste0(nrow(exceptions_df)," - Total Actioned"))
}

tracker_info(df, exceptions_df)

#--------------- CHECK FOR MISSING RACTION ---------------

# Check if any Raction reasons/tags are missing
na_Raction <- sum(is.na(df$`Raction`))

# Check missing values to manually check data
print(if_else(na_Raction > 0, "Missing Raction Reasons - Please Check Data", "No Raction Reasons Missing - Continue Forward"))

#--------------- OPEN APP SUPPORT LINK ---------------

# Send exceptions file to App Support -> opens in default browser
browseURL("https://360insights.atlassian.net/servicedesk/customer/portal/28/group/107/create/520")

print("Opening Request Portal Link")