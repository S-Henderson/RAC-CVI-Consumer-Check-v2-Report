#--------------- RACTION ---------------

# Document actions
create_Raction <- function(df) {
  df <- df %>%
    mutate(
      `Raction` = case_when(
        `Is Blackhawk` == "TRUE"                              ~ "BH TAG",
        `Previous Claim Status`  == "Invalid Submission"      ~ "IS",
        # Accounts for mutiple exception reasons
        !is.na(`Exception Reason`)                            ~ "PREV TAG",
        # Checks First Name as query pulls same Last Name matches
        `Patient First Name` == `Previous Patient First Name` ~ "TAG",
        `Patient First Name` != `Previous Patient First Name` ~ "Diff Patient"
      )
    )
}

df <- create_Raction(df)

print("Creating Raction Column")

#--------------- PATIENT NAME MATCH ---------------

# Match column mostly for audit checks
patient_name_match <- function(df) {
  df <- df %>%
    mutate(
      `Patient First Name Match` = case_when(
        `Patient First Name` == `Previous Patient First Name` ~ "TRUE",
        `Patient First Name` != `Previous Patient First Name` ~ "FALSE"
      )
    )
}

df <- patient_name_match(df)

print("Checking Patient Name Matches")

#--------------- RE-ORDER COLUMNS ---------------

# Easier to read
reorder_df_columns <- function(df) {
  df <- df %>%
    select(`Raction`, 
           1:38, 
           `Patient First Name Match`, # Put after Names
           everything()
           )
}

df <- reorder_df_columns(df)

print("Re-ordering Columns")