#--------------- CLEAN TRANSACTIONS ---------------

# Converts Transactions to numeric 
df$`Transaction Number` = as.numeric(as.character(df$`Transaction Number`))

print("Cleaning Transactions Numbers")

#--------------- CLEAN PATIENT NAMES ---------------

# Converts Patient Names to Uppercase to properly compare

df$`Patient First Name` = toupper(df$`Patient First Name`)

df$`Patient Last Name` = toupper(df$`Patient Last Name`)

df$`Previous Patient First Name` = toupper(df$`Previous Patient First Name`)

df$`Previous Patient Last Name` = toupper(df$`Previous Patient Last Name`)

print("Cleaning Patient Names")

#--------------- REMOVE DUPLICATES ---------------

# Remove duplicate transactions
remove_duplicates <- function(df) {
  df <- df %>%
    distinct(
      `Transaction Number`, 
      .keep_all = TRUE
    )
}

df <- remove_duplicates(df)

print("Removing Duplicates")