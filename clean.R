library(tidyverse)
library(readxl)
library(dplyr)
library(janitor)
library(readr)
library(stringr)
library(openxlsx)

# By: Scott Henderson
# Last Updated: Apr 1, 2020

#--------------- REMOVE DUPLICATES ---------------

# Removes duplicates by Transaction Number
remove_duplicates <- function(df) {
  df <- df %>%
    distinct(
      `Transaction Number`, 
      .keep_all = TRUE
    )
}

df <- remove_duplicates(df)

#--------------- CLEAN TRANSACTIONS ---------------

# Converts Transactions to numeric 
df$`Transaction Number` = as.numeric(as.character(df$`Transaction Number`))

#--------------- CLEAN PATIENT NAMES ---------------

# Converts Patient Names to Uppercase to properly match

df$`Patient First Name` = toupper(df$`Patient First Name`)

df$`Patient Last Name` = toupper(df$`Patient Last Name`)

df$`Previous Patient First Name` = toupper(df$`Previous Patient First Name`)

df$`Previous Patient Last Name` = toupper(df$`Previous Patient Last Name`)