library(tidyverse)
library(readxl)
library(dplyr)
library(janitor)
library(readr)
library(stringr)
library(openxlsx)

# By: Scott Henderson
# Last Updated: Apr 2, 2020

#--------------- RACTION ---------------

# Action column
create_Raction <- function(df) {
  df <- df %>%
    mutate(
      `Raction` = case_when(
        `Is Blackhawk` == "TRUE" 
        ~ "BH TAG",
        `Previous Claim Status`  == "Invalid Submission" 
        ~ "IS",
        # If exception_reason is not blank - accounts for mutiple exception reasons
        !is.na(`Exception Reason`) 
        ~ "PREV TAG",
        # TAG name matches -> only checks First Name as report pulls same Last Name
        `Patient First Name` == `Previous Patient First Name` 
        ~ "TAG",
        `Patient First Name` != `Previous Patient First Name` 
        ~ "Diff Patient"
      ))
}

df <- create_Raction(df)

print("Creating Raction Column")

#--------------- PATIENT NAME MATCH ---------------

# Check Patient Name match -> mostly for audit reason -> only checks First Name as report pulls same Last Name
patient_name_match <- function(df) {
  df <- df %>%
    mutate(
      `Patient First Name Match` = case_when(
        `Patient First Name` == `Previous Patient First Name` 
        ~ "TRUE",
        `Patient First Name` != `Previous Patient First Name` 
        ~ "FALSE"
      ))
}

df <- patient_name_match(df)

print("Checking Patient Name Matches")

#--------------- RE-ORDER COLUMNS ---------------

# Puts Raction at start -> adds `Patient First Name Match` after patient names
reorder_df_columns <- function(df) {
  df <- df %>%
    select(`Raction`, 1:38, `Patient First Name Match`, everything())
}

df <- reorder_df_columns(df)

print("Re-ordering Columns")