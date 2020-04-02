library(tidyverse)
library(readxl)
library(dplyr)
library(janitor)
library(readr)
library(stringr)
library(openxlsx)

# By: Scott Henderson
# Last Updated: Apr 2, 2020

#--------------- IMPORT DATA ---------------

# Reads excel file - opens file browser window
df <- read_excel(
  file.choose(), 
  sheet = "Sheet1",
  guess_max = Inf
)

print("Raw Data File Imported")