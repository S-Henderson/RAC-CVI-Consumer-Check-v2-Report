library(tidyverse)
library(readxl)
library(dplyr)
library(janitor)
library(readr)
library(stringr)
library(openxlsx)

# By: Scott Henderson
# Last Updated: Apr 1, 2020

#--------------- SCRIPT STARTING ---------------

start_time <- format(Sys.time(), "%X")

print(paste0("Script Starting at ", start_time))

#--------------- SCRIPTS ---------------

# Load Data
source(".R_Scripts/load.R")

# Clean Data
source(".R_Scripts/clean.R")

# Analyze Data
source(".R_Scripts/analysis.R")

# Export Data
source(".R_Scripts/reporting.R")

#--------------- SCRIPT COMPLETED ---------------

end_time <- format(Sys.time(), "%X")

print(paste0("Script Completed at ", end_time))