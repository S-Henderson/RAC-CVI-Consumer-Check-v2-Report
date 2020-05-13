#--------------- IMPORT DATA ---------------

# File name is: "RAC CVI Consumer Check v2 %m-%d-%Y.xlsx"

src_file_path <- "C:/Users/shenderson/Downloads"

src_file_pattern <- "^RAC CVI Consumer Check v2(.*)xlsx$"

src_file_name = list.files(path = src_file_path, 
                           pattern = src_file_pattern, 
                           full.names = TRUE
                           )

print(paste0("File Selected is -> ", src_file_name))

# Read File
df <- read_excel(src_file_name, 
                 sheet = "Sheet1",
                 guess_max = Inf
                 )

print("Raw Data File Imported")