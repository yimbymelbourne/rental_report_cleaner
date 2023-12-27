library(tidyverse)
library(readxl)
library(lubridate)

#Import Victoria's rental database 

rental_file <- list.files(path = "data",full.names = T)

sheet_names <- excel_sheets(path = rental_file)

import_sheet <- function(x){
# Step 2: Read the first two rows to get the column names
col_names <- read_excel(rental_file, sheet = x,n_max = 2) 


# Loop through each row and column of the dataframe
for (i in 1:nrow(col_names)) {
  for (j in 2:ncol(col_names)) { # Start at the second column
    if (is.na(col_names[i, j])) {
      col_names[i, j] <- col_names[i, j - 1] # Replace NA with the value from the previous column
    }
  }
}


new_col_names <- paste(col_names[1, -c(1,2)], col_names[2, -c(1,2)], sep = "_")

output <- read_excel(rental_file, sheet = x,skip =1 ) %>% 
  rename(lga = 2,
         `Dec 2002` = `Dec 2003...31`, 
         `Dec 2003` =  `Dec 2003...39`) %>% 
  select(-1) %>% 
  filter(!is.na(lga)) %>% 
  filter(row_number() != 1) %>% 
  pivot_longer(-lga,
               names_to = "date_messy",
               values_to = "value") %>%
  mutate(date = if_else(substr(date_messy,1,1)==".",lag(date_messy),date_messy ),
         type  = if_else(substr(date_messy,1,1)==".",
                         "count",
                         "median" )) %>%
  select(-date_messy) %>% 
  mutate(series = x)

return(output) }

data <- map_df(sheet_names,import_sheet)  %>% 
  filter(value != "-") %>% 
  mutate(date = dmy(paste0("1 ",date)),
         bedrooms = parse_number(series),
         type = case_when(str_detect(series,"ouse") ~ "House",
                          str_detect(series,"lat")  ~"Flat",
                          str_detect(series,"All")  ~ "All"))
