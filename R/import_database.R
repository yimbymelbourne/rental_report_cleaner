library(tidyverse)
library(readxl)
library(lubridate)
library(grattan)

#Import Victoria's rental database 

rental_file_list <- list.files(path = "data",full.names = T)

get_rental_file <- function(rental_file) {
sheet_names <- excel_sheets(path = rental_file)

import_sheet <- function(sheet_name){
# Step 2: Read the first two rows to get the column names
col_names <- read_excel(rental_file, sheet = sheet_name,n_max = 2) 


# Loop through each row and column of the dataframe
for (i in 1:nrow(col_names)) {
  for (j in 2:ncol(col_names)) { # Start at the second column
    if (is.na(col_names[i, j])) {
      col_names[i, j] <- col_names[i, j - 1] # Replace NA with the value from the previous column
    }
  }
}


new_col_names <- paste(col_names[1, -c(1,2)], col_names[2, -c(1,2)], sep = "_")

output <- read_excel(rental_file, sheet = sheet_name,skip =1 ) %>% 
  rename(area = 2) %>% 
  rename(any_of(c("Dec 2002" = "Dec 2003...31" ))) %>% 
  rename(any_of(c("Dec 2003" = "Dec 2003...39" ))) %>% 
  select(-1) %>% 
  filter(!is.na(area)) %>% 
  filter(row_number() != 1) %>% 
  pivot_longer(-area,
               names_to = "date_messy",
               values_to = "value") %>%
  mutate(date = if_else(substr(date_messy,1,1)==".",lag(date_messy),date_messy ),
         measure  = if_else(substr(date_messy,1,1)==".",
                         "median",
                         "count" )) %>%
  select(-date_messy) %>% 
  mutate(file_name = rental_file) %>%
  mutate(area_type = if_else(str_detect(file_name,"suburb"),"suburb","lga")) %>% 
  filter(value != "-") %>% 
  mutate(series = sheet_name,
         date = dmy(paste0("1 ",date)),
         bedrooms = parse_number(series),
         home_type = case_when(str_detect(series,"ouse") ~ "House",
                          str_detect(series,"lat")  ~"Flat",
                          str_detect(series,"All")  ~ "All"),
         value = parse_number(value),
         cpi_adjusted_value = if_else(measure == "median",
                                      value *cpi_inflator(from = date, 
                                                   to   = today(),
                                                   series = "original"),
                                      NA_real_
                                      )
         )

return(output)

}

output_sheet <- map_df(sheet_names,import_sheet)

return(output_sheet)

}



data <- map_df(rental_file_list,get_rental_file)

data %>% write_csv("Victoria's rental index.csv")

data %>% 
  filter(area == "Collingwood-Abbotsford",
         measure == "median") %>% 
  filter(home_type == "Flat",
         bedrooms == 2,
         date > ymd("2013-01-01")) %>% 
  ggplot(aes(x = date, y= value, colour = series))+
  geom_line() +
  labs(title = "Rents in Collingwood and Abbotsford",
       caption = "Source - Victoria's rental report",
       y = element_blank(),
       x = element_blank())+
  scale_y_continuous(labels = scales::dollar_format())+
  theme_minimal()
