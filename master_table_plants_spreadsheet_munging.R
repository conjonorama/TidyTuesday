library(tidyverse)
library(readxl)
library(tidyxl)

# Read in the sheet, I've opted to not use the first line as headers since I add them back in manually at the end.
x <- read_excel("master_table_plants_extinct_color.xlsx", col_names = FALSE)

# Pulling the formatting out of the sheet
formats <- xlsx_formats("master_table_plants_extinct_color.xlsx")

# Pulling the fill colours specifically out of the path
fill_colours_path <- formats$local$fill$patternFill$fgColor$rgb

#  Import all the cells (note that it also imports extra whitespace beyond column 24, this is deal with below),
# Create new columns of 'x_fill' with the fill colours, by looking up the local format id of each cell
fills <- xlsx_cells("master_table_plants_extinct_color.xlsx",
                sheet = "Sheet1") %>% 
  mutate(fill_colour = fill_colours_path[local_format_id]) %>% 
  select(row, col, fill_colour) %>% 
  spread(col, fill_colour) %>% 
  set_names(paste0(colnames(x), "_fill"))

# Combine the original sheet and the format sheet
y <- bind_cols(x, fills)

# Removed the two header rows, will be added back in below
y1 <- y[-(1:2),]

# Moved format colours from their columns to where they are on the sheet 
# (i.e. moved the hexcode info from columns 35:52 to columns 6:23)
y1[,(6:23)] <- y1[,(35:52)]

# Removed excess columns right of 'Red List Category' column
y1 <- y1[,-(25:length(y1))]

# Threats + Current_Actions (From Datapasta, god bless lol)
threats <- c("AA", "BRU", "RCD", "ISGD", "EPM", "CC", "HID", "P", "TS", "NSM", "GE", "NA")
current_actions <- c("LWP", "SM", "LP", "RM", "EA", "NA")


# Setting column names
y1 <- y1 %>% 
  set_names(c("Binomial_Name",
              "Country",
              "Continent",
              "Group",
              "Year_Last_Seen",
              paste0(rep("Threats_", length(threats)), threats),
              paste0(rep("Current_Actions_", length(current_actions)), current_actions),
              "Red_List_Category"))





