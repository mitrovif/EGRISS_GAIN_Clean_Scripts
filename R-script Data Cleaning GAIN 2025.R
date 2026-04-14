### Libraries
library(readxl)
library(dplyr)
library(countrycode)
library(readr)
library(stringr)
library(gtranslate)
library(janitor)
library(lubridate)
library(fs)
library(openxlsx)
library(tidyr)
library(purrr)
library(stringi)



rm(list = ls())
`%not in%` <- Negate(`%in%`)

# If bExport == 1, export everything
# If bExport == 0, don't export
bExport = 0


# ******************************************************
# Step 0: Define file locations ----
# ******************************************************
#Fill this to have the right path 
year <- 2025

egr_yr <- str_c('EGRISS GAIN Survey ', year)


final_version_file <- str_c(egr_yr, "/06 Data Cleaning/01 Input/GAIN_Survey_2025_Final_Version.xlsx")
data_clean_file <- str_c(egr_yr,"/06 Data Cleaning/01 Input/EGRISS_GAIN_2025_-_Data Clean.xlsx")
output_directory <- str_c(egr_yr, "/10 Data")
gain_survey_all_file <- file.path(output_directory, "GAIN Survey - All Data.xlsx")

# Define subfolder for analysis-ready files
analysis_ready_directory <- file.path(output_directory, "01 Analysis Ready Files")

# Ensure the output directory for analysis-ready files exists
if (!dir.exists(analysis_ready_directory)) {
  dir.create(analysis_ready_directory, recursive = TRUE)
}

# ******************************************************
# Step 1: Import all sheets from both files ----
# ******************************************************

final_version_sheets <- excel_sheets(final_version_file)
data_clean_sheets <- excel_sheets(data_clean_file)

# Read all sheets from Final Version
final_version_data <- lapply(final_version_sheets, function(sheet) {
  read_excel(final_version_file, sheet = sheet)
})
names(final_version_data) <- final_version_sheets

# Read all sheets from Data Clean
data_clean_data <- lapply(data_clean_sheets, function(sheet) {
  read_excel(data_clean_file, sheet = sheet)
})
names(data_clean_data) <- data_clean_sheets


# ******************************************************
# Step 2: Rename `_PRO02A` to `PRO02` and `_index` to `index` in `group_roster` ----
# ******************************************************

if ("group_roster" %in% names(final_version_data)) {
  group_roster <- final_version_data[["group_roster"]]
  
  # Rename `_PRO02A` to `PRO02` if it exists
  if ("_PRO02A" %in% colnames(group_roster)) {
    group_roster <- group_roster %>%
      rename(PRO02 = `_PRO02A`)
    message("Renamed `_PRO02A` to `PRO02` in `group_roster`.")
  }
  
  # Rename `_index` to `index` if it exists
  if ("_index" %in% colnames(group_roster)) {
    group_roster <- group_roster %>%
      rename(index = `_index`)
    message("Renamed `_index` to `index` in `group_roster`.")
  }
  
  # Save back to the list
  final_version_data[["group_roster"]] <- group_roster
}


#Step not needed as already 'index'
# Rename `_index` to `index` in `GAIN Survey 2025`
if ("GAIN Survey 2025" %in% names(final_version_data)) {
  egriss_gain_2025 <- final_version_data[["GAIN Survey 2025"]]
  
  # Rename `_index` to `index` if it exists
  if ("_index" %in% colnames(egriss_gain_2025)) {
    egriss_gain_2025 <- egriss_gain_2025 %>%
      rename(index = `_index`)
    message("Renamed `_index` to `index` in `GAIN Survey 2025`.")
  }
  
  # Save back to the list
  final_version_data[["GAIN Survey 2025"]] <- egriss_gain_2025
}

# ******************************************************
# Step 3: Remove entries in "GAIN Survey 2025" based on del_main ----
# ******************************************************

if ("GAIN Survey 2025" %in% names(final_version_data) && "del_main" %in% names(data_clean_data)) {
  egriss_gain_2025 <- final_version_data[["GAIN Survey 2025"]]
  del_egriss_gain <- data_clean_data[["del_main"]]
  
  # Remove rows from "GAIN Survey 2025" based on del_main
  egriss_gain_2025_cleaned <- egriss_gain_2025 %>%
    filter(!(index %in% del_egriss_gain$index))
  
  # Save cleaned version of "GAIN Survey 2025" back to the list
  final_version_data[["GAIN Survey 2025"]] <- egriss_gain_2025_cleaned
}


contr <- nrow(egriss_gain_2025) - nrow(egriss_gain_2025_cleaned)
contr
nrow(data_clean_data[["del_main"]])
#OK removed 108639846

egriss_gain_2025 <- egriss_gain_2025_cleaned

rm(egriss_gain_2025_cleaned)


# ******************************************************
# Step 4: Remove entries in "group_roster" based on del_group_roster ----
# ******************************************************

contr <- final_version_data[["group_roster"]] %>% 
  select(index, PRO02, `PRO03C. Please Select the country:`)

if ("group_roster" %in% names(final_version_data) && "del_group_roster" %in% names(data_clean_data)) {
  group_roster <- final_version_data[["group_roster"]]
  del_group_roster <- data_clean_data[["del_group_roster"]]
  
  # Remove rows from "group_roster" based on del_group_roster
  group_roster_cleaned <- group_roster %>%
    filter(!(index %in% del_group_roster$index))
  
  # Save cleaned version of "group_roster" back to the list
  final_version_data[["group_roster"]] <- group_roster_cleaned
}

contr <- nrow(group_roster) - nrow(group_roster_cleaned)
contr
nrow(data_clean_data[["del_group_roster"]])
#OK removed

group_roster <- group_roster_cleaned

contr <- group_roster %>% 
  select(index, PRO02, `PRO03C. Please Select the country:`)


rm(group_roster_cleaned)


# ******************************************************
# Step 5: Add text values to `PRO02` in "group_roster" ----
# ******************************************************

contr <- group_roster %>%
  group_by(PRO02) %>%
  tally()

#NO NA
#However has to be translated in English

contr <- group_roster %>% 
  select(index, PRO02, `PRO03C. Please Select the country:`)


if ("group_roster" %in% names(final_version_data) && "title_group" %in% names(data_clean_data)) {
  group_roster <- final_version_data[["group_roster"]]
  title_pro02 <- data_clean_data[["title_group"]]
  
  # Add or update `PRO02` in "group_roster" based on title_PRO02
  group_roster <- group_roster %>%
    left_join(title_pro02 %>% select(group_roster_index, `New Title`) %>% rename(index = group_roster_index,
                                                                                 title = `New Title`), by = "index") %>%
    mutate(
      PRO02 = ifelse(is.na(title), PRO02, title) # Take only the ones with a definition from data cleaning
      # PRO02 = ifelse(is.na(PRO02), title, PRO02) # Update only missing values - OLD VERSION
    ) %>%
    select(-title) # Remove temporary title column
  
  # Save updated version of "group_roster" back to the list
  final_version_data[["group_roster"]] <- group_roster
}

contr <- group_roster %>% 
  select(index, PRO02, `PRO03C. Please Select the country:`)


#Translation
contr <- group_roster %>%
  group_by(PRO02) %>%
  tally()


contr <- group_roster %>% 
  select(index, PRO02, `PRO03C. Please Select the country:`)


if ("group_roster" %in% names(final_version_data)) {
  group_roster <- final_version_data[["group_roster"]]
  
  # Translate
  group_roster <- group_roster %>%
    mutate(PRO02 = translate(PRO02, from = "auto", to = "en", trim_str = TRUE))
  
  # Save updated version of "group_roster" back to the list
  final_version_data[["group_roster"]] <- group_roster
}




contr2 <- group_roster %>%
  group_by(PRO02) %>%
  tally()
#Translation OK



# ******************************************************
# Step 6: Add `year` column to both cleaned datasets ----
# ******************************************************


# Define the datasets that need a year column
datasets_to_update <- c("GAIN Survey 2025", "group_roster")

# Loop through and update only those that exist
for (nm in datasets_to_update) {
  if (nm %in% names(final_version_data)) {
    final_version_data[[nm]] <- final_version_data[[nm]] %>%
      mutate(year = year)
  }
}

summary(final_version_data[["GAIN Survey 2025"]])
summary(final_version_data[["group_roster"]])
#OK



# ******************************************************
# Step 7: Rename variables in "GAIN Survey 2025" using var_main (sequence-based) ----
# ******************************************************
# Determine which columns to rename, excluding "year"

var_main <- data_clean_data[['var_main']] %>% clean_names()

if (all(c("newg_2024", "newg_2025") %in% names(var_main))){
  contr <- var_main %>% 
    select(newg_2024, newg_2025)%>%
    filter(newg_2024 != newg_2025)
  #Are the same, doesn't really change which one is selected
}


egriss_gain_2025 <- final_version_data[["GAIN Survey 2025"]]
cols_to_rename <- setdiff(colnames(egriss_gain_2025), "year")

# Number of columns to rename from var_main (make sure not to exceed the length)
num_vars_to_rename <- min(length(cols_to_rename), nrow(var_main))

# Get the old names (the ones you want to change) in order
old_names <- cols_to_rename[1:num_vars_to_rename]
new_names <- var_main$newg_2025[1:num_vars_to_rename]
new_type  <- var_main$data_type


# Replace the names while preserving the "year" column
for(i in seq_along(old_names)) {
  idx <- which(colnames(egriss_gain_2025) == old_names[i])
  if(length(idx) > 0) {
    colnames(egriss_gain_2025)[idx] <- new_names[i]
  }
}

ls(egriss_gain_2025)
#OK

summary(egriss_gain_2025)

# Change types of columns
for (i in seq_along(new_names)) {
  col_name <- new_names[i]
  type     <- new_type[i]
  
  if (col_name %in% names(egriss_gain_2025)) {
    egriss_gain_2025[[col_name]] <- switch(
      type,
      "numeric"   = as.numeric(egriss_gain_2025[[col_name]]),
      "integer"   = as.integer(egriss_gain_2025[[col_name]]),
      "character" = as.character(egriss_gain_2025[[col_name]]),
      "factor"    = as.factor(egriss_gain_2025[[col_name]]),
      "Date"      = as.Date(egriss_gain_2025[[col_name]]),
      
      egriss_gain_2025[[col_name]]
    )
  }
}

summary(egriss_gain_2025)

final_version_data[["GAIN Survey 2025"]] <- egriss_gain_2025

print(colnames(final_version_data[["GAIN Survey 2025"]]))


# ******************************************************
# Step 8: Rename variables in "group_roster" using var_group (sequence-based) ----
# ******************************************************

contr <- group_roster %>% 
  select(index, PRO02, `PRO03C. Please Select the country:`)

if ("var_group" %in% names(data_clean_data) && "group_roster" %in% names(final_version_data)) {
  var_group <- data_clean_data[["var_group"]]
  group_roster <- final_version_data[["group_roster"]]
  
  # Ensure the sequence matches
  num_vars_to_rename <- min(ncol(group_roster), nrow(var_group)) # Limit to the smaller size
  old_names <- colnames(group_roster)[1:num_vars_to_rename]
  new_names <- var_group$newgr_2025[1:num_vars_to_rename]
  
  # Rename variables in group_roster
  colnames(group_roster)[1:num_vars_to_rename] <- new_names
  
  # Save back to the list
  final_version_data[["group_roster"]] <- group_roster
  message("Variables renamed in 'group_roster' based on sequence.")
}

ls(group_roster)
summary(group_roster)

contr <- group_roster %>% 
  select(index, PRO02A, PRO03C)


# Save cleaned and renamed datasets back to the specified directory
if (bExport == 1){
  write.xlsx(final_version_data[["GAIN Survey 2025"]], file.path(output_directory, "renamed_egriss_gain_2025.xlsx"), rowNames = FALSE)
  write.xlsx(final_version_data[["group_roster"]], file.path(output_directory, "renamed_group_roster.xlsx"), rowNames = FALSE)
}


# View final cleaned and renamed datasets
summary(final_version_data[["GAIN Survey 2025"]])
summary(final_version_data[["group_roster"]])
summary(final_version_data[["GAIN Survey 2025"]]$year)

message("Cleaned and renamed datasets have been saved to: ", output_directory)


# Define file locations
output_directory <- str_c(egr_yr, "/10 Data")
# gain_survey_all_file <- file.path(output_directory, "GAIN Survey - All Data.xlsx")

# Define subfolder for analysis-ready files
analysis_ready_directory <- file.path(output_directory, "01 Analysis Ready Files")

# Ensure the output directory for analysis-ready files exists
if (!dir.exists(analysis_ready_directory)) {
  dir.create(analysis_ready_directory, recursive = TRUE)
}


# ******************************************************
# Step 9 Adapt PRO08 to the prior versions ----
# ******************************************************

group_roster_change <- final_version_data[["group_roster"]]

ls(group_roster_change)

contr <- group_roster_change %>% 
  group_by(PRO08, PRO08.A, PRO08.B, PRO08.C, PRO08_label, PRO08_other, `PRO08B/A` , 
           `PRO08B/B` , `PRO08B/C`, `PRO08B/D`, `PRO08B/E`, `PRO08B/F`, `PRO08B/G`, `PRO08B/H`, `PRO08B/X`, PRO08B_other,
           `_PRO08B`, `_PRO08B_a`, `_PRO08B_a2`, `_PRO08B_b`, `_PRO08B_b2`, `_PRO08B_c`, `_PRO08B_d`, `_PRO08B_e`, `_PRO08B_f`,
           `_PRO08B_g`, `_PRO08B_h`, `_PRO08B_x`) %>% 
  tally()


# If the main source used is mentioned, then will be automatically 1 in the corresponding column
group_roster_change1 <- group_roster_change %>% 
  mutate(`PRO08B/A` = ifelse(PRO08 == 'SURVEY', 1, `PRO08B/A`),
         `PRO08B/B` = ifelse(PRO08 == 'ADMINISTRATIVE DATA', 1, `PRO08B/B`),
         `PRO08B/C` = ifelse(PRO08 == 'CENSUS', 1, `PRO08B/C`),
         `PRO08B/D` = ifelse(PRO08 == 'DATA INTEGRATION', 1, `PRO08B/D`),
         `PRO08B/E` = ifelse(PRO08 == 'NON-TRADITIONAL', 1, `PRO08B/E`),
         `PRO08B/F` = ifelse(PRO08 == 'STRATEGY', 1, `PRO08B/F`),
         `PRO08B/G` = ifelse(PRO08 == 'GUIDANCE/TOOLKIT', 1, `PRO08B/G`),
         `PRO08B/H` = ifelse(PRO08 == 'WORKSHOP/TRAINING', 1, `PRO08B/H`)) %>% 
  rename(PRO08.C    = PRO08.C,
         PRO08B.A   = `PRO08B/A`,
         PRO08B.B   = `PRO08B/B`,
         PRO08B.C   = `PRO08B/C`,
         PRO08B.D   = `PRO08B/D`,
         PRO08B.E   = `PRO08B/E`,
         PRO08B.F   = `PRO08B/F`,
         PRO08B.G   = `PRO08B/G`,
         PRO08B.H   = `PRO08B/H`,
         PRO08B.X   = `PRO08B/X`) %>% 
  select(-c(`_PRO08B`, `_PRO08B_a`, `_PRO08B_a2`, `_PRO08B_b`, `_PRO08B_b2`, `_PRO08B_c`, `_PRO08B_d`, `_PRO08B_e`, `_PRO08B_f`,
            `_PRO08B_g`, `_PRO08B_h`, `_PRO08B_x`, `_PRO07`, `_PRO07_a`, 
            `_PRO07_b`, `_PRO07_c`, `_PRO07_x`)) #Remove useless columns



final_version_data[["group_roster"]] <- group_roster_change1
group_roster <- group_roster_change1



# ******************************************************
# Step 10: Load `GAIN Survey - All Data` file ----
# Historical data (2021 - 2024)
# ******************************************************
# Load relevant sheets

#Path might have to be adapted
main_roster      <- read_excel(str_c(egr_yr, "/06 Data Cleaning/01 Input/01 Historical Data/analysis_ready_main_roster.xlsx"), .name_repair = "unique") # Handle duplicate names
group_roster_all <- read_excel(str_c(egr_yr, "/06 Data Cleaning/01 Input/01 Historical Data/analysis_ready_group_roster.xlsx"), .name_repair = "unique") # Handle duplicate names


contr <- group_roster_all %>%  group_by(PRO08, PRO08.A, PRO08.B, PRO08.C, 
                                        PRO08.D, PRO08.E, PRO08.F, PRO08.G, PRO08.H, PRO08.X, PRO08a) %>%  tally()


contr <- main_roster %>% 
  group_by(year) %>% 
  summarise(
    across(
      starts_with("ACT04"),
      ~ {
        x <- suppressWarnings(as.numeric(.x))
        if (all(is.na(x))) NA_real_ else sum(x, na.rm = TRUE)
      }
    ),
    .groups = "drop"
  )


#Change ACT04 to make it uniform
main_roster <- main_roster %>%
  # Fill empty ACT04 data with character that are in ACT04.A
  mutate(ACT04   = ACT04.A,
         ACT04   = ifelse(is.na(ACT04) == T & ACT04.A %not in% c(0, 1), ACT04.A, ACT04)) %>% 
  select(-ACT04.A) %>% 
  rename(ACT04.A = ACT04.B, 
         ACT04.B = ACT04.C, 
         ACT04.C = ACT04.D, 
         ACT04.D = ACT04.E,
         ACT04.E = ACT04.F, 
         ACT04.F = ACT04.G, 
         ACT04.G = ACT04.I)

# Recode manually to numeric
# main_roster <- main_roster %>%
#   # Fill empty ACT04 data with character that are in ACT04.A
#   mutate(ACT04.A = ifelse(grepl("IRRS", ACT04), 1,
#                           ifelse(year == 2023, ACT04.A, 0)),
#          # ifelse(ACT04.A %not in% c(0, 1), NA, ACT04.A)),
# 
#          ACT04.B = ifelse(grepl("IRIS", ACT04) & year %in% c(2024), 1,
#                           ifelse(year == 2023, ACT04.B, 0)),
# 
#          ACT04.C = ifelse(grepl("IROSS", ACT04) & year %in% c(2024), 1, 
#                           ifelse(year == 2023, ACT04.C, 0)), 
# 
#          ACT04.D = ifelse(grepl("EXTERNAL", ACT04) & year %in% c(2024), 1, 
#                           ifelse(year == 2023, ACT04.D, 0)),
# 
#          ACT04.E = ifelse(grepl("ANNUAL", ACT04) & year %in% c(2024), 1, 
#                           ifelse(year == 2023, ACT04.E, 0)),
# 
#          ACT04.F = ifelse(grepl("PROMOTIONAL", ACT04) & year %in% c(2024), 1,
#                           ifelse(year == 2023, ACT04.F, 0)),
# 
#          ACT04.G = ifelse(grepl("REVISED", ACT04) & year %in% c(2024), 1, 
#                           ifelse(year == 2023, ACT04.G, 0)),
# 
#          ACT04.H = ifelse(grepl("METHODOLOGICAL", ACT04) & year %in% c(2024), 1, 
#                           ifelse(year == 2023, ACT04.H, 0)),
# 
#          ACT04.X = ifelse(grepl("OTHER", ACT04) & year %in% c(2024), 1, 
#                           ifelse(year == 2023, ACT04.X, 0)),
# 
#          ACT04.Z = ifelse(grepl("KNOW", ACT04) & year %in% c(2024), 1, 
#                           ifelse(year == 2023, ACT04.Z, 0)))

ls(main_roster)

contr <- main_roster %>%
  group_by(year, across(starts_with("ACT04"))) %>%
  tally()

contr <- main_roster %>% 
  group_by(year) %>% 
  summarise(
    across(
      starts_with("ACT04"),
      ~ {
        x <- suppressWarnings(as.numeric(.x))
        if (all(is.na(x))) NA_real_ else sum(x, na.rm = TRUE)
      }
    ),
    .groups = "drop"
  )

# Ensure consistent naming for `index`
if ("_index" %in% colnames(main_roster)) {
  main_roster <- main_roster %>%
    rename(index = `_index`)
}


#Correct dates formatting
var <- c('start', 'end','today', 'X_submission_time')

for (i in var) {
  main_roster <- main_roster %>%
    mutate(
      !!i := as.numeric(.data[[i]]),                  
      !!i := as.POSIXct(.data[[i]] * 86400,          
                        origin = "1899-12-30", tz = "UTC")
    )
}

main_roster <- main_roster %>%
  mutate(across(where(~ inherits(., "character")), ~ na_if(., "NA")))


summary(main_roster)
ls(main_roster)


# ******************************************************
# Step 11: Merge `main_roster` with `renamed_egriss_gain_2025` ----
# Aggregate new and old data
# ******************************************************

renamed_egriss_gain_2025 <- egriss_gain_2025


align_column_types <- function(df1, df2) {
  # Get the column names for both dataframes
  colnames_df1 <- colnames(df1)
  colnames_df2 <- colnames(df2)
  
  # Find the common column names between the two dataframes
  common_cols <- intersect(colnames_df1, colnames_df2)
  
  # Loop through the common columns and align their types
  for (col in common_cols) {
    # Check the types of the first element in both columns
    class_df1 <- class(df1[[col]])[1]  # Take the first element's class
    class_df2 <- class(df2[[col]])[1]  # Take the first element's class
    
    # If one of the columns is a Date or POSIXct, convert both columns to character
    if ("POSIXct" %in% c(class_df1, class_df2)) {
      # Convert both columns to character if one is POSIXct
      df1[[col]] <- as.character(df1[[col]])
      df2[[col]] <- as.character(df2[[col]])
    } else if ("Date" %in% c(class_df1, class_df2)) {
      # Convert both columns to character if one is Date
      df1[[col]] <- as.character(df1[[col]])
      df2[[col]] <- as.character(df2[[col]])
    } else {
      # For non-date types, match the type (prefer character if one column is character)
      if (class_df1 != class_df2) {
        if (class_df1 == "character" || class_df2 == "character") {
          # Convert both columns to character if either is character
          df1[[col]] <- as.character(df1[[col]])
          df2[[col]] <- as.character(df2[[col]])
        } else {
          # If one column is numeric, convert both columns to numeric
          df1[[col]] <- as.numeric(df1[[col]])
          df2[[col]] <- as.numeric(df2[[col]])
        }
      }
    }
  }
  
  # Return the aligned dataframes as a list
  return(list(df1 = df1, df2 = df2))
}

# Remove the 'X' prefix from column names that start with 'X'
colnames(renamed_egriss_gain_2025) <- gsub("^X", "", colnames(renamed_egriss_gain_2025))
colnames(main_roster)              <- gsub("^X", "", colnames(main_roster))

# Explicitly convert _submission_time to character before alignment
renamed_egriss_gain_2025$`_submission_time` <- as.character(renamed_egriss_gain_2025$`_submission_time`)
main_roster$`_submission_time` <- as.character(main_roster$`_submission_time`)


# Align column types
aligned_data <- align_column_types(renamed_egriss_gain_2025, main_roster)
renamed_egriss_gain_2025 <- aligned_data$df1 
main_roster <- aligned_data$df2

# summary data type
sapply(renamed_egriss_gain_2025, class)
sapply(main_roster, class)

colnames(renamed_egriss_gain_2025) <- gsub("/", ".", colnames(renamed_egriss_gain_2025))


setdiff(colnames(renamed_egriss_gain_2025), colnames(main_roster))
setdiff(colnames(main_roster), colnames(renamed_egriss_gain_2025))
ls(main_roster)




# Aggregate datasets
merged_main <- bind_rows(renamed_egriss_gain_2025, main_roster)

nrow(renamed_egriss_gain_2025) #92
nrow(main_roster) #248
nrow(merged_main) #340
#OK

ncol(renamed_egriss_gain_2025) #102
ncol(main_roster) #104
ncol(merged_main) #106
#OK

ls(merged_main)


# Ensure `index` exists before arranging
summary(merged_main$index)
#OK

if ("index" %in% colnames(merged_main)) {
  merged_main <- merged_main %>%
    arrange(index)
} else {
  message("`index` column not found in merged_main. Skipping `arrange()`.")
}

# Save merged dataset
output_main_roster <- file.path(analysis_ready_directory, "analysis_ready_main_roster.xlsx")

if (bExport == 1){
  write.xlsx(merged_main, output_main_roster, rowNames = FALSE)
}

message("Merged and saved main roster as analysis-ready dataset at: ", output_main_roster)



# ******************************************************
# Step 12: Merge `group_roster_all` with `renamed_group_roster` ----
# Aggregate new and old data
# Also ensures `PRO04` and `PRO05` are converted to dates and rounded up to year.
# ******************************************************


ls(group_roster_all)
summary(group_roster_all$PRO04)


# Rename columns to match new ones
if ("PRO08.A" %in% colnames(group_roster_all)) {
  group_roster_all <- group_roster_all %>%
    rename(PRO08B.A = PRO08.A, 
           PRO08B.B = PRO08.B,
           PRO08B.C = PRO08.C, 
           PRO08B.D = PRO08.D,
           PRO08B.E = PRO08.E,
           PRO08B.F = PRO08.F,
           PRO08B.G = PRO08.G, 
           PRO08B.H = PRO08.H,
           PRO08B.X = PRO08.X,
           PRO08B_other = PRO08a)
}


contr <- group_roster %>% 
  select(index, PRO02A, PRO03C)


#Check matching colnames
colnames(group_roster) <- gsub("/", ".", colnames(group_roster))

ls(group_roster)
ls(group_roster_all)


setdiff(colnames(group_roster), colnames(group_roster_all))
setdiff(colnames(group_roster_all), colnames(group_roster))
ls(group_roster)


# Dates are not working
contr <- group_roster_all %>% 
  group_by(PRO04) %>% 
  tally()


var <- c('PRO04', 'PRO05')

group_roster_all <- group_roster_all %>%
  mutate(across(all_of(var), ~ {
    
    x <- str_trim(as.character(.))
    x[x == "" | x == "9999"] <- NA
    
    res <- as.Date(rep(NA_character_, length(x)))
    
    # 1. Dates already in ISO (YYYY-MM-DD)
    idx_char_date <- which(!is.na(x) & str_detect(x, "^\\d{4}-\\d{2}-\\d{2}$"))
    if (length(idx_char_date) > 0) {
      res[idx_char_date] <- as.Date(x[idx_char_date])
    }
    
    # 2. Numerical Values
    num <- suppressWarnings(as.numeric(x))
    
    # Years
    idx_year <- which(!is.na(num) & num >= 1900 & num <= 2100)
    if (length(idx_year) > 0) {
      res[idx_year] <- as.Date(paste0(num[idx_year], "-01-01"))
    }
    
    # Excel Dates
    idx_excel <- which(!is.na(num) & num > 2100)
    if (length(idx_excel) > 0) {
      res[idx_excel] <- as.Date(num[idx_excel], origin = "1899-12-30")
    }
    
    res
  }))


summary(group_roster_all$PRO04)
summary(group_roster_all$PRO05)
#OK

contr <- group_roster_all %>% 
  group_by(submission__submission_time) %>% 
  tally()


if (exists("group_roster_all") && exists("final_version_data") && "group_roster" %in% names(final_version_data)) {
  # Ensure column names are valid and not NA
  colnames(group_roster_all) <- make.names(colnames(group_roster_all), unique = TRUE)
  colnames(final_version_data[["group_roster"]]) <- make.names(colnames(final_version_data[["group_roster"]]), unique = TRUE)
  
  # Convert all columns to character to prevent type mismatches
  group_roster_all <- group_roster_all %>% mutate(across(everything(), as.character))
  renamed_group_roster <- final_version_data[["group_roster"]] %>% mutate(across(everything(), as.character))
  
  # Align columns
  all_columns_group <- union(colnames(renamed_group_roster), colnames(group_roster_all))
  
  for (col in setdiff(all_columns_group, colnames(group_roster_all))) {
    group_roster_all[[col]] <- NA
  }
  for (col in setdiff(all_columns_group, colnames(renamed_group_roster))) {
    renamed_group_roster[[col]] <- NA
  }
  
  # Merge datasets
  merged_group <- bind_rows(renamed_group_roster, group_roster_all)
  
  # Ensure `index` exists before arranging
  if ("X_index" %in% colnames(merged_group)) {
    merged_group <- merged_group %>% arrange(X_index)
  } else {
    message("`index` column not found in merged_group. Skipping `arrange()`.")
  }
  
  # Save merged dataset
  output_group_roster <- file.path(analysis_ready_directory, "analysis_ready_group_roster.xlsx")
  tryCatch({
    if (bExport == 1){
      write.xlsx(merged_group, output_group_roster, rowNames = FALSE)
    }
    message("Merged and saved group roster as analysis-ready dataset at: ", output_group_roster)
  }, error = function(e) {
    message("Error in writing group roster file: ", e$message)
  })
} else {
  message("One or more datasets are missing for merging 'group_roster_all' with 'renamed_group_roster'.")
}

summary(merged_group)
ls(merged_group)

#If problem in columns matching come here ----



# ******************************************************
# Step 13: Confirm saved files ----
# ******************************************************

saved_files <- list.files(analysis_ready_directory, full.names = TRUE)
if (length(saved_files) > 0) {
  message("Analysis-ready files have been saved:")
  print(saved_files)
} else {
  message("No analysis-ready files were saved. Please check the script and data inputs.")
}

# Load the analysis-ready group roster file
analysis_ready_group_roster_file <- str_c(egr_yr,"/10 Data/01 Analysis Ready Files/analysis_ready_group_roster.xlsx")

if (file.exists(analysis_ready_group_roster_file)) {
  # Load the dataset
  # group_roster <- read.xlsx(analysis_ready_group_roster_file) #don't need to import / export, keep it inside
  group_roster <- merged_group
  
  # Identify correct column names for `year`
  year_cols <- grep("^year", colnames(group_roster), value = TRUE)
  
  if (length(year_cols) >= 1) {
    # Use the first detected year column
    group_roster <- group_roster %>%
      mutate(ryear = coalesce(!!!syms(year_cols)),
             ryear = as.numeric(ryear),
             year = as.numeric(year)) # Combine all potential year columns
    
    # Save the updated dataset under the same name
    if (bExport == 1){
      write.xlsx(group_roster, analysis_ready_group_roster_file, rowNames = FALSE)
    }
    message("Created `ryear` column by combining available `year` columns. Updated file saved to `analysis_ready_group_roster.xlsx`.")
  } else {
    stop("No valid `year` columns found in the dataset.")
  }
} else {
  stop("The file 'analysis_ready_group_roster.xlsx' does not exist in the specified directory.")
}



summary(group_roster$ryear)
#OK


# ******************************************************
# Step 14: Create `pindex2` - Combine `ryear` and `pindex1` into an 8-digit index ----
# ******************************************************

contr <- group_roster %>%
  mutate(parent_index = as.numeric(parent_index),
         pindex1 = as.numeric(pindex1))

summary(contr$parent_index)
summary(contr$pindex1)

summary(group_roster)

contr <- group_roster %>% 
  select(index,pindex2, PRO02A, PRO03C)


# Ensure `pindex1` exists by extracting from `X_parent_index`
if ("parent_index" %in% colnames(group_roster)) {
  group_roster <- group_roster %>%
    mutate(parent_index = as.numeric(parent_index),
           pindex1 = as.numeric(parent_index))
}

if ("pindex1" %in% colnames(group_roster) & "ryear" %in% colnames(group_roster)) {
  group_roster <- group_roster %>%
    mutate(
      pindex1 = as.numeric(pindex1),
      ryear = as.numeric(ryear),
      pindex2 = sprintf("%d%04d", ryear, pindex1) # Combine `ryear` and padded `pindex1`
    )
  
  # Save the updated dataset under the same name
  tryCatch({
    if (bExport == 1){
      write.xlsx(group_roster, analysis_ready_group_roster_file, rowNames = FALSE)
    }
    message("Created `pindex1` and `pindex2` variables. Updated file saved to `analysis_ready_group_roster.xlsx`.")
  }, error = function(e) {
    message("Failed to save the file. Check if the file is open or the path is writable.")
    stop(e)
  })
} else {
  stop("Columns `ryear` or `pindex1` are missing in the dataset.")
}



# Load the analysis-ready main roster file
analysis_ready_main_roster_file <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_main_roster.xlsx")

if (file.exists(analysis_ready_main_roster_file)) {
  # Load the dataset
  # main_roster <- read.xlsx(analysis_ready_main_roster_file)
  main_roster <- merged_main
  
  # Ensure `year` and `index` columns exist
  if ("year" %in% colnames(main_roster) & "index" %in% colnames(main_roster)) {
    # Create `pindex2` by combining `year` and `index`
    main_roster <- main_roster %>%
      mutate(
        index = as.numeric(index), # Ensure `index` is numeric
        year = as.numeric(year), # Ensure `year` is numeric
        pindex2 = sprintf("%d%04d", year, index) # Combine `year` and padded `index`
      )
    
    # Save the updated dataset under the same name
    tryCatch({
      if (bExport == 1){
        write.xlsx(main_roster, analysis_ready_main_roster_file, rowNames = FALSE)
      }
      message("Created `pindex2` variable in the main roster file. Updated file saved to `analysis_ready_main_roster.xlsx`.")
    }, error = function(e) {
      message("Failed to save the file. Check if the file is open or the path is writable.")
      stop(e)
    })
  } else {
    stop("Columns `year` or `index` are missing in the dataset.")
  }
} else {
  stop("The file 'analysis_ready_main_roster.xlsx' does not exist in the specified directory.")
}

# Load the analysis-ready main roster file
analysis_ready_main_roster_file <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_main_roster.xlsx")

if (file.exists(analysis_ready_main_roster_file)) {
  # Load the dataset
  # main_roster <- read.xlsx(analysis_ready_main_roster_file)
  
  # Ensure that the LOC03 column exists (since Country has been removed)
  if ("LOC03" %in% colnames(main_roster)) {
    # Create the `mcountry` variable directly from LOC03
    main_roster <- main_roster %>%
      mutate(
        mcountry = LOC03
      )
    # Save the updated dataset under the same name
    tryCatch({
      if(bExport == 1){
        write.xlsx(main_roster, analysis_ready_main_roster_file, rowNames = FALSE)
      }
      message("Created `mcountry` variable from LOC03. Updated file saved to `analysis_ready_main_roster.xlsx`.")
    }, error = function(e) {
      message("Failed to save the file. Check if the file is open or the path is writable.")
      stop(e)
    })
  } else {
    stop("Column `LOC03` is missing in the dataset.")
  }
} else {
  stop("The file 'analysis_ready_main_roster.xlsx' does not exist in the specified directory.")
}



# ******************************************************
# Step 15: Adapt country names ----
# Filtering based on INT02
# ******************************************************


# Check country names
contr <- main_roster %>%
  group_by(LOC03) %>%
  tally()
#Mix between names (from new version) and Iso codes (from old one)

# From geo service UNHCR
iso_map <- data_clean_data[["iso_mapping"]] %>% 
  select(iso3, mcountry)

nrow(iso_map) #253
nrow(distinct(iso_map, iso3)) #253
#OK no duplicate

contr <- group_roster %>% 
  select(pindex2, PRO03, PRO03C, mcountry)


#Join to df
main_roster <- merge(main_roster, iso_map, by.x = 'mcountry', by.y = 'iso3', all.x = T)

#Already have an old mcountry from historical data
contr <- main_roster %>% 
  select(index, mcountry, mcountry.y)

# Goal is to add the new new from 2025
main_roster <- main_roster %>% 
  mutate(mcountry2 = ifelse(is.na(mcountry.y) == T, mcountry, mcountry.y)) %>% 
  select(-c(mcountry, mcountry.y)) %>% 
  rename(mcountry = mcountry2)



# Remove rows where `INT02` is "No" but preserve `NA`
contr <- main_roster %>%
  group_by(INT02) %>%
  tally()


nrow(main_roster) #347

main_roster <- main_roster %>%
  mutate(INT02 = ifelse(INT02 == 1, NA, INT02)) %>%
  filter(is.na(INT02) | INT02 != "No") # Keeps rows with `NA` or not "No"

nrow(main_roster) #344
#Three removed


# Save the updated dataset under the same name
if (bExport == 1){
  write.xlsx(main_roster, analysis_ready_main_roster_file, rowNames = FALSE)
}
message("Removed rows where `INT02` is 'No'. Preserved rows with `NA`. Updated file saved to `analysis_ready_main_roster.xlsx`.")


# ******************************************************
# Step 16: Standardizes `LOC01` by recoding it to 1 (COUNTRY), 2 (INTERNATIONAL ORG), or 3 (CSO), ----
# and creates `morganization` with all text in the `organization` column capitalized.
# ******************************************************

# Recode `LOC01` values
contr <- main_roster %>%
  group_by(LOC01) %>%
  tally()
contr
# LOC01                                n
# 1                                  180
# 2                                   62
# 3                                    4
# CIVIL SOCIETY ORGANIZATION (CSO)     2
# COUNTRY NSS/NSO/LINE MINISTRY       64
# INTERNATIONAL ORGANIZATION          23
# NA                                   2


main_roster <- main_roster %>%
  mutate(
    LOC01 = case_when(
      LOC01 == "COUNTRY NSS/NSO/LINE MINISTRY" ~ 1,    # Recode as 1
      LOC01 == "INTERNATIONAL ORGANIZATION" ~ 2,        # Recode as 2
      LOC01 == "CIVIL SOCIETY ORGANIZATION (CSO)" ~ 3,  # Recode as 3
      LOC01 == "1" ~ 1,                                 # Keep "1" as is, converted to numeric
      LOC01 == "2" ~ 2,                                 # Keep "2" as is, converted to numeric
      LOC01 == "3" ~ 3,                                 # Keep "3" as is, converted to numeric
      TRUE ~ NA_real_                                  # For everything else, convert to NA
    )
  ) %>%
  mutate(
    LOC01 = as.numeric(LOC01)  # Ensure everything is numeric (NA values remain NA)
  )



contr <- main_roster %>%
  group_by(LOC01) %>%
  tally()
contr
# LOC01     n
#     1   244
#     2    85
#     3     6
#    NA     2


contr <- main_roster %>%
  group_by(morganization, organization) %>%
  tally()


# Standardize `organization` text
main_roster <- main_roster %>%
  mutate(
    morganization = toupper(organization), # Convert all text in `organization` to uppercase
  )

# Save the updated dataset under the same name
if (bExport == 1){
  write.xlsx(main_roster, analysis_ready_main_roster_file, rowNames = FALSE)
}
message("Recode of `LOC01` and capitalization of `organization` completed. Updated file saved to `analysis_ready_main_roster.xlsx`.")

# Step 16b: Apply org crosswalk to main_roster ----

crosswalk <- data_clean_data[["inst_code_errata"]] %>%
  filter(!action %in% c("ok", "review")) %>%
  mutate(across(c(old_code, old_name), ~if_else(is.na(.) | . %in% c("", "NA"), "__NA__", .))) %>%
  distinct(old_code, old_name, .keep_all = TRUE)

# Pass 1: join on code + name
main_roster <- main_roster %>%
  mutate(
    LOC06_4_join       = if_else(is.na(LOC06_4)       | LOC06_4       %in% c("", "NA"), "__NA__", LOC06_4),
    morganization_join = if_else(is.na(morganization) | morganization %in% c("", "NA"), "__NA__", morganization)
  ) %>%
  left_join(
    crosswalk %>% select(old_code, old_name, new_code, canonical_name),
    by = c("LOC06_4_join" = "old_code", "morganization_join" = "old_name")
  ) %>%
  mutate(
    LOC06_4       = coalesce(new_code,       LOC06_4),
    morganization = coalesce(canonical_name, morganization)
  ) %>%
  select(-new_code, -canonical_name, -LOC06_4_join, -morganization_join)

# Pass 2: fallback — join on name only for rows still not canonical
crosswalk_name_only <- data_clean_data[["inst_code_errata"]] %>%
  filter(!action %in% c("ok", "review")) %>%
  filter(!is.na(old_name), !old_name %in% c("", "NA")) %>%
  distinct(old_name, .keep_all = TRUE) %>%
  select(old_name, new_code, canonical_name)

main_roster <- main_roster %>%
  left_join(crosswalk_name_only, by = c("morganization" = "old_name")) %>%
  mutate(
    LOC06_4       = coalesce(new_code,       LOC06_4),
    morganization = coalesce(canonical_name, morganization)
  ) %>%
  select(-new_code, -canonical_name)

main_roster <- main_roster %>%
  mutate(morganization = if_else(
    LOC06_4 == "SOM_DI",
    "DIRECTORATE OF NATIONAL STATISTICS",
    morganization
  ))


# ******************************************************
# Step 17: Converts `LOC04` text values to numeric codes: 1 = NATIONAL, 2 = SUB-NATIONAL, 6 = OTHER. ----
# Standardizes `LOC04` to contain only numeric values (1, 2, or 6).
# ******************************************************


# Standardize `LOC04` values to numeric
contr <- main_roster %>%
  group_by(LOC04) %>%
  tally()
contr
# LOC04            n
# 1              156
# 2                5
# 6                1
# NATIONAL        58
# OTHER            2
# SUB-NATIONAL     4
# NA             111

main_roster <- main_roster %>%
  mutate(
    LOC04 = case_when(
      LOC04 == "01" | LOC04 == "1" | LOC04 == "NATIONAL" ~ 1,          # NATIONAL → 1
      LOC04 == "02" | LOC04 == "2" | LOC04 == "SUB-NATIONAL" ~ 2,       # SUB-NATIONAL → 2
      LOC04 == "06" | LOC04 == "6" | LOC04 == "OTHER" ~ 6,              # OTHER → 6
      TRUE ~ NA_real_                                                    # For all other cases, convert to NA
    )
  ) %>%
  mutate(
    LOC04 = as.numeric(LOC04)  # Ensure column is numeric
  )


contr <- main_roster %>%
  group_by(LOC04) %>%
  tally()
contr
# LOC04     n
#    1   214
#    2     9
#    6     3
#   NA   111
#OK


# Save the updated dataset under the same name
if (bExport == 1){
  write.xlsx(main_roster, analysis_ready_main_roster_file, rowNames = FALSE)
}
message("Standardized `LOC04` to numeric values (1, 2, or 6). Updated file saved to `analysis_ready_main_roster.xlsx`.")



# ******************************************************
# Step 18: Standardizes response variables in `analysis_ready_main_roster`----
# to numeric values: 01 = YES, 02 = NO, 08 = DON'T KNOW, 09 = NO RESPONSE.
# Applies to: PRO01A, FPR01, GRF02, ACT02, ACT03, ACT05, FOL01, FOC04A
# ******************************************************

main_roster_file <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_main_roster.xlsx")

ls(main_roster)

contr <- main_roster %>% 
  group_by(FOL01) %>% 
  tally()

# Recode to numeric using unified logic
main_roster <- main_roster %>%
  mutate(across(
    c(PRO01, PRO01A, FPR01, GRF02, ACT02, ACT03, ACT05, FOL01, FOC04A),
    ~ case_when(
      .x %in% c("01", "1", "YES", "OUI", "SÍ") ~ 1,
      .x %in% c("02", "2", "NO", "NON") ~ 2,
      .x %in% c("08", "8", "DON'T KNOW", "NE SAIT PAS", "NO SABE") ~ 8,
      .x %in% c("09", "9", "NO RESPONSE") ~ 9,
      TRUE ~ NA_real_
    ),
    .names = "{.col}"
  )) %>%
  mutate(across(c(PRO01, PRO01A, FPR01, GRF02, ACT02, ACT03, ACT05, FOL01, FOC04A), as.numeric))

# Save the updated dataset
if (bExport == 1){
  write.xlsx(main_roster, main_roster_file, rowNames = FALSE)
}
message("Response variables standardized in `analysis_ready_main_roster.xlsx` and saved.")



# ******************************************************
# Step 19: Correct `ACT04` in `analysis_ready_main_roster`----
# Historical data don't have ACT04, it directly starts with ACT04.A
# The idea is to recode to have unified new and old data
# ******************************************************
# 
# main_roster_file <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_main_roster.xlsx")
# 
# ls(main_roster)
# 
# contr <- main_roster %>% 
#   group_by(year, ACT04, ACT04.A, ACT04.B, ACT04.C, ACT04.D, ACT04.E, ACT04.F, 
#            ACT04.G, ACT04.H, ACT04.X, ACT04.Z) %>% 
#   tally()
# #Problem in 2024
# 
# # Recode manually to numeric
# main_roster <- main_roster %>%
#   # Fill empty ACT04 data with character that are in ACT04.A
#   mutate(ACT04   = ifelse(is.na(ACT04) == T & ACT04.A %not in% c(0, 1), ACT04.A, ACT04),
#          
#          ACT04.A = ifelse(grepl("IRRS", ACT04), 1,
#                    ifelse(ACT04.A %not in% c(0, 1), NA, ACT04.A)),
#          
#          ACT04.B = ifelse(grepl("IRIS", ACT04), 1, ACT04.B),
#          
#          ACT04.C = ifelse(grepl("IROSS", ACT04), 1, ACT04.C),
#          
#          ACT04.D = ifelse(grepl("EXTERNAL", ACT04), 1, ACT04.D),
# 
#          ACT04.E = ifelse(grepl("ANNUAL", ACT04), 1, ACT04.E),
#          
#          ACT04.F = ifelse(grepl("PROMOTIONAL", ACT04), 1, ACT04.F),
#          
#          ACT04.G = ifelse(grepl("REVISED", ACT04), 1, ACT04.G),
#          
#          ACT04.H = ifelse(grepl("METHODOLOGICAL", ACT04), 1, ACT04.H),
#          
#          ACT04.X = ifelse(grepl("OTHER", ACT04), 1, ACT04.X),
#          
#          ACT04.Z = ifelse(grepl("KNOW", ACT04), 1, ACT04.Z)) %>%
#   
#   mutate(across(c(ACT04.A, ACT04.B, ACT04.C, ACT04.D, ACT04.F, ACT04.G, ACT04.H, 
#                   ACT04.I, ACT04.X, ACT04.Z), as.numeric)) 
# 
# # Save the updated dataset
# if (bExport == 1){
#   write.xlsx(main_roster, main_roster_file, rowNames = FALSE)
# }
# message("Response variables standardized in `analysis_ready_main_roster.xlsx` and saved.")
# 



# ******************************************************
# Step 20: Rename `ACT06` in `analysis_ready_main_roster`----
# Variables ACT06 starts from ACT06.A, but should start with ACT06
# ******************************************************

main_roster_file <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_main_roster.xlsx")

ls(main_roster)

contr <- main_roster %>% 
  group_by(year, ACT06.A, ACT06.B, ACT06.C, ACT06.D, ACT06.E, ACT06.F, ACT06.X, ACT06.Z) %>% 
  tally()
#OK, just names are wrong

# Rename variables
main_roster <- main_roster %>%
  rename(ACT06   = ACT06.A,
         ACT06.A = ACT06.B, 
         ACT06.B = ACT06.C,
         ACT06.C = ACT06.D, 
         ACT06.D = ACT06.E, 
         ACT06.E = ACT06.F) %>% 
  mutate(across(c(ACT06.A, ACT06.B, ACT06.C, ACT06.D, ACT06.E, ACT06.X, ACT06.Z), as.numeric)) 


contr <- main_roster %>% 
  group_by(ACT06, ACT06.A, ACT06.B, ACT06.C, ACT06.D, ACT06.E, ACT06.X, ACT06.Z) %>% 
  tally()


# Save the updated dataset
if (bExport == 1){
  write.xlsx(main_roster, main_roster_file, rowNames = FALSE)
}
message("Response variables standardized in `analysis_ready_main_roster.xlsx` and saved.")



# ******************************************************
# Step 21: Recode `LOC01A` in `analysis_ready_main_roster`----
# Standardize LOC01A to character
# ******************************************************

main_roster_file <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_main_roster.xlsx")
main_roster_file_csv <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_main_roster.csv")

ls(main_roster)

contr <- main_roster %>% 
  group_by(LOC01A) %>% 
  tally()

# Recode to numeric using unified logic
main_roster <- main_roster %>%
  mutate(LOC01A   = ifelse(LOC01A %in% c('GLOBAL'), 1, 
                           ifelse(LOC01A %in% c('NATIONAL'), 2, LOC01A))) %>% 
  mutate(across(c(LOC01A), as.numeric))


contr <- main_roster %>% 
  group_by(LOC01A) %>% 
  tally()


# Save the updated dataset
if (bExport == 1){
  write.xlsx(main_roster, main_roster_file, rowNames = FALSE)
}
message("Response variables standardized in `analysis_ready_main_roster.xlsx` and saved.")




# ******************************************************
# Step 22: Add `morganization` and `mcountry` from `analysis_ready_main_roster` ----
# to `analysis_ready_group_roster` based on `pindex2`.
# Handles multiple entries for the same `pindex2`.
# ******************************************************


# Convert `pindex2` in both datasets to numeric for a proper join
main_roster <- main_roster %>%
  mutate(pindex2 = as.numeric(pindex2),
         bAll    = 1) 


group_roster <- group_roster %>%
  mutate(pindex2 = as.numeric(pindex2),
         bAll    = 1)


nrow(group_roster) #410

# Merge `morganization` and `mcountry` into `group_roster` based on `pindex2`
group_roster <- group_roster %>%
  left_join(
    main_roster %>% select(pindex2, morganization, mcountry, bAll), # Select relevant columns
    by = "pindex2" # Join on `pindex2`
  )

nrow(group_roster) #410
# OK no duplicates


contrmerge <- group_roster %>%
  group_by(bAll.x, bAll.y) %>%
  tally()
contrmerge
# bAll.x bAll.y     n
#   1      1      409
#   1     NA        1


test <- group_roster %>%
  filter(is.na(bAll.y))
#One is missing, because index 73 has been deleted in the Data Clean files


contr <- group_roster %>%
  group_by(mcountry.x, mcountry.y) %>%
  tally()

# Take what was already done, and add the ones from 2025 
#Also remove the data deleted from the data_cleaning sheet
group_roster <- group_roster %>%
  filter(is.na(bAll.y) == F) %>%
  select(-bAll.x, -bAll.y) %>%
  mutate(bAll = 1) %>%
  # Take what was already done, and add the ones from 2025 
  mutate(mcountry      = ifelse(is.na(mcountry.y), mcountry.x, mcountry.y), 
         morganization = ifelse(is.na(morganization.y), morganization.x, morganization.y))


contr_cnt <- group_roster %>%
  group_by(mcountry, mcountry.x, mcountry.y) %>%
  tally()

contr_org <- group_roster %>%
  group_by(morganization, morganization.x, morganization.y) %>%
  tally()


contr <- group_roster %>%
  group_by(morganization, mcountry) %>%
  tally()
#Still the one missing from data cleaning



# Save the updated group roster
group_roster_file <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_group_roster.xlsx")
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Added `morganization` and `mcountry` to `analysis_ready_group_roster.xlsx`. Updated file saved.")

# ******************************************************
# Step 23: Standardizes `PRO03B` in `analysis_ready_group_roster` to numeric values: ----
# 1 = GLOBAL, 2 = REGIONAL, 3 = COUNTRY, 8 = DON'T KNOW.
# ******************************************************


contr <- group_roster %>%
  group_by(PRO03B) %>%
  tally()
contr
# PRO03B       n
# 1           34
# 2           34
# 3           85
# 8            1
# COUNTRY     25
# GLOBAL       8
# NA         150
# REGIONAL    18
# NA          54



# Recode `PRO03B` values to numeric
group_roster <- group_roster %>%
  select(-c(mcountry.x, mcountry.y, morganization.x, morganization.y)) %>%
  mutate(
    PRO03B = case_when(
      PRO03B == "01" | PRO03B == "1" | PRO03B == "GLOBAL" ~ 1,            # GLOBAL → 1
      PRO03B == "02" | PRO03B == "2" | PRO03B == "REGIONAL" ~ 2,          # REGIONAL → 2
      PRO03B == "03" | PRO03B == "3" | PRO03B == "COUNTRY" ~ 3,           # COUNTRY → 3
      PRO03B == "08" | PRO03B == "8" | PRO03B == "DON'T KNOW" ~ 8,        # DON'T KNOW → 8
      TRUE ~ NA_real_                                                     # Keep numeric if already valid
    ) 
  ) %>%
  mutate(
    PRO03B = as.numeric(PRO03B)  # Ensure column is numeric
  )

contr <- group_roster %>%
  group_by(PRO03B) %>%
  tally()
contr
# PRO03B    n
#     1    42
#     2    52
#     3   110
#     8     1
#    NA   204
#OK

# Save the updated dataset under the same name
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Recode of `PRO03B` completed. All values are now numeric. Updated file saved to `analysis_ready_group_roster.xlsx`.")



# ******************************************************
# Step 24: Standardizes `PRO03D` in `analysis_ready_group_roster` to numeric values: ----
# 1 = NATIONAL, 2 = INSTITUTIONAL, 8 = DON'T KNOW.
# ******************************************************


contr_fun <- function(df, var) {
  df %>%
    dplyr::group_by({{ var }}) %>%
    dplyr::tally()
}

contr_fun(group_roster, PRO03D)
# PRO03D            n
# 1                56
# 2                47
# 3                 2
# 8                 1
# INSTITUTIONAL    12
# NA              206
# NATIONAL          5
# NA               80


# Recode `PRO03D` values to numeric
group_roster <- group_roster %>%
  mutate(
    PRO03D = case_when(
      PRO03D == "1" | PRO03D == "NATIONAL" ~ 1,             # NATIONAL → 1
      PRO03D == "2" | PRO03D == "INSTITUTIONAL" ~ 2,        # INSTITUTIONAL → 2
      PRO03D == "8" | PRO03D == "DON'T KNOW" ~ 8,           # DON'T KNOW → 8
      TRUE ~ NA_real_                            # Keep numeric if already valid
    )
  ) %>%
  mutate(
    PRO03D = as.numeric(PRO03D)  # Ensure column is numeric
  )

contr_fun(group_roster, PRO03D)
# PRO03D     n
#     1     61
#     2     59
#     8      1
#    NA    288




# Save the updated dataset under the same name
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Recode of `PRO03D` completed. All values are now numeric. Updated file saved to `analysis_ready_group_roster.xlsx`.")

# ******************************************************
# Step 25: Standardizes `PRO06` in `analysis_ready_group_roster` to numeric values: ----
# 01 = DESIGN/PLANNING, 02 = IMPLEMENTATION, 03 = COMPLETED,
# 06 = OTHER, 08 = DON’T KNOW.
# ******************************************************


contr_fun(group_roster, PRO06)
# PRO06               n
#  1                  62
#  2                 139
#  3                  81
#  6                   8
#  8                  11
#  COMPLETED          32
#  DESIGN/PLANNING    16
#  IMPLEMENTATION     46
#  NA                  3
#  OTHER              11


# Recode `PRO06` values to numeric
group_roster <- group_roster %>%
  mutate(
    PRO06 = case_when(
      PRO06 == "01" | PRO06 == "1" | PRO06 == "DESIGN/PLANNING" | PRO06 == "CONCEPTION/PLANIFICATION" | PRO06 == "DISEÑO/PLANIFICACIÓN" ~ 1,
      PRO06 == "02" | PRO06 == "2" | PRO06 == "IMPLEMENTATION" | PRO06 == "MISE EN ŒUVRE" | PRO06 == "IMPLEMENTACIÓN" ~ 2,
      PRO06 == "03" | PRO06 == "3" | PRO06 == "COMPLETED" | PRO06 == "ACHEVÉ" | PRO06 == "FINALIZADA" ~ 3,
      PRO06 == "06" | PRO06 == "6" | PRO06 == "OTHER" | PRO06 == "AUTRE" | PRO06 == "OTROS" ~ 6,
      PRO06 == "08" | PRO06 == "8" | PRO06 == "DON’T KNOW" | PRO06 == "NE SAIT PAS" | PRO06 == "NO SABE" ~ 8,
      TRUE ~ NA_real_ # Keep numeric values as is
    )
  ) %>%
  mutate(
    PRO06 = as.numeric(PRO06)  # Ensure column is numeric
  )

contr_fun(group_roster, PRO06)
# PRO06     n
#    1    78
#    2   185
#    3   113
#    6    19
#    8    11
#   NA     3



# Save the updated dataset under the same name
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Recode of `PRO06` completed. All values are now numeric. Updated file saved to `analysis_ready_group_roster.xlsx`.")




# ******************************************************
# Step 26: Rename and standardize `PRO08` in `analysis_ready_group_roster`: ----
# PRO08A: 1 = YES,  2 = NO,  8 = DON’T KNOW
# PRO08C: 1 = REFUGEES, IDPs, OR STATELESS PERSONS ONLY, 2 = NATIONAL POPULATION INCLUDED, 8 = DON’T KNOW
# ******************************************************


# Recode `PRO06` values to numeric
group_roster <- group_roster %>%
  rename(PRO08A = PRO08.A,
         PRO08B = PRO08.B,
         PRO08C = PRO08.C) %>% 
  mutate(
    PRO08A = ifelse(PRO08A %in% c("DON'T KNOW"), 8, 
                    ifelse(PRO08A %in% c("YES"), 1, 
                           ifelse(PRO08A %in% c("NO"), 2, NA))),
    
    PRO08C = ifelse(PRO08C %in% c("DON'T KNOW"), 8, 
                    ifelse(PRO08C %in% c("REFUGEES, IDPs, OR STATELESS PERSONS ONLY"), 1, 
                           ifelse(PRO08C %in% c("NATIONAL POPULATION INCLUDED"), 2, NA)))
  ) %>%
  mutate(across(c(PRO08A, PRO08C, 
                  PRO08B.A, PRO08B.B, PRO08B.C, PRO08B.D, PRO08B.E, PRO08B.F, 
                  PRO08B.G, PRO08B.H, PRO08B.X), as.numeric))

contr_fun(group_roster, PRO06)
# PRO06     n
#    1    78
#    2   185
#    3   113
#    6    19
#    8    11
#   NA     3



contr <- group_roster %>% 
  group_by(year, PRO08) %>% 
  tally()



# Save the updated dataset under the same name
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Recode of `PRO06` completed. All values are now numeric. Updated file saved to `analysis_ready_group_roster.xlsx`.")


# ******************************************************
# Step 27: Standardizes `PRO09`, `PRO13B`, `PRO19`, and `PRO21` in `analysis_ready_group_roster` ----
# to numeric values: 01 = YES, 02 = NO, 08 = DON'T KNOW, 09 = NO RESPONSE.
# ******************************************************


contr_fun(group_roster, PRO09)
contr_fun(group_roster, PRO13B)
contr_fun(group_roster, PRO19)
contr_fun(group_roster, PRO21)



# Recode PRO09, PRO13B, PRO19, and PRO21 to numeric with consistent logic
group_roster <- group_roster %>%
  mutate(across(
    c(PRO09, PRO13B, PRO19, PRO21),
    ~ case_when(
      .x %in% c("01", "1", "YES", "OUI", "SÍ") ~ 1,                                           # YES → 1
      .x %in% c("02", "2", "NO", "NON") ~ 2,                                                  # NO → 2
      .x %in% c("08", "8", "DON'T KNOW", "NE SAIT PAS", "NO SABE") ~ 8,                      # DON'T KNOW → 8
      .x %in% c("09", "9", "NO RESPONSE") ~ 9,                                               # NO RESPONSE → 9
      TRUE ~ NA_real_                                                                        # All others → NA
    ),
    .names = "{.col}"
  )) %>%
  mutate(across(c(PRO09, PRO13B, 
                  PRO13CA, PRO13CB, PRO13CC, PRO13CD, PRO13CE, PRO13CX,
                  PRO19, PRO21, PRO22A, PRO22B, PRO22C, PRO22D, PRO22E, PRO22F, PRO22X,
                  index,
                  PRO07.A, PRO07.B, PRO07.C, PRO07.X,
                  PRO10.A, PRO10.B, PRO10.C, PRO10.Z,
                  PRO18.A, PRO18.B, PRO18.C,
                  PRO20.A, PRO20.B, PRO20.C, PRO20.D, PRO20.E, PRO20.F, PRO20.G, PRO20.H, 
                  PRO20.I, PRO20.J, PRO20.X, PRO20.Z,
                  UPD25), as.numeric))  # Ensure numeric type

contr_fun(group_roster, PRO09)
contr_fun(group_roster, PRO13B)
contr_fun(group_roster, PRO19)
contr_fun(group_roster, PRO21)



# Save the updated dataset under the same name
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Recode of `PRO09` completed. All values are now numeric. Updated file saved to `analysis_ready_group_roster.xlsx`.")

# ******************************************************
# Step 28: Standardizes `PRO14` in `analysis_ready_group_roster` ----
# 1 = YES, 2 = NO.
# ******************************************************


contr_fun(group_roster, PRO14)
# PRO14          n
# 1             92
# 2            132
# DON'T KNOW    13
# NA            80
# NO            58
# YES           34


# Recode `PRO14` & PRO14B` values to numeric
group_roster <- group_roster %>%
  mutate(
    PRO14 = case_when(
      PRO14 == "01" | PRO14 == "1" | PRO14 == "YES" | PRO14 == "OUI" | PRO14 == "SÍ" ~ 1,  # YES → 1
      PRO14 == "02" | PRO14 == "2" | PRO14 == "NO" | PRO14 == "NON" ~ 2,                  # NO → 2
      TRUE ~ NA_real_                                                            # Keep numeric values as is
    ),
    PRO14B = case_when(
      PRO14B == "01" | PRO14B == "1" | PRO14B == "YES" | PRO14B == "OUI" | PRO14B == "SÍ" ~ 1,  # YES → 1
      PRO14B == "02" | PRO14B == "2" | PRO14B == "NO" | PRO14B == "NON" ~ 2,                  # NO → 2
      PRO06 == "08" | PRO06 == "8" | PRO06 == "DON’T KNOW" | PRO06 == "NE SAIT PAS" | PRO06 == "NO SABE" ~ 8, # DON'T KNOW → 8
      TRUE ~ NA_real_                                                            # Keep numeric values as is
    )
  ) %>%
  mutate(
    PRO14  = as.numeric(PRO14),  # Ensure column is numeric
    PRO14B = as.numeric(PRO14B)   
  )

contr_fun(group_roster, PRO14)
# PRO14     n
#    1    126
#    2    190
#   NA     93

contr_fun(group_roster, PRO14B)
# PRO14B     n
#     1     20
#     2     48
#     8     11
#    NA    330


# Save the updated dataset under the same name
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Recode of `PRO14` completed. All values are now numeric. Updated file saved to `analysis_ready_group_roster.xlsx`.")



# ******************************************************
# Step 29: Standardizes `PRO15` in `analysis_ready_group_roster` ----
# 1 = YES, 2 = NO.
# ******************************************************


contr_fun(group_roster, PRO15)
# PRO15          n
# 1             36
# 2            151
# DON'T KNOW     5
# NA           117
# NO            51
# YES           27
# NA            22



# Recode `PRO15` values to numeric
group_roster <- group_roster %>%
  mutate(
    PRO15 = case_when(
      PRO15 == "01" | PRO15 == "1" | PRO15 == "YES" | PRO15 == "OUI" | PRO15 == "SÍ" ~ 1,  # YES → 1
      PRO15 == "02" | PRO15 == "2" | PRO15 == "NO" | PRO15 == "NON" ~ 2,                  # NO → 2
      TRUE ~ NA_real_                                                           # Keep numeric values as is
    ) 
  ) %>%
  mutate(
    PRO15 = as.numeric(PRO15)  # Ensure column is numeric
  )


contr_fun(group_roster, PRO15)
# PRO15     n
#    1      63
#    2     202
#   NA     144
#OK



# Save the updated dataset under the same name
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Recode of `PRO15` completed. All values are now numeric. Updated file saved to `analysis_ready_group_roster.xlsx`.")



# ******************************************************
# Step 30: Standardizes `PRO17` in `analysis_ready_group_roster` ----
# 1 = YES, 2 = NO.
# ******************************************************

contr_fun(group_roster, PRO17)
# PRO17          n
# 1            165
# 2             85
# DON'T KNOW     3
# NA            54
# NO            21
# YES           82


# Recode `PRO17` values to numeric
# Cat DON'T KNOW is now with NA - CHECK IF IN ORDER
group_roster <- group_roster %>%
  mutate(
    PRO17 = case_when(
      PRO17 == "01" | PRO17 == "1" | PRO17 == "YES" | PRO17 == "OUI" | PRO17 == "SÍ" ~ 1,  # YES → 1
      PRO17 == "02" | PRO17 == "2" | PRO17 == "NO" | PRO17 == "NON" ~ 2,                  # NO → 2
      TRUE ~ NA_real_                                                           # Keep numeric values as is
    )
  ) %>%
  mutate(
    PRO17 = as.numeric(PRO17)  # Ensure column is numeric
  )



contr_fun(group_roster, PRO17)
# PRO17     n
#    1    247
#    2    106
#   NA     57



# Save the updated dataset under the same name
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Recode of `PRO17` completed. All values are now numeric. Updated file saved to `analysis_ready_group_roster.xlsx`.")


# ******************************************************
# Step 31: Add manual data from `PRO18` in `analysis_ready_group_roster`: ----
# _Step 1: Export data to manually fill
# _Step 2: join the data manually filled
# ******************************************************

# Step 1:
#Take historical data
#Already filled from previous years
#The path might have to be updated
old_PRO <- read_csv("EGRISS GAIN Survey 2024/10 Data/Analysis Ready Files/analysis_ready_group_roster.csv")

#Select relevant columns

old_PRO1 <- old_PRO %>% 
  select(pindex2, index, PRO18, PRO18.A, PRO18.B, PRO18.C) %>% 
  mutate(bAll = 1)


nrow(old_PRO1)
nrow(distinct(old_PRO1, pindex2, index))

old_PRO1a <- old_PRO1 %>% 
  group_by(pindex2, index) %>% 
  mutate(seq = seq_along(pindex2),
         maxSeq = max(seq)) %>%  ungroup() %>% 
  filter(seq == maxSeq) %>% 
  select(-seq, -maxSeq)


#Take new data (already ready file exported)
new_PRO <- group_roster

new_PRO1 <- new_PRO %>% 
  select(pindex2, index, PRO18) %>% 
  mutate(bAll = 1)

nrow(new_PRO1)
nrow(distinct(new_PRO1, pindex2, index))

new_PRO1a <- new_PRO1 %>% 
  group_by(pindex2, index) %>% 
  mutate(seq = seq_along(pindex2),
         maxSeq = max(seq)) %>%  ungroup() %>% 
  filter(seq == maxSeq) %>% 
  select(-seq, -maxSeq)



#Put two data together 
tot_PRO <- merge(new_PRO1a, old_PRO1a, by = c('pindex2', 'index'), all = T)

contrmerge <- tot_PRO %>% 
  group_by(bAll.x, bAll.y) %>% 
  tally()
contrmerge
# bAll.x bAll.y     n
#   1      1      304
#   1     NA      120



final_PRO <- tot_PRO %>% 
  mutate(
    PRO18 = ifelse(is.na(PRO18.y), PRO18.x, PRO18.y)) %>% 
  select(-PRO18.x, -PRO18.y)

# Export to manually fill. To be carefull because it will overwrite work
# if (bExport == 10){
#   write.xlsx(final_PRO, str_c(egr_yr, "/06 Data Cleaning/01 Input/PRO18_ToComplete.xlsx"))
# }


# Step 2:
# Import filled data, and join it to current data
Pro18_actu <- read_excel(str_c(egr_yr, "/06 Data Cleaning/01 Input/PRO18_ToComplete.xlsx")) %>% 
  mutate(bAll = 1)


nrow(Pro18_actu)
nrow(distinct(Pro18_actu, index, pindex2))
nrow(group_roster)

group_roster_temp <- merge(group_roster %>% mutate(bAll = 1) %>% select(-c(PRO18, PRO18.A, PRO18.B, PRO18.C)), Pro18_actu, by = c('index', 'pindex2'), all.x = T)

nrow(group_roster)

contrmerge <- group_roster_temp %>% 
  group_by(bAll.x, bAll.y) %>% 
  tally()
contrmerge
# bAll.x bAll.y     n
#     1      1     424
#OK


nrow(group_roster) #424
nrow(group_roster_temp) #424
nrow(Pro18_actu) #329
#OK

ls(group_roster_temp)

group_roster <- group_roster_temp %>% 
  select(-bAll.x, -bAll.y)


# Save the updated group roster
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("PRO18 updates have been added in `analysis_ready_group_roster.xlsx` based on `pindex2`. Updated file saved.")



# ******************************************************
# Step 32: Add manual data from `FOC02` in `analysis_ready_main_roster`: ----
# _Step 1: Export data to manually fill
# _Step 2: join the data manually filled
# ******************************************************

# Step 1
#Take historical data
#Already filled from previous years
#The path might have to be updated
old_FOC <- read_csv("EGRISS GAIN Survey 2024/10 Data/Analysis Ready Files/analysis_ready_main_roster.csv")

#Select relevant columns

old_FOC1 <- old_FOC %>% 
  select(pindex2, FOC02, FOC02A, FOC02B) %>% 
  mutate(bAll = 1)



#Take new data (already ready file exported)
new_FOC <- main_roster

new_FOC1 <- new_FOC %>% 
  select(pindex2, FOC02) %>% 
  mutate(bAll = 1)



#Put two data together
tot_FOC <- merge(new_FOC1, old_FOC1, by = c('pindex2'), all = T)

contrmerge <- tot_FOC %>% 
  group_by(bAll.x, bAll.y) %>% 
  tally()
contrmerge
# bAll.x bAll.y     n
#     1      1      248
#     1     NA      96

contr <- tot_FOC %>% 
  filter(FOC02.x != FOC02.y)
#Just character type causing differences


#Takes time
if (bExport == 10){
  final_FOC <- tot_FOC %>% 
    mutate(
      FOC02 = ifelse(is.na(FOC02.y), FOC02.x, FOC02.y),
      
      FOC02 = map_chr(FOC02, function(x) {
        if (is.na(x) || x == "NA" || x == "") {
          return(NA_character_)
        }
        
        translate(x, from = "auto", to = "en", trim_str = TRUE)
      })
    ) %>% 
    select(-FOC02.x, -FOC02.y)
  
  #Carefull it overwrites
  # write.xlsx(final_FOC, str_c(egr_yr, "/06 Data Cleaning/01 Input/FOC02_ToComplete.xlsx"))
}


# Step 2:
# Import filled data, and join it to current data
Foc02_actu <- read_excel(str_c(egr_yr, "/06 Data Cleaning/01 Input/FOC02_ToComplete.xlsx")) %>% 
  mutate(bAll = 1)


main_roster_temp <- merge(main_roster %>% mutate(bAll = 1) %>% select(-FOC02), Foc02_actu, by = c('pindex2'), all.x = T)

contrmerge <- main_roster_temp %>% 
  group_by(bAll.x, bAll.y) %>% 
  tally()
contrmerge
# bAll.x bAll.y     n
#     1      1     344

# contr <- main_roster_temp %>%
#   select(FOC02.x, FOC02.y, FOC02A, FOC02B) %>%
#   filter(FOC02.x != FOC02.y)

nrow(main_roster) #344
nrow(main_roster_temp) #344
nrow(Foc02_actu) #344
#OK

ls(main_roster_temp)

main_roster <- main_roster_temp %>% 
  select(-bAll.x, -bAll.y)


# Save the updated group roster
if (bExport == 1){
  write.xlsx(main_roster, main_roster_file, rowNames = FALSE)
  write.csv(main_roster, main_roster_file_csv)
}
message("PRO18 updates have been added in `analysis_ready_main_roster.xlsx` based on `pindex2`. Updated file saved.")




# ******************************************************
# Step 33: Copies `LOC01` from `analysis_ready_main_roster` to `gLOC01` in `analysis_ready_group_roster` based on `pindex2`. ----
# ******************************************************

# Ensure `pindex2` is numeric in both datasets
summary(main_roster$pindex2) #OK
summary(group_roster$pindex2) #OK

summary(main_roster$LOC01) #OK
summary(group_roster$gLOC01) #OK

contr <- group_roster %>%
  mutate(gLOC01 = as.numeric(gLOC01))
summary(contr$gLOC01) #OK

contr_fun(group_roster, gLOC01)
# gLOC01     n
# 1        150
# 2        152
# 3          2
# NA       106

contr <- group_roster %>% 
  group_by(year, gLOC01) %>% 
  tally()


ls(main_roster)
ls(group_roster)

nrow(group_roster) #409

# Join `LOC01` from `main_roster` to `group_roster` and create `gLOC01`
group_roster <- group_roster %>%
  select(-gLOC01) %>% 
  left_join(
    main_roster %>% select(pindex2, LOC01) %>% rename(gLOC01 = LOC01), # Select `pindex2` and `LOC01` from `main_roster`
    by = "pindex2" # Match on `pindex2`
  )

nrow(group_roster)

ls(group_roster)

contr <- group_roster %>% 
  group_by(year, gLOC01) %>% 
  tally()

summary(group_roster$gLOC01)
#OK numeric


# Save the updated group roster
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Copied `LOC01` to `gLOC01` in `analysis_ready_group_roster.xlsx` based on `pindex2`. Updated file saved.")



# ******************************************************
# Step 34: Creates `g_conled` in `analysis_ready_group_roster` to categorize projects: ----
# 1 = Country-led, 2 = Institutional-led, 3 = Other (based on `gLOC01` and `PRO03D`).
# ******************************************************

# Check column names for debugging
print(colnames(group_roster))  # Ensure `PRO03D` and `gLOC01` are present


contr <- group_roster %>% 
  group_by(year, PRO03D, gLOC01) %>% 
  tally()
#Some PRO03D are missing




# Create `g_conled` based on `gLOC01` and `PRO03D`
group_roster <- group_roster %>%
  mutate(
    g_conled = case_when(
      gLOC01 == 1 ~ 1, # Country-led if `gLOC01` is 1
      gLOC01 == 2 & PRO03D == 1 ~ 1, # Country-led if `gLOC01` is 2 and `PRO03D` is 1
      gLOC01 == 2 ~ 2, # Institutional-led if `gLOC01` is 2 and `PRO03D` is not 1
      gLOC01 == 3 ~ 3, # Other if `gLOC01` is 3
      TRUE ~ NA_real_ # Assign NA for missing or unmatched cases
    )
  )



contr <- group_roster %>% 
  group_by(g_conled, gLOC01, PRO03D) %>% 
  tally()


# Save the updated group roster
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Created `g_conled` in `analysis_ready_group_roster.xlsx` based on `gLOC01` and `PRO03D`. Updated file saved.")



# ******************************************************
# Step 35: Recodes ISO country codes in `PRO03C` in `analysis_ready_group_roster` ----
# to full country names (e.g., "SOM" → "Somalia"), preserving existing names.
# ******************************************************

contr1 <- group_roster %>% 
  select(pindex2, PRO03, PRO03C, mcountry)

# Define region
group_roster <- group_roster %>% 
  mutate(mcountry = ifelse(mcountry == 'NA', NA, mcountry),
         PRO03C   = ifelse(PRO03C == 'NA', NA, PRO03C)) %>% 
  mutate(mcountry = ifelse(is.na(mcountry), PRO03C, mcountry))


contr2 <- group_roster %>%
  select(pindex2, PRO03, PRO03C, mcountry)

contr <- group_roster %>% 
  group_by(PRO03C, mcountry) %>% 
  tally()



# ******************************************************
# Step 36: This script updates the 'mcountry' field in the 'analysis_ready_group_roster' dataset----
# by mapping ISO country codes to their full names where 'g_conled' equals 1.
# It only updates entries where 'mcountry' is NA and 'PRO03C' contains an ISO code,
# ensuring all updates are relevant to country-led examples.
# ******************************************************

contr <- group_roster %>% 
  group_by(PRO03C, mcountry) %>% 
  tally()


# Update specific cases
# Rename some countries
group_roster <- group_roster %>%
  mutate(mcountry = ifelse(mcountry %in% c("United Kingdom of Great Britain and Northern Ireland"), "United Kingdom",
                    ifelse(mcountry %in% c('Côte d’Ivoire'), "Côte d'Ivoire",
                    ifelse(mcountry %in% c('Palestinian Territories'), 'State of Palestine',
                    ifelse(mcountry %in% c('Turkey'), 'Turkiye',
                    ifelse(mcountry %in% c('Rep. of Chad'), 'Chad',
                    ifelse(mcountry %in% c('Moldova'), 'Republic of Moldova', 
                    ifelse(mcountry %in% c('Congo - Kinshasa'), 'Democratic Republic of the Congo',
                    ifelse(mcountry %in% c('Congo - Brazzaville'), 'Republic of the Congo', 
                    ifelse(mcountry %in% c('Netherlands'), 'Netherlands (Kingdom of the)', mcountry)))))))))) #Has to rename it for the merging


contr <- group_roster %>% 
  group_by(mcountry) %>% 
  tally()

# Save the updated dataset
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("Updated `mcountry` in `analysis_ready_group_roster.xlsx` based on `PRO03C`. Saved to the same file.")



# ******************************************************
# Step 37: Assign Regions to Countries in `analysis_ready_group_roster`----
# - This script maps country names (`mcountry`) to their respective regions.
# - If a country is not in the predefined list, it is assigned "Other".
# ******************************************************

#Was manually added
#If not all country names are matched, please add them, with their unhcr_region and egriss_region
country_region_mapping <- data_clean_data[["country_name"]]


# Ensure `region` column exists in `group_roster` before updating

contr <- group_roster %>% 
  select(pindex2, index, PRO03, PRO03C, mcountry)



contr <- group_roster %>% 
  group_by(region, mcountry) %>% 
  tally()


nrow(group_roster)
# Assign regions to the dataset
group_roster <- group_roster %>%
  mutate(bAlle = 1) %>% 
  select(-region) %>% 
  left_join(country_region_mapping %>% mutate(bAlle = 1), by = "mcountry") 
nrow(group_roster)


contrmerge <- group_roster %>% 
  group_by(bAlle.x, bAlle.y) %>% 
  tally()
contrmerge


#Control here the countries not mapped
contr <- group_roster %>% 
  filter(is.na(bAlle.y)) %>% 
  select(pindex2, mcountry)
#If only NAs, then OK


contr <- group_roster %>% 
  group_by(egriss_region, mcountry) %>% 
  tally()
#OK


group_roster <- group_roster %>% 
  select(-c(bAlle.x, -bAlle.y)) %>% 
  mutate(egriss_region = ifelse(is.na(egriss_region), "Other", egriss_region))


# Save the updated dataset
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file)
}
message("Updated `region` variable in `analysis_ready_group_roster.xlsx`. Saved to the same file.")



# ******************************************************
# Step 38: Assign Regions to Countries in analysis_ready_group_roster2 ----
# This script assigns regions to countries in the analysis_ready_group_roster2 dataset.
# A predefined mapping of country names to their respective regions is used for categorization.
# The region column is added or updated to ensure consistency in geographic classifications.
# The cleaned dataset is then saved in the Analysis Ready Files folder.
# ******************************************************


output_group_roster2_file     <- file.path(analysis_ready_directory, "analysis_ready_group_roster2.xlsx")
output_group_roster2_file_csv <- file.path(analysis_ready_directory, "analysis_ready_group_roster2.csv")

# Load datasets
group_roster2 <- final_version_data[["group_roster2"]]

# Step 1: Check if `X_parent_index` exists
if ("_parent_index" %in% colnames(group_roster2)) {
  group_roster2 <- group_roster2 %>%
    mutate(
      index1 = as.numeric(`_parent_index`)  # Create a new numeric variable without altering the original column
    )
  message("Created `index1` as a numeric version of `X_parent_index`.")
} else {
  stop("Column `X_parent_index` not found in `group_roster2`. Check the dataset.")
}

# Step 2: Add `year` variable as 2025
group_roster2 <- group_roster2 %>%
  mutate(year = year)

# Step 3: Rename `index1` to `pindex1`
group_roster2 <- group_roster2 %>%
  rename(pindex1 = index1)

# Step 4: Create `pindex2` - 8-digit identifier using `year` and `pindex1`
group_roster2 <- group_roster2 %>%
  mutate(
    pindex2 = ifelse(is.na(pindex1), NA, as.numeric(sprintf("%d%04d", year, pindex1))) # Ensure numeric
  )


# Ensure `pindex2` in `main_roster` is also numeric
summary(main_roster$pindex2)

nrow(group_roster2)

# Step 5: Merge `gLOC01`, `morganization`, and `mcountry` from `main_roster`
group_roster2 <- group_roster2 %>%
  inner_join(main_roster %>% select(pindex2, LOC01, morganization, mcountry), by = "pindex2")

nrow(group_roster2)
#Remove the ones filtered in main_roster

# Ensure `region` column is created in `group_roster2`
group_roster2 <- group_roster2 %>%
  mutate(region = NA_character_)  # Create an empty `region` column

# Assign regions to `group_roster2` dataset using `country_region_mapping`
group_roster2 <- group_roster2 %>%
  left_join(country_region_mapping, by = "mcountry") %>%
  mutate(egriss_region = ifelse(is.na(egriss_region), 'Other', egriss_region))  # Use region from mapping, fill missing with "Other"

contr <- group_roster2 %>% 
  group_by(mcountry, region, egriss_region) %>% 
  tally()

# Create `q2025` based on the quarter variable
# Identify the column that starts with "FPR07"
column_name <- colnames(group_roster2)[startsWith(colnames(group_roster2), "FPR07")][1]

# Use the identified column in mutate
contr <- group_roster2 %>% 
  group_by(.data[[column_name]]) %>% 
  tally()

# newcolname <- str_c('q', as.character(year+1))

group_roster2 <- group_roster2 %>%
  mutate(
    q2025 = case_when(
      grepl("Quarter 1", .data[[column_name]]) ~ 1,
      grepl("Quarter 2", .data[[column_name]]) ~ 2,
      grepl("Quarter 3", .data[[column_name]]) ~ 3,
      grepl("Quarter 4", .data[[column_name]]) ~ 4,
      TRUE ~ NA_real_
    )
  )



# Recode FPRO05 
group_roster2 <- group_roster2 %>% 
  rename(FPR05 = `FPR05. What will be the **main** data source or tool used to improve future data collection on these populations through  <span style='color:#3b71b9; font-weight: bold;'>${_FPR02}</span>? Please select one.`) %>% 
  mutate(FPR05 = ifelse(FPR05 %in% c('SURVEY'), 11, 
                 ifelse(FPR05 %in% c('ADMINISTRATIVE DATA'), 12, 
                 ifelse(FPR05 %in% c('CENSUS'), 13, 
                 ifelse(FPR05 %in% c('DATA INTEGRATION'), 14,
                 ifelse(FPR05 %in% c('NON-TRADITIONAL'), 15, 
                 ifelse(FPR05 %in% c('STRATEGY'), 16,
                 ifelse(FPR05 %in% c('GUIDANCE/TOOLKIT'), 17, 
                 ifelse(FPR05 %in% c('WORKSHOP/TRAINING'), 18, 
                 ifelse(FPR05 %in% c('OTHER'), 96, NA))))))))))




# Save the final dataset
if (bExport == 1){
  write.xlsx(group_roster2, output_group_roster2_file, rowNames = FALSE)
  write.csv(group_roster2, output_group_roster2_file_csv)
}
message("Saved `analysis_ready_group_roster2.xlsx` with `pindex2`, `gLOC01`, `morganization`, and `mcountry`. File located at: ", output_group_roster2_file)



# ******************************************************
# Step 39: Update `PRO02A` and `PRO03` Based on `ryear`----
# - If `ryear == 2024`: 
#   - Copy `PRO02A` → `PRO03`
# - If `ryear == 2023, 2022, or 2021`: 
#   - Copy `PRO03` → `PRO02A`
# ******************************************************

# Ensure required columns exist
required_columns <- c("ryear", "PRO02A", "PRO03")
missing_columns <- setdiff(required_columns, colnames(group_roster))

if (length(missing_columns) > 0) {
  stop(paste("Missing required columns:", paste(missing_columns, collapse = ", ")))
}

contr <- group_roster %>% 
  select(index, pindex2, PRO02A, PRO03, PRO03C, mcountry) %>% 
  filter(pindex2 == 20250055)

# Apply conditional updates
group_roster <- group_roster %>%
  mutate(
    PRO03 = ifelse(ryear == year, PRO02A, PRO03),   # Copy `PRO02A` → `PRO03`
    PRO02A = if_else(ryear >= 2021 & ryear < year, PRO03, PRO02A)  # Copy `PRO03` → `PRO02A`
  )

contr <- group_roster %>% 
  group_by(PRO03, PRO02A) %>% 
  tally()


# Save the updated dataset
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file)
}
message("Successfully updated `PRO02A` and `PRO03` based on `ryear` conditions (Step 32).")



# ******************************************************
# Step 40: Overwrite `group_roster` Values Using JDC Data----
# - Matches rows where `pindex2` & `X_index` are identical.
# - Replaces values for specified columns directly.
# - Ensures data integrity by only updating matching records.
# ******************************************************


# Change pindex2 == 20230138 to g_conled = 2 
contr <- group_roster %>% 
  filter(pindex2 == 20230138) %>% 
  select(pindex2, index, PRO02A, g_conled, gLOC01, PRO03D)


group_roster <- group_roster %>%
  mutate(
    g_conled = case_when(
      pindex2 == 20240115 & gLOC01 == 1 ~ 1,  # Country-led if `gLOC01` is 1
      pindex2 == 20240115 & gLOC01 == 2 & PRO03D == 1 ~ 1,  # Country-led if `gLOC01` is 2 and `PRO03D` is 1
      pindex2 == 20240115 & gLOC01 == 2 ~ 2,  # Institutional-led if `gLOC01` is 2 and `PRO03D` is not 1
      pindex2 == 20240115 & gLOC01 == 3 ~ 3,  # Other if `gLOC01` is 3
      pindex2 == 20240115 & gLOC01 == 3 ~ 3,  # Other if `gLOC01` is 3
      pindex2 == 20230138 & index == 91 ~ 2,  # Change the US from national to institutional
      TRUE ~ g_conled  # Keep existing values for other rows
    )
  )

contr <- group_roster %>% 
  filter(pindex2 == 20240115) %>% 
  group_by(pindex2, gLOC01, PRO03D, g_conled) %>% 
  tally()
#OK


contr <- group_roster %>% 
  filter(pindex2 == 20230138) %>% 
  select(pindex2, index, gLOC01, PRO03D, g_conled)



# Save updated version to "analysis_ready_group_roster.xlsx"
analysis_ready_file <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_group_roster.xlsx")
analysis_ready_file_csv <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_group_roster.csv")
if (bExport == 1){
  write.xlsx(group_roster, analysis_ready_file)
  write.csv(group_roster, analysis_ready_file_csv)
}

message("✅ Updated `group_roster` has been saved to: ", analysis_ready_file)

# Save updated version of "group_roster" back to the list
final_version_data[["group_roster"]] <- group_roster


# ******************************************************
#Step 41: Correct errata from previous year (remove duplicates, recode and rename some values) -----
# ******************************************************

errata <- data_clean_data[["errata_gain"]]


#Remove duplicates from previous years
errata_delete <- errata %>% 
  filter(action == 'Delete') %>% 
  select(pindex2, index)


nrow(group_roster) #421
group_roster <- group_roster %>% 
  anti_join(errata_delete, by = c("index", "pindex2"))
nrow(group_roster) #413
#OK removed only 8 


#Rename names
errata_rename <- errata %>% 
  filter(action == 'Rename') %>% 
  select(pindex2, index, errata_value)


group_roster <- merge(group_roster, errata_rename, by = c('index', 'pindex2'), all = T)

contr1 <- group_roster %>% 
  group_by(errata_value) %>% 
  tally()


group_roster <- group_roster %>% 
  mutate(PRO02A = ifelse(is.na(errata_value) == F, errata_value, PRO02A)) %>%
  select(-errata_value)



#Recode
errata_recode <- errata %>% 
  filter(action == 'Recode') %>% 
  select(pindex2, index, errata_value, errata_column) %>% 
  pivot_wider(id_cols = c(index, pindex2),
              names_from = errata_column,
              names_prefix = 'errata_',
              values_from = errata_value)


nrow(errata_recode) #2
nrow(distinct(errata_recode, index, pindex2)) #2 
# OK


group_roster <- merge(group_roster, errata_recode, by = c('index', 'pindex2'), all = T)



group_roster <- group_roster %>% 
  mutate(g_conled = ifelse(is.na(errata_g_conled) == F, errata_g_conled, g_conled),
         g_conled = as.numeric(g_conled),
         PRO03B   = ifelse(is.na(errata_PRO03B) == F, errata_PRO03B, PRO03B),
         PRO03B   = as.numeric(PRO03B)) %>%
  select(-errata_g_conled, -errata_PRO03B)


contr <- group_roster %>% 
  filter(pindex2 == 20230138)



#Change name Netherlands
group_roster <- group_roster %>% 
  mutate(mcountry = ifelse(mcountry == 'Netherlands (Kingdom of the)', 'Netherlands', mcountry))


# ******************************************************
# Step 42: Write and Clean PRO12 roster for challenges and use of recommendations----
# ******************************************************

repeat_data <- final_version_data[["repeat_PRO11_PRO12"]]

ls(repeat_data)

# Ensure `_parent_index` is numeric in `repeat_data`
repeat_data <- repeat_data %>%
  mutate(across(c(`_parent_index`), as.numeric)) %>% 
  mutate(
    `_recommendation` = na_if(`_recommendation`, "NA"),
    
    `_recommendation` = map_chr(`_recommendation`, ~ {
      if (is.na(.x)) {
        NA_character_
      } else if (.x == "إيرس") {
        "IRIS"
      } else {
        translate(
          .x,
          from = "auto",
          to = "en",
          trim_str = TRUE
        )
      }
    })
  )


# Remove entries in "repeat_data" based on del_group_roster

# Remove rows from "group_roster" based on del_group_roster
#Related to PRO11 & PRO12
repeat_data_cleaned <- repeat_data %>%
  filter(!(`_parent_index` %in% del_group_roster$index))

nrow(repeat_data)
nrow(repeat_data_cleaned)


contr <- nrow(repeat_data) - nrow(repeat_data_cleaned)
contr

repeat_data <- repeat_data_cleaned

rm(repeat_data_cleaned)


# Map values from group_roster without duplicating rows in repeat_data
nrow(repeat_data)


contr <- group_roster %>% 
  group_by(PRO04, PRO05) %>% 
  tally()

contr <- repeat_data %>% 
  group_by(`_recommendation`) %>% 
  # group_by(recommendations, `_recommendation`) %>% 
  tally()


summary(group_roster$PRO04)
summary(group_roster$PRO05)


#check how manu entried per index
sub_group_roster <- group_roster %>% 
  mutate(PRO04 = as.Date(PRO04),
         PRO05 = as.Date(PRO05),
         submission__submission_time = as.Date(submission__submission_time)) %>% 
  select(index, morganization, mcountry, PRO04, PRO05, gLOC01, g_conled, egriss_region, recommendations, submission__submission_time) %>% 
  arrange(index, desc(submission__submission_time)) %>% 
  group_by(index) %>% 
  mutate(seq = seq_along(index),
         maxSeq = max(seq)) %>% ungroup()



#will still have to be translated
contr <- sub_group_roster %>% 
  group_by(recommendations) %>% 
  tally()


#To avoid duplicates, only take the last submission 
sub_group_roster <- sub_group_roster %>% 
  filter(seq == 1) %>% 
  select(-seq, maxSeq)


#Duplicates on index

contr <- repeat_data %>% 
  group_by(`_parent_index`) %>% 
  tally()

contr <- repeat_data %>% 
  mutate(`_submission__submission_time` = as.Date(`_submission__submission_time`)) 


summary(contr$`_submission__submission_time`)
#only 2025

contr <- sub_group_roster %>% 
  group_by(recommendations) %>% 
  tally()


nrow(repeat_data)

repeat_data <- repeat_data %>%
  mutate(
    morganization = sub_group_roster$morganization[match(`_parent_index`, sub_group_roster$index)],
    mcountry = sub_group_roster$mcountry[match(`_parent_index`, sub_group_roster$index)],
    gPRO04 = sub_group_roster$PRO04[match(`_parent_index`, sub_group_roster$index)],
    gPRO05 = sub_group_roster$PRO05[match(`_parent_index`, sub_group_roster$index)],
    gLOC01 = sub_group_roster$gLOC01[match(`_parent_index`, sub_group_roster$index)],
    g_conled = sub_group_roster$g_conled[match(`_parent_index`, sub_group_roster$index)],
    region = sub_group_roster$egriss_region[match(`_parent_index`, sub_group_roster$index)],
    recommendations = sub_group_roster$recommendations[match(`_parent_index`, sub_group_roster$index)],
  )

nrow(repeat_data)
#no duplicates
#OK


contr <- repeat_data %>% 
  group_by(recommendations, `_recommendation`) %>% 
  tally()


# ******************************************************
# Step 43: Translate variables ----
# ******************************************************

#Takes 5-10 minutes, so don't always run it...
if (bExport == 1){ 
  ls(group_roster)
  
  vars_to_translate <- c(
    'phase', 'PRO02A', 'PRO03', 'PRO06A', 'PRO07a',
    'PRO08_label', 'PRO13', 'PRO13C_other',
    'PRO14C', 'PRO18', 'PRO20A'
    # 'recommendations'
  )
  
  na_safe_translate <- function(x) {
    if (is.na(x))
      return(NA_character_)
    
    if (x == "")
      return(NA_character_)
    
    if (x == "NA")
      return(NA_character_)
    
    translate(x, from = "auto", to = "en", trim_str = TRUE)
  }
  
  group_roster <- group_roster %>%
    mutate(across(
      all_of(vars_to_translate),
      ~ vapply(.x,
               na_safe_translate,
               FUN.VALUE = character(1),
               USE.NAMES = FALSE)
    ))
} 


# ******************************************************
# Step 44: Recode UPD02 ----
# ******************************************************

contr <- group_roster %>% 
  group_by(year, UPD02) %>% 
  tally()


group_roster <- group_roster %>% 
  mutate(
    UPD02 = case_when(
      UPD02 %in% c("01","1","DESIGN/PLANNING","CONCEPTION/PLANIFICATION","DISEÑO/PLANIFICACIÓN") ~ 1,
      UPD02 %in% c("02","2","IMPLEMENTATION","MISE EN ŒUVRE","IMPLEMENTACIÓN")             ~ 2,
      UPD02 %in% c("03","3","COMPLETED","ACHEVÉ","FINALIZADA")                           ~ 3,
      UPD02 %in% c("06","6","OTHER","AUTRE","OTROS")                                     ~ 6,
      UPD02 %in% c("08","8","DON’T KNOW","NE SAIT PAS","NO SABE")                       ~ 8,
      TRUE                                                                                ~ NA_real_
    ))

# ******************************************************
# Step 45: Recode PRO10 and g_recuse ----
# ******************************************************

group_roster <- group_roster %>%
  mutate(
    PRO10.A = as.numeric(gsub("[^0-9]", "", PRO10.A)),
    PRO10.B = as.numeric(gsub("[^0-9]", "", PRO10.B)),
    PRO10.C = as.numeric(gsub("[^0-9]", "", PRO10.C)),
    PRO10.Z = as.numeric(gsub("[^0-9]", "", PRO10.Z)),
    PRO09 = as.numeric(gsub("[^0-9]", "", PRO09)),
    g_recuse = case_when(
      PRO09 == 1 & PRO10.A == 1 & (is.na(PRO10.B) | PRO10.B != 1) &
        (is.na(PRO10.C) | PRO10.C != 1) ~ "IRRS",
      PRO09 == 1 & PRO10.A != 1 & PRO10.B == 1 & PRO10.C != 1 ~ "IRIS",
      PRO09 == 1 & PRO10.A != 1 & PRO10.B != 1 & PRO10.C == 1 ~ "IROSS",
      PRO09 == 1 & rowSums(cbind(PRO10.A, PRO10.B, PRO10.C), na.rm = TRUE) > 1 ~ "Mixed",
      PRO09 == 1 & PRO10.Z == 1 ~ "Undetermined",
      PRO09 == 1 & PRO10.A != 1 & PRO10.B != 1 & PRO10.C != 1 & PRO10.Z != 1 ~ "Undetermined",
      TRUE ~ "Undetermined"  # Changed NA entries to "Undetermined"
    )
  )

contr <- group_roster %>% 
  group_by(PRO10, g_recuse, PRO10.A, PRO10.B, PRO10.C, PRO10.Z) %>% 
  tally()



# ******************************************************
# Step 46: Standardize PRO04 to year-only format from mixed inputs----
# Handles: MDY strings, Excel serials, plain years, and special values like "9999"
# ******************************************************

summary(group_roster$PRO04)

contr <- group_roster %>% 
  group_by(PRO04, PRO04_year) %>% 
  tally()

group_roster <- group_roster %>%
  mutate(PRO04 = as.Date(PRO04),
         PRO05 = as.Date(PRO05),
         PRO04_year = ifelse(is.na(PRO04_year), year(PRO04), PRO04_year),
         PRO05_year = ifelse(is.na(PRO05_year), year(PRO05), PRO05_year))

contr <- group_roster %>% 
  group_by(PRO04, PRO04_year) %>% 
  tally()


# === Preview before saving ===
cat("\n===== Preview of cleaned PRO04_year values =====\n")
print(head(group_roster$PRO04_year, 20))         # Print first 20 entries
cat("\n===== Frequency table of PRO04_year =====\n")
print(table(group_roster$PRO04_year, useNA = "ifany"))
cat("\nFirst 20 values of PRO05_year:\n")
print(head(group_roster$PRO05_year, 20))

cat("\nFrequency table:\n")
print(table(group_roster$PRO05_year, useNA = "ifany"))

# === Save updated dataset ===
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file, rowNames = FALSE)
}
message("PRO04 standardized to year format in new variable `PRO04_year`.")



# ******************************************************
# Step 47: Write and Clean PRO12 roster for challenges and use of recommendations----
# Rename PRO12 variables systematically
# Rename PRO11 variables systematically
# ******************************************************

#PRO12
# Identify PRO12 columns (ensure they contain "PRO12" somewhere)
pro12_columns <- grep("PRO12", names(repeat_data), value = TRUE)

# Define standard labels starting with "PRO12" and then "PRO12A" to "PRO12I"
standard_labels <- c("PRO12", "PRO12A", "PRO12B", "PRO12C", "PRO12D", 
                     "PRO12E", "PRO12F", "PRO12G", "PRO12H", "PRO12I",
                     "PRO12J", "PRO12K", "PRO12_A")

# Assign names in sequence
if (length(pro12_columns) >= 13) {
  main_pro12_names <- setNames(pro12_columns[1:13], standard_labels)
} else {
  main_pro12_names <- setNames(pro12_columns, standard_labels[seq_along(pro12_columns)])
}

# Identify the "Other (Specify)" and "Don't Know" columns
pro12_other <- grep("OTHER|SPECIFY", pro12_columns, value = TRUE, ignore.case = TRUE)
pro12_dont_know <- grep("DON.TKNOW|DONâ€™TKNOW|DONTKNOW|DON’T KNOW", pro12_columns, value = TRUE, ignore.case = TRUE)

# Assign PRO12X for "Other (Specify)" and PRO12Z for "Don't Know"
if (length(pro12_other) > 0) {
  main_pro12_names["PRO12X"] <- pro12_other[1]
}
if (length(pro12_dont_know) > 0) {
  main_pro12_names["PRO12Z"] <- pro12_dont_know[1]
}



# Rename columns in repeat_data
names(repeat_data)[match(unlist(main_pro12_names), names(repeat_data))] <- names(main_pro12_names)


#PRO11
# Identify PRO11 columns (ensure they contain "PRO11" somewhere)
pro11_columns <- grep("PRO11", names(repeat_data), value = TRUE)

# Define standard labels starting with "PRO12" and then "PRO12A" to "PRO12I"
standard_labels <- c("PRO11", "PRO11A", "PRO11B", "PRO11C", "PRO11D", 
                     "PRO11E", "PRO11F", "PRO11G", "PRO11H", "PRO11X", "PRO11Z", "PRO11_A")

# Assign names in sequence
if (length(pro11_columns) >= 12) {
  main_pro11_names <- setNames(pro11_columns[1:12], standard_labels)
} else {
  main_pro11_names <- setNames(pro11_columns, standard_labels[seq_along(pro11_columns)])
}


# Rename columns in repeat_data
names(repeat_data)[match(unlist(main_pro11_names), names(repeat_data))] <- names(main_pro11_names)



ls(repeat_data)

# Convert PRO12 variables to numeric
pro12_11_numeric_vars <- c("PRO12A", "PRO12B", "PRO12C", "PRO12D", 
                        "PRO12E", "PRO12F", "PRO12G", "PRO12H", "PRO12I",
                        "PRO12X", "PRO12Z",
                        "PRO11A", "PRO11B", "PRO11C", 
                        "PRO11D", "PRO11E", "PRO11F", "PRO11G",
                        "PRO11H", "PRO11X", "PRO11Z")

# Ensure these columns exist in the dataset before converting
existing_pro12_vars <- intersect(pro12_11_numeric_vars, names(repeat_data))

repeat_data <- repeat_data %>%
  mutate(year = year) %>% 
  mutate(across(all_of(existing_pro12_vars), as.numeric))



# Save the cleaned dataset
if (bExport == 1){
  write.xlsx(repeat_data, str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_repeat_PRO11_PRO12.xlsx"))
  write.csv(repeat_data, str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_repeat_PRO11_PRO12.csv"))
}
# Confirm success
message("Updated repeat_data saved successfully with properly renamed PRO12 variables as numeric!")


summary(repeat_data)



# ******************************************************
# Step 48: Put historical PRO11_PRO12 with new PRO11_PRO12 data - 2024 ----
# ******************************************************

#2024 data
pro11_12_histo_path <- "EGRISS GAIN Survey 2024/10 Data/Analysis Ready Files/analysis_ready_repeat_PRO11_PRO12.csv"

pro11_12_histo <- read_csv(pro11_12_histo_path) %>% 
  mutate(year = year-1)


contr <- pro11_12_histo %>% 
  select(year, PRO12A, `_parent_index`)


ls(pro11_12_histo)


# rename columns to match new columns names
#To manually rename as they don't have the same order as in 2025
pro11_columns <- grep("PRO11", names(pro11_12_histo), value = TRUE)

pro11_12_histo <- pro11_12_histo %>% 
  rename(PRO12_A = "PRO12A. Please describe how you applied the EGRISS recommendations in your <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span>:\r\r\nFor illustration, did you refer to them during the design phase to help frame statistical categories, coordinate team training, select appropriate survey questions, or in other ways? \r\r\n\r\r\nProvide specific instances or steps where these guidelines influenced your ${_PRO02A} approach.",
         PRO12_Z = "PRO12. Which elements of the the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations have been/are being used within the design or implementation of <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span>?/DON’T KNOW",
         PRO11   = "PRO11. For what purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?\r\r\nFor instance, did you apply the recommendations while including populations in a census, coordinating and planning national surveys, or using non-traditional data sources? If you select 'Other', please describe",
         PRO11A  = "PRO11. For what purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?\r\r\nFor instance, did you apply the recommendations while including populations in a census, coordinating and planning national surveys, or using non-traditional data sources? If you select 'Other', please describe/INCLUDING REFUGEES, IDPS, OR STATELESS PERSONS IN A POPULATION CENSUS",
         PRO11B  = "PRO11. For what purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?\r\r\nFor instance, did you apply the recommendations while including populations in a census, coordinating and planning national surveys, or using non-traditional data sources? If you select 'Other', please describe/INCLUDING REFUGEES IN A SAMPLE SURVEY OF THE NATIONAL POPULATION, OR RUNNING A STAND-ALONE SURVEY OF REFUGEES",
         PRO11C  = "PRO11. For what purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?\r\r\nFor instance, did you apply the recommendations while including populations in a census, coordinating and planning national surveys, or using non-traditional data sources? If you select 'Other', please describe/INCLUDING IDPS IN A SAMPLE SURVEY OF THE NATIONAL POPULATION, OR RUNNING A STAND-ALONE SURVEY OF IDPS",
         PRO11D  = "PRO11. For what purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?\r\r\nFor instance, did you apply the recommendations while including populations in a census, coordinating and planning national surveys, or using non-traditional data sources? If you select 'Other', please describe/INCLUDING STATELESS PERSONS IN A SAMPLE SURVEY OF THE NATIONAL POPULATION, OR RUNNING A STAND-ALONE SURVEY OF STATELESS PERSONS",
         PRO11E  = "PRO11. For what purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?\r\r\nFor instance, did you apply the recommendations while including populations in a census, coordinating and planning national surveys, or using non-traditional data sources? If you select 'Other', please describe/USING GOVERNMENT ADMINISTRATIVE DATA" ,
         PRO11F  = "PRO11. For what purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?\r\r\nFor instance, did you apply the recommendations while including populations in a census, coordinating and planning national surveys, or using non-traditional data sources? If you select 'Other', please describe/SOURCES OF OPERATIONAL DATA FROM HUMANITARIAN ORGANISATIONS",
         PRO11G  = "PRO11. For what purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?\r\r\nFor instance, did you apply the recommendations while including populations in a census, coordinating and planning national surveys, or using non-traditional data sources? If you select 'Other', please describe/NON-TRADITIONAL DATA SOURCES",
         PRO11H  = "PRO11. For what purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?\r\r\nFor instance, did you apply the recommendations while including populations in a census, coordinating and planning national surveys, or using non-traditional data sources? If you select 'Other', please describe/CO-ORDINATING AND PLANNING REFUGEE, IDP, AND STATELESS STATISTICS IN NATIONAL STATISTICAL SYSTEMS" ,
         PRO11X  = "PRO11. For what purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?\r\r\nFor instance, did you apply the recommendations while including populations in a census, coordinating and planning national surveys, or using non-traditional data sources? If you select 'Other', please describe/OTHER (SPECIFY)",
         PRO11Z  = "PRO11. For what purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?\r\r\nFor instance, did you apply the recommendations while including populations in a census, coordinating and planning national surveys, or using non-traditional data sources? If you select 'Other', please describe/DON’T KNOW",
         PRO11_A = "PRO11A. For what other purpose in <span style='color:#3b71b9; font-weight: bold;'>${_PRO02A}</span> did you use the <span style='color:#3b71b9; font-weight: bold;'>${_recommendation}</span> recommendations for?")

pro11_columns <- grep("PRO11", names(pro11_12_histo), value = TRUE)


#Put the two together
repeat_data_all_dat <- bind_rows(repeat_data, pro11_12_histo)

ls(repeat_data_all_dat)

summary(repeat_data_all_dat)



# ******************************************************
# Step 49: Put historical PRO11_PRO12 with new PRO11_PRO12 data - 2023 ----
# ******************************************************
#2023 data
pro11_12_2023 <- "EGRISS GAIN Survey 2023/07 Data alignment/Alignment - EGRISS_GAIN_Survey_2023_2022_2021_realigmnent_with_PRO3D_clean.xlsx"


pro11_12_2023 <- read_excel(pro11_12_2023, sheet = 'repeat_PRO11_PRO12') %>% 
  mutate(year = year-2)


ls(pro11_12_2023)

contr <- pro11_12_2023 %>% 
  select(year, `PRO12/A`, `_parent_index`)



pro11_12_2023 <- pro11_12_2023 %>% 
  rename(PRO12_A = PRO12A,
         PRO11_A = PRO11A) %>%
  rename_with(~ gsub("/", "", .x))

ls(pro11_12_2023)

#Put the two together
repeat_data_all_dat2 <- bind_rows(repeat_data_all_dat, pro11_12_2023)


ls(repeat_data_all_dat2)
summary(repeat_data_all_dat2)

if (bExport == 1){
  write.csv(repeat_data_all_dat2, str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_repeat_PRO11_PRO12T.csv"))
}



# ******************************************************
# Step 50: Write and Clean GRF Repeat Pledge File (Without Increasing Rows)----
# ******************************************************

# File paths
output_path <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/repeat_pledges_cleaned.xlsx")

# Load datasets
repeat_pledges <- final_version_data[["repeat_pledges"]]

# Ensure `_parent_index` is numeric
repeat_pledges <- repeat_pledges %>%
  mutate(across(c(`_parent_index`), as.numeric))


main_roster_sub <- main_roster %>% 
  mutate(`_submission_time` = as.Date(`_submission_time`)) %>% 
  select(index, mcountry, morganization, LOC01, `_submission_time`) %>% 
  arrange(index, desc(`_submission_time`)) %>% 
  group_by(index) %>% 
  mutate(seq = seq_along(index),
         maxSeq = max(seq)) %>%  ungroup()


#remove duplicates from past years
main_roster_sub <- main_roster_sub %>% 
  filter(seq == 1) %>% 
  select(-seq, -maxSeq)

summary(main_roster_sub)


# Map values from `main_roster` without increasing rows in `repeat_pledges`
nrow(repeat_pledges) #20

repeat_pledges_cleaned <- repeat_pledges %>%
  mutate(
    mcountry = main_roster_sub$mcountry[match(`_parent_index`, main_roster_sub$index)],
    morganization = main_roster_sub$morganization[match(`_parent_index`, main_roster_sub$index)],
    LOC01 = main_roster_sub$LOC01[match(`_parent_index`, main_roster_sub$index)]
  )

nrow(repeat_pledges_cleaned) #20
#OK


# Rename column for pledge status
colnames(repeat_pledges_cleaned)[
  colnames(repeat_pledges_cleaned) == "GRF04. What is the current status of the pledge implementation for pledge: **${pledge_name}?**"
] <- "GRF04"

ls(repeat_pledges_cleaned)

# Save cleaned dataset as CSV
if (bExport == 1){
  write.xlsx(repeat_pledges_cleaned, output_path)
}

# Print success message
cat("The repeat_pledges dataset has been cleaned and saved as 'repeat_pledges_cleaned.xlsx'.\n")



# ******************************************************
# Step 51: Merge & recode UPD02/UPD03A, create UPD25----
# ******************************************************

# Load data
pp <- final_version_data[["repeat_prev_projects"]]

# Shorten any long UPD* headers so we have plain UPD02 & UPD03A in pp
shorten_upd <- function(df) {
  for(pref in c("UPD02","UPD03A")) {
    hits <- grep(paste0("^", pref), names(df), perl = TRUE, value = TRUE)
    if(length(hits)>0) df <- df %>% rename(!!pref := all_of(hits[1]))
  }
  df
}
pp <- shorten_upd(pp)

# 4) Merge in the two new columns
ls(group_roster)
ls(pp)

pp <- pp %>% 
  rename(index   = `_index`) %>% 
  mutate(pindex2 = as.numeric(sprintf("%d%04d", year(`_submission__submission_time`), index)),
         bAll    = 1)

nrow(pp)



group_roster <- group_roster %>%
  mutate(bAll = 1) %>% 
  select(-UPD02, -UPD03A) %>% #from last version
  left_join(
    pp %>% select(index, pindex2, UPD02, UPD03A, bAll),
    by = c("pindex2", "index")
  )


contrmerge <- group_roster %>% 
  group_by(bAll.x, bAll.y) %>% 
  tally()
contrmerge
# bAll.x bAll.y     n
#   1      1        8
#   1     NA      416


# Now recode UPD02 into PRO06 categories, then build UPD25

group_roster <- group_roster %>%
  mutate(
    UPD02 = as.character(UPD02),
    UPD02 = case_when(
      UPD02 %in% c("01","1","DESIGN/PLANNING","CONCEPTION/PLANIFICATION","DISEÑO/PLANIFICACIÓN") ~ 1,
      UPD02 %in% c("02","2","IMPLEMENTATION","MISE EN ŒUVRE","IMPLEMENTACIÓN")             ~ 2,
      UPD02 %in% c("03","3","COMPLETED","ACHEVÉ","FINALIZADA")                           ~ 3,
      UPD02 %in% c("06","6","OTHER","AUTRE","OTROS")                                     ~ 6,
      UPD02 %in% c("08","8","DON’T KNOW","NE SAIT PAS","NO SABE")                       ~ 8,
      TRUE                                                                                ~ NA_real_
    ),
    UPD25 = coalesce(as.numeric(UPD02), PRO06)
  )


# 6) Save back over the CSV
group_roster_file_csv <- "EGRISS GAIN Survey 2025/10 Data/01 Analysis Ready Files//analysis_ready_group_roster.csv"
if (bExport == 1){
  write.xlsx(group_roster, group_roster_file)
  write.csv(group_roster, group_roster_file_csv)
  
}
message("✔ Step 41 complete: UPD02/UPD03A merged, UPD25 created.")



# ******************************************************
# Step 52: GRF File Merging ----
# ******************************************************


# Read the data files
stat_pledges <- read_excel("EGRISS GAIN Survey 2025/10 Data/03 GRF Files External and Internal/20260107_PledgeData.xlsx") %>%
  filter(grepl('Development – Inclusion of Forcibly Displaced', `Multistakeholder pledge`)) %>% 
  mutate(pledge_id = as.character(`Pledge ID`))

pledge_updates <- read_excel("EGRISS GAIN Survey 2025/10 Data/03 GRF Files External and Internal/20260107_PledgeData.xlsx") %>%
  mutate(pledge_id = as.character(`Pledge ID`))

repeat_pledges <- repeat_pledges_cleaned %>%
  mutate(
    Pledge.ID = gsub("GRF_", "GRF-", pledge_name),
    pledge_id = as.character(Pledge.ID)
  )

# Deduplicate to avoid many-to-many joins
nrow(pledge_updates) #52
nrow(distinct(pledge_updates, pledge_id)) #37

pledge_updates <- pledge_updates %>% 
  arrange(pledge_id, `Follow-up Submitted`) %>% 
  group_by(pledge_id) %>% 
  mutate(seq = seq_along(pledge_id),
         maxSeq = max(seq)) %>%  ungroup() %>% 
  filter(seq == maxSeq) %>% 
  select(-seq, -maxSeq)

nrow(pledge_updates) #37
nrow(distinct(pledge_updates, pledge_id)) #37


nrow(repeat_pledges) #20
nrow(distinct(repeat_pledges, pledge_id)) #17

repeat_pledges <- repeat_pledges %>% 
  arrange(pledge_id, `_submission__submission_time`) %>%
  group_by(pledge_id) %>% 
  mutate(seq = seq_along(pledge_id),
         maxSeq = max(seq)) %>%  ungroup() %>% 
  filter(seq == maxSeq) %>% 
  select(-seq, -maxSeq)



# From repeat_pledges, map GRF04 codes → pledge-style text so we can fall back on them
repeat_pledges <- repeat_pledges %>%
  mutate(grf4_pledge = case_when(
    GRF04 == "COMPLETED"        ~ "Fulfilled",
    GRF04 == "DESIGN/PLANNING"  ~ "Planning stage",
    GRF04 == "IMPLEMENTATION"   ~ "In progress",
    TRUE                        ~ NA_character_
  ))


contr <- repeat_pledges %>% 
  group_by(GRF04) %>% 
  tally()


# Prepare lookup tables
pledge_updates_clean <- pledge_updates %>%
  select(pledge_id, `Implementation Stage_FU`) %>%
  rename(Implementation_Stage_FU_updates = `Implementation Stage_FU`)

repeat_pledges_clean <- repeat_pledges %>%
  select(pledge_id, grf4_pledge)

# Identify which IDs are unique to each set
ids_updates_only <- setdiff(pledge_updates_clean$pledge_id, repeat_pledges_clean$pledge_id)
ids_repeat_only  <- setdiff(repeat_pledges_clean$pledge_id, pledge_updates_clean$pledge_id)

# Join everything and compute the final stage + source flag
stat_pledges <- stat_pledges %>%
  left_join(pledge_updates_clean, by = "pledge_id") %>%
  left_join(repeat_pledges_clean, by = "pledge_id") %>%
  mutate(
    # Final stage: prefer raw-text update; else use grf4_pledge from repeat_pledges
    stage_final = coalesce(
      Implementation_Stage_FU_updates,
      grf4_pledge
    ),
    
    # Source flag: 1 if it came from the raw-text update,
    #              2 if it came from the grf4_pledge fallback,
    #             NA if neither
    source_pledge = case_when(
      !is.na(Implementation_Stage_FU_updates)                         ~ 1L,
      is.na(Implementation_Stage_FU_updates) & !is.na(grf4_pledge)    ~ 2L,
      TRUE                                                            ~ NA_integer_
    )
  ) %>%
  select(-Implementation_Stage_FU_updates, -grf4_pledge)

# Country–region mapping
df_country_region <- data_clean_data[["country_name"]] %>% 
  select(mcountry, egriss_region, bAll)


stat_pledges <- stat_pledges %>%
  mutate(bAll = 1) %>% 
  left_join(
    df_country_region,
    by = c("Country - Submitting Entity" = "mcountry")
  )

contrmerge <- stat_pledges %>% 
  group_by(bAll.x, bAll.y) %>% 
  tally()
contrmerge
# bAll.x bAll.y     n
#   1      1       83
#   1     NA       22

contrmerge <- stat_pledges %>% 
  group_by(`Country - Submitting Entity`, bAll.x, bAll.y) %>% 
  tally()
#OK


# Final deduplication and save
stat_pledges   <- stat_pledges   %>% distinct(pledge_id, .keep_all = TRUE)
pledge_updates <- pledge_updates %>% distinct(pledge_id, .keep_all = TRUE)
repeat_pledges <- repeat_pledges %>% distinct(pledge_id, .keep_all = TRUE)

output_dir <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/")

if (bExport == 1){
  write.xlsx(stat_pledges,   file.path(output_dir, "Stat_Inclusion_Pledges_Upd.xlsx"), rowNames = FALSE)
  write.xlsx(pledge_updates, file.path(output_dir, "Pledge_Updates_2024.xlsx"),                   rowNames = FALSE)
  write.xlsx(repeat_pledges, file.path(output_dir, "repeat_pledges_cleaned.xlsx"),                rowNames = FALSE)
}


# ******************************************************
# Step 53: Remove specified variables from main roster and save as "main.xlsx"----
# ******************************************************

# Define file paths:
analysis_ready_main_roster_file <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/analysis_ready_main_roster.xlsx")
output_file <- str_c(egr_yr, "/10 Data/01 Analysis Ready Files/main.xlsx")

# Create a vector of columns to remove:
cols_to_remove <- c(
  "LOC06_2", "Country_UNHCR", "prev_projects", "PRO02Note", "END01",
  "X_id", "X_uuid", "X_submission_time", "X_validation_status", "X_notes",
  "X_status", "X_submitted_by", "X__version__", "X_tags", "FOL01",
  "FOL02A", "FOL02B", "FOL02C", "FOL02D", "FOL03", "FOL04", "ACT06.A",
  "count_ACT04", "ACT04", "count_pledges", "GRF03", "pledgesavailable",
  "organizationGRF", "FPR02", "count_FPR02", "count_PRO02A", "PRO02",
  "NameFOC04_1", "NameFOC04_2", "NameFOC04_3", "NameFOC04_4", "NameFOC04_5",
  "NameFOC04_6", "NameFOC04_7", "NameFOC04_8", "NameFOC04_9", "NameFOC04_10",
  "UPD01", "count_prev_projects", "LOC01B", "LOC02", "LOC03", "LOC04",
  "LOC04A", "LOC05", "LOC06", "LOC06_UNCT", "LOC06A", "UNHCR_Level",
  "Bureau", "LOC06C", "LOC06_3", "LOC06_2_other", "LOC06_2_label",
  "LOC06_2_label2", "LOC06_label", "LOC06_label2", "LOC01B_label",
  "LOC06_4", "organization", "FOC01A", "FOC01B", "FOC01C", "NameFOC01",
  "FOC02", "FOC03A", "start", "end", "today", "logo"
)

# Check if the main roster file exists, then remove listed columns and save as main.xlsx
if (file.exists(analysis_ready_main_roster_file)) {
  # main_roster <- read.xlsx(analysis_ready_main_roster_file)
  
  main_roster_sub <- main_roster %>%
    select(-any_of(cols_to_remove))
  
  if (bExport == 1){
    write.xlsx(main_roster_sub, output_file, rowNames = FALSE)
  }
  
  message(paste("Successfully removed requested columns. File saved as:", output_file))
} else {
  stop("The file 'analysis_ready_main_roster.xlsx' does not exist in the specified directory.")
}

# ******************************************************
# Step 54: Create final version of group_roster as `pro.xlsx`----
# Removes administrative, unused, and auxiliary fields
# ******************************************************

vars_to_remove <- c(
  "submission__tags",
  "submission___version__",
  "submission__status",
  "submission__submitted_by",
  "submission__validation_status",
  "submission__submission_time",
  "submission__uuid",
  "submission__id",
  "year/ryear",
  "PRO14A",
  "PRO18",
  "PRO13",
  "PRO16",
  "parent_table_name",
  "index",
  "parent_index",
  "PRO22AA",
  "PRO22",
  "PRO20A",
  "PRO13C_other",
  "PRO13C",
  "recommendations",
  "count_recommendations",
  "PRO10",
  "PRO08a",
  "PRO08",
  "PRO07a",
  "PRO07",
  "phase",
  "PRO06A",
  "project",
  "PRO02A",
  "bAll.x",
  "bAll.y",
  "X_blanc"
)

# Keep only variables not in the removal list
group_roster_final <- group_roster %>%
  select(-any_of(vars_to_remove))

# Save final version as pro.xlsx
if (bExport == 1){
  write.xlsx(group_roster_final, str_c(egr_yr,"/10 Data/01 Analysis Ready Files/pro.xlsx"), rowNames = FALSE)
}

message("Final cleaned group roster saved as `pro.xlsx` with selected variables removed.")


# ******************************************************
# Step 55: Backup Analysis Ready Files with a Timestamp----
# ******************************************************

if (bExport == 1){
  # Define the base directory for analysis-ready files
  analysis_ready_directory <- normalizePath(
    file.path(egr_yr, "10 Data/01 Analysis Ready Files"),
    mustWork = TRUE
  )
  
  # Define the backup folder with timestamp
  timestamp <- format(Sys.time(), "%Y-%m-%d_%H-%M-%S")  # Generate timestamp
  backup_directory <- file.path(
    analysis_ready_directory,
    paste0("Backup_", timestamp)
  )
  
  # Ensure the backup directory exists
  dir_create(backup_directory, recurse = TRUE)
  message("✅ Backup folder created: ", backup_directory)
  
  # List all analysis-ready files (excluding previous backups)
  analysis_files <- dir_ls(analysis_ready_directory, type = "file")
  # analysis_files <- analysis_files[!str_detect(analysis_files, "Backup_")]
  analysis_files <- analysis_files[
    !str_detect(analysis_files, "Backup_") &
      !str_detect(analysis_files, "\\.docx$")
  ]
  
  # Copy each file to the backup folder
  file_copy(analysis_files, backup_directory, overwrite = TRUE)
  
  # List and print the backed-up files
  backup_files <- dir_ls(backup_directory, type = "file")
  message("📂 Backup Completed. Files in the backup folder:")
  print(backup_files)
}



# ******************************************************
# Step 56: Frequency tables ----
# ******************************************************

# group_roster 
cols <- c(ls(group_roster))


summary_group_roster <- map_dfr(cols, function(i) {
  
  variable_data <- group_roster[[i]]
  var_label    <- attr(variable_data, "label")
  value_labels <- attr(variable_data, "labels")
  
  n_unique  <- n_distinct(variable_data, na.rm = TRUE)
  is_date   <- inherits(variable_data, c("Date", "POSIXct", "POSIXt"))
  is_numeric <- is.numeric(variable_data)
  
  var_type <- paste(class(variable_data), collapse = ", ")
  
  group_roster %>%
    group_by(.data[[i]]) %>%
    tally(name = "n") %>%
    ungroup() %>%
    mutate(
      Tot       = sum(n),
      Perc      = round(n / Tot * 100, 1),
      
      Variable  = i,
      Type      = var_type,
      
      Value     = as.character(.data[[i]]),
      
      Label = if (!is.null(var_label)) var_label else NA_character_,
      
      ValueLabel = if (!is.null(value_labels)) {
        vl <- setNames(names(value_labels), as.character(value_labels))
        vl[Value]
      } else {
        NA_character_
      }
    ) %>%
    select(Variable, Type, Value, n, Tot, Perc)
})




# main_roster
cols <- c(ls(main_roster))

# cols <- cols[cols %not in% c("Respondent_Serial", "Weightvar_XW", "Weightvar_RN")]

summary_main_roster <- map_dfr(cols, function(i) {
  
  variable_data <- main_roster[[i]]
  var_label    <- attr(variable_data, "label")
  value_labels <- attr(variable_data, "labels")
  
  n_unique   <- n_distinct(variable_data, na.rm = TRUE)
  is_date    <- inherits(variable_data, c("Date", "POSIXct", "POSIXt"))
  is_numeric <- is.numeric(variable_data)
  
  var_type <- paste(class(variable_data), collapse = ", ")
  
  main_roster %>%
    group_by(.data[[i]]) %>%
    tally(name = "n") %>%
    ungroup() %>%
    mutate(
      Tot  = sum(n),
      Perc = round(n / Tot * 100, 1),
      
      Variable = i,
      Type     = var_type,
      
      Value = as.character(.data[[i]]),
      
      Label = if (!is.null(var_label)) var_label else NA_character_,
      
      ValueLabel = if (!is.null(value_labels)) {
        vl <- setNames(names(value_labels), as.character(value_labels))
        vl[Value]
      } else {
        NA_character_
      }
    ) %>%
    select(Variable, Type, Value, n, Tot, Perc)
  
})







if (bExport == 3) {
  DATA <- list(summary_main_roster, summary_group_roster)
  sheet_names <- c("main_roster", "group_roster")
  
  # Create a new workbook
  wb <- createWorkbook()
  
  # Add each dataframe to the workbook as a sheet
  for (i in seq_along(DATA)) {
    addWorksheet(wb, sheetName = sheet_names[i])
    writeData(wb, sheet = sheet_names[i], DATA[[i]])
  }
  
  # Save the workbook
  output_file <- str_c('EGRISS GAIN Survey 2025/06 Data Cleaning/03 Frequencies', '/summary_roster_', today(), '.xlsx')
  saveWorkbook(wb, file = output_file, overwrite = TRUE)
}

