#' This script produces the LMO files for WorkBC. Requires inputs:
fyod <- 2023 # first year of data: need to reset each year
#' "employment.csv" (4castviewer)
#' "ftpt2125NOCp1.csv" (run the SAS scripts... will need to be changed 2027ish)
#' "ftpt2125NOCp2.csv" (run the SAS scripts... will need to be changed 2027ish)
#' "industry_characteristics_2023.xlsx" (???)
#' "job_openings.csv"  (4castviewer)
#' "LMO 2023E HOO BC and Regions 2023-08-23.xlsx" (Feng)
#' "Occupational Characteristics w skills interests wages" (created by file add_skills_interests_wages.R)
#' "Occupational Characteristics based on LMO 2023E 2023-08-30.xlsx" (Feng) NOT CALLED BY THIS SCRIPT, BUT USED TO CREATE ^
#' "Occupational interest by NOC2021 occupation.xlsx" (Amy) NOT CALLED BY THIS SCRIPT, BUT USED TO CREATE ^
#' "Top skills by NOC2021 occupations.xlsx" (Amy) NOT CALLED BY THIS SCRIPT, BUT USED TO CREATE ^
#' "WorkBC_Career_Trek_2023-2033_CEU EDIT.xlsx" (WorkBC)
#'
#' To Run: source me, and then you will have to verify the hoo geography is right...
#'
#' some code to correct the senior vs seniors management can probably be removed...

# "constants"------------------------
fyfn <- fyod + 5
tyfn <- fyod + 10
# libraries---------------------------
library(tidyverse)
library(here)
library(vroom)
library(readxl)
library(rvest)
library(openxlsx)
library(XLConnect)
# deal with naming conflicts between packages
library(conflicted)
conflicts_prefer(vroom::cols)
conflicts_prefer(vroom::col_double)
conflicts_prefer(vroom::col_character)
conflicts_prefer(dplyr::filter)
conflicts_prefer(XLConnect::loadWorkbook)
conflicts_prefer(XLConnect::saveWorkbook)
# functions------------------
fill_skills <- function(tbbl){
  if(nrow(tbbl)==1){ #if there is no skill data...
  tbbl <- tbbl|>
    slice_sample(n=length(skill_names), replace=TRUE) #replicate columns to fit 35ish skill names
  tbbl$`Skills & Competencies` <- skill_names #add in the skill names
  }
  return(tbbl)
}

fill_interests <- function(tbbl){
  if(nrow(tbbl)==1){ #if there is no data...
    tbbl <- tbbl|>
      slice_sample(n=3, replace=TRUE) #replicate columns to fit 3 interests
    tbbl$Options <- c("Primary","Secondary", "Tertiary") #add in the option names
  }
  return(tbbl)
}



make_na <- function(vec){
  vec <- tibble("vec"=vec)|>
    mutate(vec=if_else(vec==0, NA_character_, vec),
           vec=as.numeric(vec))
  return(vec$vec)
}
get_levels <- function(tbbl) {
  #' take a tbbl containing ONLY columns name and value,
  #' and returns a tbbl containing current employment, employment in 5 years, employment in 10 years.
  employment_current <- tbbl$value[tbbl$name == fyod]
  employment_five <- tbbl$value[tbbl$name == fyfn]
  employment_ten <- tbbl$value[tbbl$name == tyfn]
  tibble(
    employment_current = employment_current,
    employment_five = employment_five,
    employment_ten = employment_ten
  )
}
get_emp_levels <- function(tbbl) {
  #' takes a tbbl containing columns name, value and Variable and returns a tbbl
  #' containing current employment, employment in 5 years, employment in 10 years.
  employment_current <- tbbl$value[tbbl$Variable == "Employment" & tbbl$name == fyod]
  employment_five <- tbbl$value[tbbl$Variable == "Employment" & tbbl$name == fyfn]
  employment_ten <- tbbl$value[tbbl$Variable == "Employment" & tbbl$name == tyfn]
  tibble(
    employment_current = employment_current,
    employment_five = employment_five,
    employment_ten = employment_ten
  )
}
make_wide <- function(tbbl) {
  # this function take a n by k tbbl, and returns a 1 by nk dataframe
  tbbl |>
    mutate(row = row_number()) |>
    tidyr::pivot_wider(
      names_from = row,
      values_from = c(`Industry (aggregate)`, job_openings, `%`),
      names_vary = "slowest"
    )
}
get_cagrs <- function(tbbl) {
  #' takes tbbl with columns name, value, Variable and returns a tbbl
  #' with the CAGR % for the first five years and second five years
  now <- tbbl$value[tbbl$Variable == "Employment" & tbbl$name == fyod]
  fyfn <- tbbl$value[tbbl$Variable == "Employment" & tbbl$name == fyfn]
  tyfn <- tbbl$value[tbbl$Variable == "Employment" & tbbl$name == tyfn]
  ffy_cagr <- 100 * round((fyfn / now)^.2 - 1, 3)
  sfy_cagr <- 100 * round((tyfn / fyfn)^.2 - 1, 3)
  tibble(ffy_cagr = ffy_cagr, sfy_cagr = sfy_cagr)
}
get_10_cagr <- function(tbbl) {
  #' takes tbbl with columns name, value, Variable and returns the
  #' CAGR % for the ten year period.
  now <- tbbl$value[tbbl$Variable == "Employment" & tbbl$name == fyod]
  tyfn <- tbbl$value[tbbl$Variable == "Employment" & tbbl$name == tyfn]
  cagr <- 100 * round((tyfn / now)^.1 - 1, 3)
  cagr
}
get_jos <- function(tbbl) {
  #' takes tbbl with columns name, value, Variable and returns a tbbl with
  #' the sum of job openings for the first five year and second five year period
  ffy <- sum(tbbl$value[tbbl$Variable == "Job Openings" & tbbl$name %in% (fyod + 1):fyfn])
  sfy <- sum(tbbl$value[tbbl$Variable == "Job Openings" & tbbl$name %in% (fyfn + 1):tyfn])
  tibble(ffy = ffy, sfy = sfy)
}
get_breakdown <- function(tbbl) {
  #' Takes a tbbl with columns name, value, Variable, and returns a tbbl containing
  #' job openings, replacement %, replacement demand, expansion % and expansion demand.
  jo <- sum(tbbl$value[tbbl$Variable == "Job Openings" & tbbl$name %in% (fyod + 1):tyfn])
  rep <- sum(tbbl$value[tbbl$Variable == "Replacement Demand" & tbbl$name %in% (fyod + 1):tyfn])
  exp <- sum(tbbl$value[tbbl$Variable == "Expansion Demand" & tbbl$name %in% (fyod + 1):tyfn])
  exp_p <- round(exp / jo, 3)
  exp_p <- case_when(
    exp_p > 1 ~ 1,
    exp_p < 0 ~ 0,
    TRUE ~ exp_p
  )
  rep_p <- round(rep / jo, 3)
  rep_p <- case_when(
    rep_p > 1 ~ 1,
    rep_p < 0 ~ 0,
    TRUE ~ rep_p
  )
  tibble(jo = jo, rep_p = 100 * rep_p, rep = rep, exp_p = 100 * exp_p, exp = exp)
}
get_current <- function(tbbl) {
  #' Takes a tbbl with columns name, value, Variable, and returns the
  #' current level of employment.
  tbbl$value[tbbl$Variable == "Employment" & tbbl$name == fyod]
}
get_10_jo <- function(tbbl) {
  #' Takes a tbbl with columns name, value, Variable, and returns the sum of
  #' job openings over the 10 year forecast.
  sum(tbbl$value[tbbl$Variable == "Job Openings" & tbbl$name %in% (fyod + 1):tyfn])
}
# READ IN THE DATA------------------------
# list of NOCs that have videos--------------------------------
career_trek_have_videos <- read_excel(here("data", "WorkBC_Career_Trek_2023-2033_CEU EDIT.xlsx")) |>
  select(-contains("Updates"), -contains("employment"), -contains("openings")) |>
  rename(NOC = `NOC 2021...2`) |>
  select(-contains("..."))
# full time part time---------------------
ftpt <- vroom(
  here(
    "data",
    list.files(here("data"),
      pattern = "ftpt"
    )
  ),
  col_types = cols(
    SYEAR = col_double(),
    FTPT = col_character(),
    NOC_5 = col_character(),
    `_COUNT_` = col_double()
  )
) |>
  na.omit() |>
  pivot_wider(id_cols = c(SYEAR, NOC_5), names_from = FTPT, values_from = `_COUNT_`) |>
  group_by(NOC_5) |>
  summarize(
    `Full-time` = sum(`Full-time`, na.rm = TRUE),
    `Part-time` = sum(`Part-time`, na.rm = TRUE)
  ) |>
  mutate(
    prop_ft = `Full-time` / (`Full-time` + `Part-time`),
    `Part-time/full-time` = if_else(prop_ft < .70,
      "Higher Chance of Part-Time",
      "Higher Chance of Full-Time"
    )
  ) |>
  select(NOC_2021 = NOC_5, `Part-time/full-time`) |>
  mutate(NOC_2021 = paste0("#", NOC_2021))
senior_managers <- tibble(NOC_2021 = "#00018", `Part-time/full-time` = "Higher Chance of Full-Time") # based on components
ftpt <- bind_rows(ftpt, senior_managers)
# get teer descriptions from web------------------------------------
teer_description <- read_html("https://noc.esdc.gc.ca/Training/TeerCategory") |>
  html_table()
teer_description <- teer_description[[1]]
teer_description <- teer_description |>
  mutate(TEER = as.character(`TEER category`)) |>
  select(-`TEER category`, TEER_description = starts_with("Nature"))
# occupation characteristics--------------------------------------
occ_char <- read_excel(
  here("data", "Occupational Characteristics w skills interests wages.xlsx"),
  skip = 0,
  na = "x"
)|>
  mutate(`2021 Census Median Employment Income (Employed)` = as.numeric(`2021 Census Median Employment Income (Employed)`),
         `Calculated Median Annual Salary \n2023`=round(as.numeric(`Calculated Median Annual Salary \n2023`))
         )
# industry characteristics---------------------------------------------------
ind_char <- read_excel(here(
  "data",
  list.files(here("data"),
    pattern = "industry"
  )
)) |>
  transmute(
    `Industry (sub-industry)` = lmo_detailed_industry,
    `Industry (aggregate)` = str_to_title(str_replace_all(aggregate_industry, "_", " "))
  ) |>
  distinct()
# job openings----------------------------------------------
jo <- vroom(here("data", "job_openings.csv"), skip = 3) |>
  janitor::remove_empty() |>
  filter(
    Variable == "Job Openings",
    !`Geographic Area` %in% c("North", "South East")
  ) |>
  pivot_longer(cols = starts_with("2")) |>
  group_by(NOC, Description, Industry, `Geographic Area`) |>
  summarize(jo = sum(value)) |>
  rename(
    NOC_2021 = NOC,
    NOC_2021_Description = Description
  ) |>
  mutate(NOC_2021_Description=if_else(NOC_2021_Description=="Seniors managers - public and private sector",
                                      "Senior managers - public and private sector",
                                      NOC_2021_Description
                                      ))
# job openings and components-------------------------
jore <- vroom(here("data", "job_openings.csv"), skip = 3) |>
  janitor::remove_empty() |>
  filter(
    Variable %in% c("Job Openings", "Expansion Demand", "Replacement Demand"),
    !`Geographic Area` %in% c("North", "South East"),
    Industry == "All industries"
  ) |>
  pivot_longer(cols = starts_with("2")) |>
  select(-Industry)|>
  mutate(Description=if_else(Description=="Seniors managers - public and private sector",
                                      "Senior managers - public and private sector",
                                      Description
  ))

# employment to add to job openings and components
emp_to_add_to_jore <- vroom(here("data", "employment.csv"), skip = 3) |>
  janitor::remove_empty() |>
  filter(
    !`Geographic Area` %in% c("North", "South East"),
    Industry == "All industries"
  ) |>
  pivot_longer(cols = starts_with("2")) |>
  select(-Industry)|>
  mutate(Description=if_else(Description=="Seniors managers - public and private sector",
                             "Senior managers - public and private sector",
                             Description
  ))

long <- bind_rows(jore, emp_to_add_to_jore) |>
  group_by(NOC, Description, `Geographic Area`) |>
  nest()

# high opportunity occupations------------------------------------
hoo_sheets <- head(excel_sheets(here("data", "LMO 2023E HOO BC and Regions 2023-08-23.xlsx")), -1) |>
  sort()
jo_regions <- unique(jo$`Geographic Area`) |>
  sort()

hoo_geography <- tibble(
  name = hoo_sheets,
  Geography = jo_regions
)
View(hoo_geography)
continue <- readline("Does the hoo geography shown in the tab above match ? (answer y or n)")
assertthat::assert_that(continue == "y", msg = "You need to manually need to fix hoo_geography")
hoo <- map(set_names(hoo_sheets),
  read_excel,
  path = here("data", "LMO 2023E HOO BC and Regions 2023-08-23.xlsx"),
  range = "A5:C500" # can't possibly be more than 500 hoos?
) |>
  enframe() |>
  full_join(hoo_geography) |>
  select(-name) |>
  unnest(value) |>
  na.omit()

# employment----------------------------------

emp <- vroom(here("data", "employment.csv"), skip = 3) |>
  janitor::remove_empty() |>
  filter(!`Geographic Area` %in% c("North", "South East")) |>
  pivot_longer(cols = starts_with("2"))|>
  mutate(Description=if_else(Description=="Seniors managers - public and private sector",
                             "Senior managers - public and private sector",
                             Description
  ))

# PROCESS THE DATA-----------------------------
# All Occupations' TEERs.xlsx------------------------------

occ_char |>
  select(NOC_2021 = NOC, NOC_2021_Description = Description) |>
  mutate(TEER = str_sub(NOC_2021, 3, 3)) |>
  inner_join(teer_description) |>
  write.xlsx(here(
    "out",
    paste0(
      "All_Occupations'_TEERs_",
      fyod,
      ".xlsx"
    )
  ))

# CareerDiscoveryQuizzesJobOpenings- master.csv--------------------------------

cdqjom <- occ_char |>
  select(
    NOC_2021 = NOC,
    NOC_2021_Description = Description,
    jo = contains("Job Openings") & !contains("Ave"),
    TEER,
    `Salary (calculated median salary)` = starts_with("Calculated Median Annual"),#might be the wrong one
  ) |>
  mutate(
    jo = round(jo, -1),
    Link = paste0(
      "https://www.workbc.ca/Jobs-Careers/Explore-Careers/Browse-Career-Profile/",
      str_sub(NOC_2021, 2)
    ),
    JobBank2 = paste0(
      "https://www.workbc.ca/jobs-careers/find-jobs/jobs.aspx?searchNOC=",
      str_sub(NOC_2021, 2)
    ),
    TEER = as.character(TEER)
  ) |>
  inner_join(teer_description) |>
  relocate(TEER_description, .after = TEER) |>
  rename("Job Openings {fyod}-{tyfn}" := jo)

cdqjom %>%
  write.xlsx(here(
    "out",
    paste0(
      "Career_Discovery_Quizzes_Job_Openings_master_",
      fyod,
      ".xlsx"
    )
  ))

# Career Search Tool Job Openings.xlsx-----------------------------

temp <- cdqjom %>%
  select(-contains("Job Openings")) |>
  full_join(jo) |>
  mutate(
    `Industry (sub-industry)` = str_to_lower(str_replace_all(Industry, " ", "_")),
    jo = round(jo, -1)
  ) |>
  full_join(ind_char) |>
  select(NOC_2021,
    NOC_2021_Description,
    `Industry (sub-industry)` = Industry,
    Region = `Geographic Area`,
    "Job Openings {fyod}-{tyfn}" := jo,
    TEER,
    TEER_description,
    `Salary (calculated median salary)`,
    Link,
    JobBank2,
    `Industry (aggregate)`
  ) |>
  mutate(`Industry (aggregate)` = if_else(is.na(`Industry (aggregate)`),
    "All industries",
    `Industry (aggregate)`
  )) |>
  filter(NOC_2021 != "#T") |>
  left_join(ftpt)

# map(temp, ~sum(is.na(.)))

temp |>
  write.xlsx(here(
    "out",
    paste0(
      "Career_Search_Tool_Job_Openings_",
      fyod,
      ".xlsx"
    )
  ))

# career_search_tool_occupation_groups_manual_update.xlsx-----------------

cstogmu_hoo <- hoo |>
  select(NOC = "#NOC (2021)", Region = Geography) |>
  mutate(`Occupational category` = "High opportunity occupations")

all_regions <- unique(cstogmu_hoo$Region) #to create the redundant information WorkBC wants...

cstogmu_stem <- occ_char |>
  filter(`Occ Group: STEM` == "STEM") |>
  select(NOC) |>
  mutate(`Occupational category` = "STEM")

cstogmu_trades <- occ_char |>
  filter(`Occ Group: Trades` == "Trades") |>
  select(NOC,`Occupational category` = `Occ Group: Construction Trades`)

cstogmu_care <- occ_char |>
  filter(`Occ Group: Deloitte Care Economy` == "Care Economy") |>
  select(NOC) |>
  mutate(`Occupational category` = "Care Economy")

cstogmu_manage <- occ_char |>
  select(NOC, TEER) |>
  mutate(`Occupational category` = if_else(TEER == 0,
      "Management occupations",
      "Non-management occupations"
    )
  ) |>
  select(-TEER)

cstogmu_all <- occ_char |>
  select(NOC) |>
  mutate(`Occupational category` = "All")

no_regions <- bind_rows(
  cstogmu_all,
  cstogmu_manage,
  cstogmu_care,
  cstogmu_trades,
  cstogmu_stem
)

crossing(no_regions, Region=all_regions)|> #adds in the redundant region info
  bind_rows(cstogmu_hoo)|>
  write.xlsx(here(
    "out",
    paste0(
      "career_search_tool_occupation_groups_manual_update_",
      fyod,
      ".xlsx"
    )
  ))

# HOO BC and Region for new tool.xlsx-----------------------------

occ_char |>
  select(-contains("Job Openings")) |>
  right_join(hoo, by = c("NOC" = "#NOC (2021)")) |>
  rename(jo = contains("Job Openings") & !contains("Ave")) |>
  mutate(jo = round(jo, -1)) |>
  select(
    `Occupation Title` = Description,
    "Job Openings {fyod}-{tyfn}" := jo,
    "Wage Rate Low {fyod}" := contains("Wage Rate Low"),
    "Wage Rate Median {fyod}" := contains("Wage Rate Median"),
    "Wage Rate High {fyod}" := contains("Wage Rate High"),
    `Median Annual Salary` = starts_with("Calculated Median Annual"),#might be the wrong one
    Interests,
    `Skills and Competencies (Top 3 together)` = `Skills: Top 3`,
    `First` = Skill1,
    `Second` = Skill2,
    `Third` = Skill3,
    `#NOC` = NOC,
    Geography
  ) |>
  mutate(NOC = str_sub(`#NOC`, 2, -1), .before = `#NOC`) |>
  mutate(TEER = str_sub(NOC, 2, 2), .before = Geography) |>
  write.xlsx(here(
    "out",
    paste0(
      "HOO_BC_and_Region_for_new_tool_",
      fyod,
      ".xlsx"
    )
  ))

# HOO List.xlsx-----------------------------

hoo |>
  filter(Geography == "British Columbia") |>
  select(-contains("Job Openings")) |>
  left_join(occ_char, by = c("#NOC (2021)" = "NOC")) |>
  rename(jo = contains("Job Openings") & !contains("Ave")) |>
  mutate(jo = round(jo, -1)) |>
  select(
    `Occupation Title` = Description,
    "Job Openings {fyod}-{tyfn}" := jo,
    "Wage Rate Low {fyod}" := contains("Wage Rate Low"),
    "Wage Rate Median {fyod}" := contains("Wage Rate Median"),
    "Wage Rate High {fyod}" := contains("Wage Rate High"),
    `Median Annual Salary` = starts_with("Calculated Median Annual"),#might be the wrong one
    NOC = `#NOC (2021)`,
    `Occupational Interest` = Interests,
    `Skills and Compentencies (Top 3)` = `Skills: Top 3`
  ) |>
  mutate(
    NOC = str_sub(NOC, 2, -1),
    TEER = str_sub(NOC, 2, 2)
  ) |>
  write.xlsx(here(
    "out",
    paste0(
      "HOO_List_",
      fyod,
      ".xlsx"
    )
  ))

# Job Openings by Industry_LMO.xlsx------------------------

career_profiles <- jo |>
  filter(`Geographic Area` == "British Columbia") |>
  mutate(Industry = str_to_lower(str_replace_all(Industry, " ", "_"))) |>
  full_join(ind_char, by = c("Industry" = "Industry (sub-industry)")) |>
  filter(!is.na(`Industry (aggregate)`)) |>
  group_by(NOC_2021, NOC_2021_Description, `Industry (aggregate)`) |>
  summarize(job_openings = sum(jo)) |>
  slice_max(job_openings, n = 5, with_ties = FALSE) |>
  mutate(
    `%` = 100 * round(job_openings / sum(job_openings), 2),
    job_openings = round(job_openings, -1)
  ) |>
  group_by(NOC_2021, NOC_2021_Description) |>
  nest() |>
  mutate(data = map(data, make_wide)) |>
  unnest(data) |>
  relocate(`%_1`, .before = `job_openings_1`) |>
  relocate(`%_2`, .before = `job_openings_2`) |>
  relocate(`%_3`, .before = `job_openings_3`) |>
  relocate(`%_4`, .before = `job_openings_4`) |>
  relocate(`%_5`, .before = `job_openings_5`)

`Job Openings by industry` <- jo |>
  filter(
    `Geographic Area` == "British Columbia",
    NOC_2021 == "#T"
  ) |>
  mutate(Industry = str_to_lower(str_replace_all(Industry, " ", "_"))) |>
  full_join(ind_char, by = c("Industry" = "Industry (sub-industry)")) |>
  na.omit() |>
  group_by(`Industry (aggregate)`) |>
  summarize(jo = round(sum(jo), -1)) |>
  rename("Job Openings {fyod}-{tyfn}" := jo)

wb <- loadWorkbook(here("new_templates", "Job Openings by Industry_2016 Census_2023 LMO.xlsx"))
writeWorksheet(wb, career_profiles, "Career Profiles", 5, 1, header = FALSE)
writeWorksheet(wb, `Job Openings by industry`, "Job Openings by industry", 4, 1, header = FALSE)
saveWorkbook(wb, here(
  "out",
  paste0(
    "Job_Openings_by_Industry_LMO_",
    fyod,
    ".xlsx"
  )
))

# LMO_Forecasted Employment Growth_2023-24.xlsx------------------------------

jo_by_region <- jo |>
  filter(
    NOC_2021 == "#T",
    Industry == "All industries"
  ) |>
  mutate(jo = round(jo, -1)) |>
  ungroup() |>
  select(`Geographic Area`, ten_sum_jo = jo)

emp |>
  filter(
    NOC == "#T",
    Industry == "All industries"
  ) |>
  select(-NOC, -Description, -Variable, -Industry) |>
  group_by(`Geographic Area`) |>
  nest() |>
  mutate(levels = map(data, get_levels)) |>
  unnest(levels) |>
  mutate(
    cagr_first_five = 100 * round((employment_five / employment_current)^(.2) - 1, 3),
    cagr_second_five = 100 * round((employment_ten / employment_five)^(.2) - 1, 3),
    cagr_ten_year = 100 * round((employment_ten / employment_current)^(.1) - 1, 3)
  ) |>
  mutate(across(contains("employment"), ~ round(.x, -1))) |>
  select(contains("Area"), contains("cagr"), contains("employment")) |>
  full_join(jo_by_region) |>
  write.xlsx(here(
    "out",
    paste0(
      "LMO_Forecasted_Employment_Growth_",
      fyod,
      ".xlsx"
    )
  ))

# top_5_careers_by_aggregate_industry.xlsx------------------------------

emp |>
  filter(
    NOC != "#T",
    `Geographic Area` == "British Columbia",
    name == fyod
  ) |>
  mutate(Industry = str_to_lower(str_replace_all(Industry, " ", "_"))) |>
  full_join(ind_char, by = c("Industry" = "Industry (sub-industry)")) |>
  group_by(NOC, Description, `Industry (aggregate)`) |>
  summarize(emp = sum(value)) |>
  group_by(`Industry (aggregate)`, .add = FALSE) |>
  slice_max(emp, n = 5, with_ties = FALSE) |>
  select(`Aggregate Industry` = `Industry (aggregate)`, NOC, Occupation = Description, emp) |>
  mutate("Employment {fyod}" := round(emp, -1)) |>
  select(-emp) |>
  write.xlsx(here(
    "out",
    paste0(
      "top_5_careers_by_aggregate_industry_",
      fyod,
      ".xlsx"
    )
  ))

# top_10_careers_by_aggregate_industry.xlsx----------------------------

jo |>
  filter(
    NOC_2021 != "#T",
    `Geographic Area` == "British Columbia"
  ) |>
  mutate(Industry = str_to_lower(str_replace_all(Industry, " ", "_"))) |>
  full_join(ind_char, by = c("Industry" = "Industry (sub-industry)")) |>
  group_by(NOC_2021, NOC_2021_Description, `Industry (aggregate)`) |>
  summarize(jo = sum(jo)) |>
  group_by(`Industry (aggregate)`, .add = FALSE) |>
  slice_max(jo, n = 10, with_ties = FALSE) |>
  select(`Aggregate Industry` = `Industry (aggregate)`, NOC = NOC_2021, Occupation = NOC_2021_Description, jo) |>
  mutate("Job Openings {fyod}-{tyfn}" := round(jo, -1)) |>
  select(-jo) |>
  na.omit() |>
  write.xlsx(here(
    "out",
    paste0(
      "top_10_careers_by_aggregate_industry_",
      fyod,
      ".xlsx"
    )
  ))

# WorkBC_Career_Profile_Data.xlsx----------------------------------------------------
#' 1. Suppression rule:
#'   If the current year's employment is below 20,
#'   then all numeric variables should be changed to "N/A"
#'

wbccpd <- long |>
  filter(
    `Geographic Area` == "British Columbia",
    NOC != "#T"
  ) %>%
  mutate(
    cagr = map(data, get_cagrs),
    jos = map(data, get_jos),
    breakdown = map(data, get_breakdown),
    current_employment = map_dbl(data, get_current)
  )|>
  select(-data) |>
  unnest(cagr) |>
  unnest(jos) |>
  unnest(breakdown) |>
  mutate(
    across(c(ffy, sfy, jo, rep, exp), ~ round(.x, -1)),
    across(where(is.numeric), ~ if_else(current_employment < 20, NA_real_, .x))
  ) |>
  ungroup() |>
  select(-`Geographic Area`, -current_employment)

#' This is the "by NOC and geographic_area" breakdown of the labour market.
#' 1. Suppression rule:
#'   If the current year's employment is below 20, then all numeric variables should be changed to "N/A"
wbccpd_regional <- long |>
  filter(NOC != "#T") |>
  mutate(
    current_employment = map_dbl(data, get_current),
    ten_cagr = map_dbl(data, get_10_cagr),
    jos = map_dbl(data, get_10_jo)
  ) |>
  select(-data) |>
  mutate(
    across(contains("current_employment") | contains("jos"), ~ round(.x, -1)),
    across(where(is.numeric), ~ if_else(current_employment < 20, NA_real_, .x))
  ) |>
  arrange(`Geographic Area`) |>
  pivot_wider(
    names_from = `Geographic Area`,
    values_from = -all_of(c("NOC", "Description", "Geographic Area")),
    names_glue = "{`Geographic Area`}_{.value}",
    names_vary = "slowest"
  ) %>%
  select(!contains("British Columbia"), everything())

career_profile_regional_excel <- wbccpd_regional %>%
  add_column(` ` = "", .after = "Cariboo_jos") %>%
  add_column(`  ` = "", .after = "Kootenay_jos") %>%
  add_column(`   ` = "", .after = "Mainland South West_jos") %>%
  add_column(`    ` = "", .after = "North Coast & Nechako_jos") %>%
  add_column(`     ` = "", .after = "North East_jos") %>%
  add_column(`      ` = "", .after = "Thompson Okanagan_jos") %>%
  select(-contains("British Columbia"))


wb <- loadWorkbook(here("new_templates", "WorkBC_Career_Profile_Data.xlsx"))
writeWorksheet(wb, wbccpd, "Provincial Outlook", 4, 1, header = FALSE)
writeWorksheet(wb, career_profile_regional_excel, "Regional Outlook", 5, 1, header = FALSE)
saveWorkbook(wb, here(
  "out",
  paste0(
    "WorkBC_Career_Profile_Data_",
    fyod,
    ".xlsx"
  )
))

# WorkBC_Career_Trek.xlsx------------------------------
#' Provide data for all the NOCS because have not provided the filter

wbcct_jo <- jo |>
  filter(
    Industry == "All industries",
    `Geographic Area` == "British Columbia"
  ) |>
  ungroup() |>
  select(NOC = NOC_2021, Description = NOC_2021_Description, jo) |>
  mutate(jo = round(jo, -1))

wbcct_emp <- emp |>
  filter(
    NOC != "#T",
    Industry == "All industries",
    `Geographic Area` == "British Columbia"
  ) |>
  ungroup() |>
  select(-Industry, -`Geographic Area`) |>
  group_by(NOC, Description) |>
  nest() |>
  mutate(cagr = map_dbl(data, get_10_cagr)) |>
  select(-data)

wbcct <- full_join(wbcct_emp, wbcct_jo) |>
  right_join(career_trek_have_videos) |>
  select(
    `NOC 2016`,
    NOC,
    `Sr No`,
    `NOC Title (2016)`,
    Description,
    starts_with("Occupation"),
    cagr,
    jo
  )

wb <- loadWorkbook(here("new_templates", "WorkBC_Career_Trek.xlsx"))
writeWorksheet(wb, wbcct, "LMO", 2, 1, header = FALSE)
saveWorkbook(wb, here(
  "out",
  paste0(
    "WorkBC_Career_Trek_",
    fyod,
    ".xlsx"
  )
))

# WorkBC_Industry_Profile.xlsx------------------

wbcip_jo <- jo |>
  filter(
    NOC_2021 != "#T",
    `Geographic Area` == "British Columbia",
    Industry != "All industries"
  ) |>
  mutate(`Industry (sub-industry)` = str_to_lower(str_replace_all(Industry, " ", "_"))) |>
  full_join(ind_char) |>
  group_by(`Industry (aggregate)`) |>
  summarize(jo = sum(jo))

wbcip <- emp |>
  filter(
    NOC != "#T",
    `Geographic Area` == "British Columbia",
    Industry != "All industries"
  ) |>
  mutate(`Industry (sub-industry)` = str_to_lower(str_replace_all(Industry, " ", "_"))) |>
  full_join(ind_char) |>
  group_by(`Industry (aggregate)`, name, Variable) |>
  summarize(value = sum(value)) |>
  group_by(`Industry (aggregate)`, .add = FALSE) |>
  nest() |>
  mutate(
    levels = map(data, get_levels),
    cagrs = map(data, get_cagrs),
    ty_cagr = map_dbl(data, get_10_cagr)
  ) |>
  unnest(levels) |>
  unnest(cagrs) |>
  select(-data) |>
  ungroup() |>
  mutate(
    current_share = round(100 * employment_current / sum(employment_current), 1),
    five_share = round(100 * employment_five / sum(employment_five), 1),
    ten_share = round(100 * employment_ten / sum(employment_ten), 1)
  ) |>
  full_join(wbcip_jo) |>
  select(
    contains("Industry"),
    jo,
    contains("share"),
    contains("employment"),
    contains("cagr")
  ) |>
  mutate(
    jo = round(jo, -1),
    across(contains("share"), ~ round(.x, 3)),
    across(contains("employment"), ~ round(.x, -1))
  )

wb <- loadWorkbook(here("new_templates", "WorkBC_Industry_Profile.xlsx"))
writeWorksheet(wb, wbcip, "Sheet1", 4, 1, header = FALSE)
saveWorkbook(wb, here(
  "out",
  paste0(
    "WorkBC_Industry_Profile_",
    fyod,
    ".xlsx"
  )
))

# WorkBC_Regional_Profile_Data.xlsx-----------------------

wbcrpd_s1 <- long |>
  filter(NOC == "#T") |>
  ungroup() |>
  select(-NOC, -Description) |>
  mutate(
    jos = map(data, get_jos),
    breakdown = map(data, get_breakdown),
    levels = map(data, get_emp_levels)
  ) |>
  unnest(jos) |>
  unnest(breakdown) |>
  unnest(levels) |>
  mutate(
    diff = employment_ten - employment_current,
    ty_cagr = map_dbl(data, get_10_cagr),
    cagrs = map(data, get_cagrs)
  ) |>
  unnest(cagrs) |>
  select(-data) |>
  mutate(across(c(ffy, sfy, jo, rep, exp, contains("employment"), diff), ~ round(.x, -1))) |>
  select(`Geographic Area`, ffy, sfy, jo, rep, rep_p, exp, exp_p, everything())

wbcrpd_s2 <- long |>
  filter(
    NOC != "#T",
    `Geographic Area` != "British Columbia"
  ) |>
  mutate(
    jo = map_dbl(data, get_10_jo),
    ten_cagr = map_dbl(data, get_10_cagr)
  ) |>
  group_by(`Geographic Area`) |>
  slice_max(order_by = jo, n = 10, with_ties = FALSE) |>
  select(NOC, Description, jo, ten_cagr, `Geographic Area`) |>
  mutate(jo = round(jo, -1))

wbcrpd_s3a <- jo |>
  filter(
    NOC_2021 == "#T",
    Industry != "All industries",
    `Geographic Area` != "British Columbia"
  ) |>
  group_by(`Geographic Area`) |>
  slice_max(order_by = jo, n = 10, with_ties = FALSE) |>
  select(-contains("NOC")) |>
  mutate(jo = round(jo, -1))

wbcrpd_s3b <- emp |>
  filter(
    NOC == "#T",
    Industry != "All industries",
    `Geographic Area` != "British Columbia"
  ) |>
  select(-NOC, -Description) |>
  group_by(`Geographic Area`, Industry) |>
  nest() |>
  mutate(ten_cagr = map_dbl(data, get_10_cagr)) |>
  select(-data)

wbcrpd_s3 <- inner_join(wbcrpd_s3a, wbcrpd_s3b) |>
  select(Industry, jo, ten_cagr, `Geographic Area`) |>
  arrange(`Geographic Area`)

wb <- loadWorkbook(here("new_templates", "WorkBC_Regional_Profile_Data.xlsx"))
writeWorksheet(wb, wbcrpd_s1, "Regional Profiles - LMO", 5, 1, header = FALSE)
writeWorksheet(wb, wbcrpd_s2, "Top Occupation", 4, 1, header = FALSE)
writeWorksheet(wb, wbcrpd_s3, "Top Industries", 4, 1, header = FALSE)
saveWorkbook(wb, here(
  "out",
  paste0(
    "WorkBC_Regional_Profile_Data_",
    fyod,
    ".xlsx"
  )
))


#wage data-----------------------------------

desired_nocs <- occ_char|>
    select(NOC, `Occupation Title`=Description)|>
    mutate(NOC=str_remove(NOC,"#"))

wage <-  read_excel(here("data","WorkBC_2023_Wage_Data.xlsx"))|>
  right_join(desired_nocs)|>
  arrange(NOC)|>
  mutate(across(starts_with("ESDC")|starts_with("Calculated"), make_na))

write.xlsx(wage, here(
  "out",
  paste0(
    "Wages_where_NOCS_00011:00015_rolled_into_00018_(missing)_",
    fyod,
    ".xlsx"
  )
))

#' top skills by occupation---------------------------------
#' The skills data does not match LMO occupations:

raw_skills <-  read_excel(here("data","Top skills by NOC2021 occupations.xlsx")) #this file has missing data

skill_names <- unique(raw_skills$`Skills & Competencies`)

skills <- raw_skills|>
  right_join(desired_nocs, by=c("NOC2021"="NOC", "NOC2021 Title"="Occupation Title"))|> #the desired NOCs
  arrange(NOC2021)|>
  group_by(NOC2021, `NOC2021 Title`)|>
  nest()|>
  mutate(data=map(data, fill_skills))|>
  unnest(data)


write.xlsx(skills, here(
  "out",
  paste0(
    "Top skills_where_NOCS_00011:00015_rolled_into_00018_(missing)_",
    fyod,
    ".xlsx"
  )
))

#Occupational interests---------------------------------

occ_int <- read_excel(here("data","Occupational interest by NOC2021 occupation.xlsx"))|>
  select(NOC=`NOC 2021`, Options, `Occupational Interest`)|>
  right_join(desired_nocs)|>
  select(-`Occupation Title`)|>
  group_by(NOC)|>
  nest()|>
  mutate(data=map(data, fill_interests))|>
  unnest(data)|>
  arrange(NOC)|>
  mutate(NOC=paste0("#",NOC))

write.xlsx(occ_int, here(
  "out",
  paste0(
    "Occupational_Interests_",
    fyod,
    ".xlsx"
  )
))

