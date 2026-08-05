#' This script produces the LMO files for WorkBC. Requires inputs:
fyod <- 2026 # first year of data: need to reset each year
#' "employment.csv" (4castviewer: all nocs, all industries, all regions)
#' "job_openings.xlsx"  (4castviewer: all nocs, all industries, all regions)
#' "expansion.xlsx" (4castviewer: all nocs, all industries, all regions)
#' "replacement.xlsx" (4castviewer: all nocs, all industries, all regions)
#' "ftpt2125NOCp1.csv" (run the SAS scripts... will need to be changed 2027)
#' "ftpt2125NOCp2.csv" (run the SAS scripts... will need to be changed 2027)
#'  industry_mapping_****.xlsx
#'  one file that contains "Occupational Characteristics" in its name (Nicole)
#'  one file that contains "Salary" in its name (Nicole)
#' "Onet data for skills, knowledge, abilites, work activities (ONET)
#'
#' To Run: source me, and then you will have to verify the hoo geography is right...

# "constants"------------------------
fyfn <- fyod + 5
tyfn <- fyod + 10
# WorkBC wants mixed case names...
exceptions <- c("And", "For", "Of", "The", "In", "At", "On", "With", "As", "By")
pattern <- paste0("\\b(", paste(exceptions, collapse = "|"), ")\\b")

# libraries---------------------------
library(tidyverse)
library(here)
library(vroom)
library(readxl)
library(rvest)
library(janitor)
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
      values_from = c(aggregate_industry, job_openings, `%`),
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
  tibble(jo_ffy = ffy, jo_sfy = sfy)
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
  tibble(jo_ty = jo, rep_p = 100 * rep_p, rep = rep, exp_p = 100 * exp_p, exp = exp)
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
get_skill_names <- function(tbbl){
  on_own <- tbbl$`Skills & Competencies`
  together <- paste(on_own, collapse = ", ")
  temp <- as_tibble(t(c(on_own, together)), .name_repair = "minimal")
  colnames(temp) <- c("Skill1", "Skill2", "Skill3", "Skills: Top 3")
  temp
}
fix_region_names <- function(region){
  region|>
    str_replace_all("Mainland South West", "Mainland/Southwest")|>
    str_replace_all("Vancouver Island Coast", "Vancouver Island/Coast")|>
    str_replace_all("North Coast & Nechako", "North Coast and Nechako")|>
    str_replace_all("Thompson Okanagan", "Thompson-Okanagan")|>
    str_replace_all("North East", "Northeast")
}
read_data <- function(file_name){
  read_excel(here("data",file_name))%>%
    clean_names()%>%
    select(o_net_soc_code, element_name, scale_name, data_value)%>%
    pivot_wider(names_from = scale_name, values_from = data_value)%>%
    mutate(Importance=10*(Importance-1)/4,
           Level=10*Level/7,
           score=sqrt(Importance*Level), #geometric mean of importance and level
           category=(str_split(file_name,"\\.")[[1]][1]))%>%
    unite(element_name, category, element_name, sep=": ")%>%
    select(-Importance, -Level)
}
get_nn <- function(q, tbbl){
  nn <- dbscan::kNN(tbbl, k = 21, sort=TRUE,  query = q)
  tibble(nearest_neighbours = rownames(tbbl)[as.vector(nn[["id"]])],
         distance = as.vector(nn[["dist"]]))
}

#Read in the data------------------------------

#top skills by NOC--------------------------------

essential_skills <- read_excel(here("data","Essential Skills.xlsx"))
transferable_skills <- read_excel(here("data","Transferable Skills.xlsx"))

skills <- bind_rows(essential_skills, transferable_skills)|>
  clean_names()|>
  select(onet_soc_code=contains("o_net"), element_name, scale_name, data_value)|>
  pivot_wider(names_from = scale_name, values_from = data_value)|>
  mutate(Importance=(Importance-min(Importance))/(max(Importance)-min(Importance))*100, #0-100 scale
         Level=(Level-min(Level))/(max(Level)-min(Level))*100 #0-100 scale
  )

onet_mapping <-  readr::read_csv(
  "https://raw.githubusercontent.com/bcgov/onet-noc2021-crosswalk/main/output/3.0.0/onet_to_noc2021_mapping.csv"
)

mapped_skills <- inner_join(onet_mapping, skills)|>
  mutate(Importance=Importance*down_weight,
         Level=Level*down_weight)|>
  group_by(noc_2021, noc_plus_title, element_name)|>
  summarize(Importance=round(sum(Importance, na.rm = TRUE),1),
            Level=round(sum(Level, na.rm = TRUE), 1))|>
  mutate(importance_description=case_when(
    Importance == 0 ~ "Not required",
    Importance < 40 ~ "Less important",
    Importance < 60 ~ "Important",
    Importance < 80 ~ "Very Important",
    TRUE ~ "Extremely Important"),
    level_description=case_when(
      Level == 0 ~ "Not required",
      Level < 33 ~ "Basic proficiency",
      Level < 66 ~ "Medium proficiency",
      TRUE ~ "High proficiency"),
    noc_title =str_sub(noc_plus_title, 8)
  )|>
  ungroup()|>
  select(NOC2021 = noc_2021, `NOC2021 Title` = noc_title,
         `Skills & Competencies` = element_name, `Importance Score`=Importance,
         `Importance Description`=importance_description, `Level Score`=Level,
         `Level Description`=level_description)

length(unique(mapped_skills$NOC2021)) #missing skills data for some NOCs

openxlsx::write.xlsx(mapped_skills, here("out", paste0("skills_data_for_career_profiles_",today(),".xlsx")))

skills_df <- mapped_skills |>
  group_by(NOC=NOC2021)|>
  mutate(NOC=paste0("#",NOC))|>
  slice_max(`Importance Score`, n=3, with_ties = FALSE)|>
  nest()|>
  mutate(skills=map(data, get_skill_names))|>
  select(-data)|>
  unnest(skills)

#ONET interests---------------------------

onet_interests <- read_excel(here("data","Career Interest Types.xlsx"))|>
  clean_names()|>
  select(onet_soc_code=o_net_soc_code, element_name, data_value)|>
  filter(!str_detect(element_name, "High-Point"))

mapped_interests <- inner_join(onet_mapping, onet_interests)|>
  mutate(weighted=down_weight*data_value)|>
  group_by(noc_2021, element_name)|>
  summarize(score=sum(weighted))|>
  group_by(noc_2021)|>
  slice_max(score, n=3, with_ties = FALSE)|>
  mutate(rank=rank(-score,ties.method = "random"))|>
  pivot_wider(id_cols = noc_2021, names_from = rank, values_from = element_name, names_prefix = "interest")|>
  mutate(Interests=paste(interest1, interest2, interest3, sep=", "),
         noc_2021=paste0("#",noc_2021))|>
  select(noc_2021, Interests)

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

ftpt <- bind_rows(ftpt, senior_managers)|>
  filter(!NOC_2021 %in% c("#00011", "#00012", "#00013", "#00014", "#00015"))

teer_description <- tibble(
  TEER_description = c("Management",
                       "University Degree",
                       "College Diploma or Apprenticeship, 2 or more years",
                       "College Diploma or Apprenticeship, less than 2 years",
                       "High School Diploma",
                        "No Formal Education"),
  TEER=as.character(0:5)
  )

# salary-----------------------------------

salary <- read_excel(here("data", list.files(here("data"), pattern = "Salary")), na = "N/A")|>
  select(NOC, contains("Rate"), median_salary=contains("salary"))|>
  mutate(NOC=paste0("#",str_pad(NOC, width=5, pad=0)))
colnames(salary) <- str_remove_all(colnames(salary),"Job Bank")

# occupation characteristics (this should have proper naming (Senior))
occ_groups <- read_excel(
  here("data",
       list.files(here("data"),
                  pattern = "Occupational Characteristics")),
  skip = 3,
  na = c("x","-", "N/A", "n/a"))|>
  select(NOC, Description, TEER, contains("OCC group"))|>
  full_join(salary)|>
  select(-`Occ Group: Skilled Trades Certification`)

# industry characteristics---------------------------------------------------

ind_char <- read_excel(here("data",list.files(here("data"),pattern = "industry_mapping")))|>
  select(contains("detailed"), contains("aggregate"))|>
  distinct()

# job openings----------------------------------------------
jo <- read_excel(here("data", "job_openings.xlsx"), skip = 3) |>
  janitor::remove_empty() |>
  filter(
    Variable == "Job Openings",
    !`Geographic Area` %in% c("North", "South East")
  ) |>
  pivot_longer(cols = starts_with("2")) |>
  filter(name > fyod)|>
  summarize(jo = sum(value), .by = c(NOC, Industry, `Geographic Area`)) |>
  rename(
    NOC_2021 = NOC
  )

# job openings and components-------------------------

jo_temp <- read_excel(here("data", "job_openings.xlsx"), skip = 3) |>
  janitor::remove_empty()

r_temp <- read_excel(here("data", "replacement.xlsx"), skip = 3) |>
  janitor::remove_empty()

e_temp <- read_excel(here("data", "expansion.xlsx"), skip = 3) |>
  janitor::remove_empty()

jore <- bind_rows(jo_temp, r_temp, e_temp)|>
  filter(
    Variable %in% c("Job Openings", "Expansion Demand", "Replacement Demand"),
    !`Geographic Area` %in% c("North", "South East"),
    Industry == "All industries"
  ) |>
  pivot_longer(cols = starts_with("2")) |>
  filter(name > fyod)|>
  select(-Industry, -Description)

# employment to add to job openings and components
emp <- read_excel(here("data", "employment.xlsx"), skip = 3) |>
  janitor::remove_empty()

emp_to_add_to_jore <-emp|>
  filter(
    !`Geographic Area` %in% c("North", "South East"),
    Industry == "All industries"
  ) |>
  pivot_longer(cols = starts_with("2")) |>
  select(-Industry, -Description)

noc_descriptions <- occ_groups|>
  select(NOC, Description)

long <- bind_rows(jore, emp_to_add_to_jore) |>
  left_join(noc_descriptions, by=c("NOC"="NOC"))|>
  group_by(NOC, Description, `Geographic Area`) |>
  nest()

# high opportunity occupations------------------------------------

hoo_region_names <- read_excel(here("data","0. LMO 2026E Regional Summary.xlsx"),
                          sheet="All Regions",
                          skip=3,
                          col_names = FALSE,
                          n_max = 1)|>
  t()|>
  na.omit()

hoo_data <- read_excel(here("data","0. LMO 2026E Regional Summary.xlsx"),
                       sheet="All Regions",
                       skip=4)|>
  select(contains("NOC"), contains("HOO"))

colnames(hoo_data)[3:10] <- hoo_region_names

occ_char <- full_join(occ_groups, hoo_data, by=c("NOC"="NOC 2021"))|>
  select(-`NOC 2021 Occupation Title`)

# employment----------------------------------

emp <- read_excel(here("data", "employment.xlsx"), skip = 3) |>
  janitor::remove_empty() |>
  filter(!`Geographic Area` %in% c("North", "South East")) |>
  pivot_longer(cols = starts_with("2"))

# PROCESS THE DATA-----------------------------

# CareerDiscoveryQuizzesJobOpenings- master.csv--------------------------------

cdqjom <- occ_char |>
  select(
    NOC_2021 = NOC,
    NOC_2021_Description = Description,
    TEER,
    median_salary)|>
  mutate(Link = paste0(
      "https://www.workbc.ca/Jobs-Careers/Explore-Careers/Browse-Career-Profile/",
      str_sub(NOC_2021, 2)),
      JobBank2 = paste0(
      "https://www.workbc.ca/jobs-careers/find-jobs/jobs.aspx?searchNOC=",
      str_sub(NOC_2021, 2)),
      TEER = as.character(TEER))|>
  inner_join(teer_description) |>
  relocate(TEER_description, .after = TEER)

# Career_Search_Tool_Job_Openings----------------------------

temp <- cdqjom %>%
  left_join(jo, by=c("NOC_2021"="NOC_2021"))|>
  mutate(jo = round(jo, -1))|>
  fuzzyjoin::stringdist_left_join(ind_char, by=c("Industry"="lmo_detailed_industry"))|>
  select(NOC_2021,
    NOC_2021_Description,
    `Industry (sub-industry)` = Industry,
    Region = `Geographic Area`,
    "Job Openings {fyod}-{tyfn}" := jo,
    TEER,
    TEER_description,
    `Salary (calculated median salary)`=median_salary,
    Link,
    JobBank2,
    `Industry (aggregate)`=aggregate_industry
  ) |>
  mutate(`Industry (aggregate)` = if_else(is.na(`Industry (aggregate)`),"All industries", `Industry (aggregate)`),
         `Industry (aggregate)` = str_replace_all(str_to_title(`Industry (aggregate)`), regex(pattern, ignore_case = TRUE), tolower),
         `Industry (sub-industry)` = str_replace_all(str_to_title(`Industry (sub-industry)`), regex(pattern, ignore_case = TRUE), tolower),
         Region=fix_region_names(Region)
         ) |>
  filter(NOC_2021 != "#T") |>
  left_join(ftpt)

length(unique(temp$NOC_2021))

write.xlsx(temp, here(
    "out",
    paste0(
      "Career_Search_Tool_Job_Openings_",
      today(),
      ".xlsx"
      )
    ),
    keepNA=TRUE,
    na.string="N/A"
    )

# career_search_tool_occupation_groups_manual_update.xlsx-----------------

workbc_regions <- fix_region_names(unique(jo$`Geographic Area`))
region_mapping <- tibble(Region=c("British Columbia",
                                "Mainland/Southwest",
                                "Vancouver Island/Coast",
                                "North Coast and Nechako",
                                "Cariboo",
                                "Kootenay",
                                "Northeast",
                                "Thompson-Okanagan"),
                         short=c("BC","MSW","VIC","NCN", "CAR", "KOO", "NE", "TOK"))

cstogmu_hoo <- occ_char|>
  select(NOC, Description, BC:NE)|>
  pivot_longer(cols=-c(NOC, Description))|>
  filter(!str_detect(value, "Non-"))|>
  select(NOC, Description, short=name)|>
  left_join(region_mapping)|>
  mutate(`Occupational category`="High opportunity occupations")|>
  select(-short)

cstogmu_stem <- occ_char |>
  filter(`Occ Group: STEM` == "STEM") |>
  select(NOC, Description) |>
  mutate(`Occupational category` = "STEM")

cstogmu_trades <- occ_char |>
  filter(`Occ Group: Trades` == "Trades") |>
  select(NOC, Description, `Occupational category` = `Occ Group: Construction Trades`)

cstogmu_care <- occ_char |>
  filter(`Occ Group: Deloitte Care Economy` == "Care Economy") |>
  select(NOC, Description) |>
  mutate(`Occupational category` = "Care Economy")

cstogmu_manage <- occ_char |>
  select(NOC, Description, TEER) |>
  mutate(`Occupational category` = if_else(TEER == 0,
      "Management occupations",
      "Non-management occupations"
    )
  ) |>
  select(-TEER)


no_regions <- bind_rows(
  cstogmu_manage,
  cstogmu_care,
  cstogmu_trades,
  cstogmu_stem
)|>
  mutate(`Occupational category`=str_to_sentence(`Occupational category`),
         `Occupational category`=str_replace_all(`Occupational category`,"Stem","STEM"))

crossing(no_regions, Region=workbc_regions)|>
  bind_rows(cstogmu_hoo)|>
  na.omit()|>
  write.xlsx(here(
    "out",
    paste0(
      "career_search_tool_occupation_groups_manual_update_",
      today(),
      ".xlsx"
    )
  ),
  keepNA=TRUE,
  na.string="N/A"
  )

# HOO BC and Region for new tool.xlsx-----------------------------

hoo_no_jo <- cstogmu_hoo |>
  left_join(skills_df)|>
  left_join(salary)|>
  left_join(mapped_interests, by=c("NOC"="noc_2021"))

#job openings by occupation and region

jo_by_noc_region <- jo|>
  filter(Industry=="All industries")|>
  ungroup()|>
  select(-Industry)|>
  rename(NOC=NOC_2021, Geography=`Geographic Area`)|>
  mutate(Geography=fix_region_names(Geography),
         jo=round(jo, -1))

#merge organize, write

left_join(hoo_no_jo, jo_by_noc_region, by = c("NOC"="NOC","Region"="Geography"))|>
  select(
    `Occupation Title` = Description,
    "Job Openings {fyod}-{tyfn}" := jo,
    "Wage Rate Low {fyod}" := contains("Wage Rate Low"),
    "Wage Rate Median {fyod}" := contains("Wage Rate Median"),
    "Wage Rate High {fyod}" := contains("Wage Rate High"),
    `Median Annual Salary` = median_salary,
    Interests,
    `Skills and Competencies (Top 3 together)`=`Skills: Top 3`,
    First=Skill1,
    Second=Skill2,
    Third=Skill3,
    `#NOC` = NOC,
    Geography=Region
  ) |>
  mutate(NOC = str_sub(`#NOC`, 2, -1), .before = `#NOC`) |>
  mutate(TEER = str_sub(NOC, 2, 2), .before = Geography) |>
  mutate(Geography=fix_region_names(Geography))|>
write.xlsx(here(
    "out",
    paste0(
      "HOO_BC_and_Region_for_new_tool_",
      today(),
      ".xlsx"
    )
  ),
  keepNA=TRUE,
  na.string="N/A"
  )

# Job Openings by Industry_LMO.xlsx------------------------

noc_descriptions <- occ_groups|>
  select(NOC, Description)

career_profiles <- jo |>
  inner_join(noc_descriptions, by=c("NOC_2021"="NOC"))|>
  filter(`Geographic Area` == "British Columbia") |>
  fuzzyjoin::stringdist_join(ind_char, by=c("Industry"="lmo_detailed_industry"))|>
  group_by(NOC_2021, Description, aggregate_industry) |>
  summarize(job_openings = sum(jo)) |>
  slice_max(job_openings, n = 5, with_ties = FALSE) |>
  mutate(
    `%` = 100 * round(job_openings / sum(job_openings), 2),
    job_openings = round(job_openings, -1)
  ) |>
  group_by(NOC_2021) |>
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
  fuzzyjoin::stringdist_join(ind_char, by=c("Industry"="lmo_detailed_industry"))|>
  na.omit() |>
  group_by(aggregate_industry) |>
  summarize(jo = round(sum(jo), -1)) |>
  rename("Job Openings {fyod}-{tyfn}" := jo)

wb <- loadWorkbook(here("templates", list.files(here("templates"), pattern = "Job_Openings_by_Industry")))
setMissingValue(wb, value = "N/A")
writeWorksheet(wb, career_profiles, "Career Profiles", 5, 1, header = FALSE)
writeWorksheet(wb, `Job Openings by industry`, "Job Openings by industry", 3, 1, header = FALSE)
saveWorkbook(wb, here(
  "out",
  paste0(
    "Job_Openings_by_Industry_LMO_",
    today(),
    ".xlsx"
  )
))

# jo_by_region------------------------------

jo_by_region <- jo |>
  filter(
    NOC_2021 == "#T",
    Industry == "All industries"
  ) |>
  mutate(jo = round(jo, -1)) |>
  ungroup() |>
  select(`Geographic Area`, ten_sum_jo = jo)

# top_10_careers_by_aggregate_industry.xlsx----------------------------

jo |>
  filter(
    NOC_2021 != "#T",
    `Geographic Area` == "British Columbia"
  ) |>
  fuzzyjoin::stringdist_join(ind_char, by=c("Industry"="lmo_detailed_industry"))|>
  left_join(noc_descriptions, by=c("NOC_2021"="NOC"))|>
  group_by(NOC_2021, NOC_2021_Description=Description, aggregate_industry)|>
  summarize(jo = sum(jo)) |>
  group_by(aggregate_industry, .add = FALSE) |>
  slice_max(jo, n = 10, with_ties = FALSE) |>
  select(`Aggregate Industry` = aggregate_industry, NOC = NOC_2021, Occupation = NOC_2021_Description, jo) |>
  mutate("Job Openings {fyod}-{tyfn}" := round(jo, -1)) |>
  select(-jo) |>
  na.omit() |>
  write.xlsx(here(
    "out",
    paste0(
      "top_10_careers_by_aggregate_industry_",
      today(),
      ".xlsx"
    )
  ),
  keepNA=TRUE,
  na.string="N/A"
  )

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
    ty_cagr=map_dbl(data, get_10_cagr),
    jos = map(data, get_jos),
    breakdown = map(data, get_breakdown),
    current_employment = map_dbl(data, get_current)
  )|>
  select(-data) |>
  unnest(cagr) |>
  unnest(jos) |>
  unnest(breakdown)|>
  mutate(
    across(c(jo_ffy, jo_sfy, jo_ty, rep, exp), ~ round(.x, -1)),
    across(where(is.numeric), ~ if_else(current_employment < 20, NA_real_, .x))
  ) |>
  ungroup() |>
  select(-`Geographic Area`, -current_employment)

length(unique(wbccpd$NOC))

#' This is the "by NOC and geographic_area" breakdown of the labour market.
#' 1. Suppression rule:
#'   If the current year's employment is below 20, then all numeric variables should be changed to "N/A"
wbccpd_regional_long <- long |>
  filter(NOC != "#T") |>
  mutate(
    current_employment = map_dbl(data, get_current),
    ty_cagr = map_dbl(data, get_10_cagr),
    jo_ty = map_dbl(data, get_10_jo)
  ) |>
  select(-data) |>
  mutate(
    across(contains("current_employment") | contains("jo_ty"), ~ round(.x, -1)),
    across(where(is.numeric), ~ if_else(current_employment < 20, NA_real_, .x))
  ) |>
  arrange(`Geographic Area`)

#try to convince them tidy data is better...

openxlsx::write.xlsx(wbccpd_regional_long,
                     here("out", paste0("career_profile_regional_long_",today(),".xlsx")),
                     na.string="N/A",
                     keepNA=TRUE)

#they are unlikely to agree...

wbccpd_regional <- wbccpd_regional_long|>
  pivot_wider(
    names_from = `Geographic Area`,
    values_from = -all_of(c("NOC", "Description", "Geographic Area")),
    names_glue = "{`Geographic Area`}_{.value}",
    names_vary = "slowest"
  ) %>%
  select(!contains("British Columbia"), everything())

career_profile_regional_excel <- wbccpd_regional %>%
  add_column(` ` = "", .after = "Cariboo_jo_ty") %>%
  add_column(`  ` = "", .after = "Kootenay_jo_ty") %>%
  add_column(`   ` = "", .after = "Mainland South West_jo_ty") %>%
  add_column(`    ` = "", .after = "North Coast & Nechako_jo_ty") %>%
  add_column(`     ` = "", .after = "North East_jo_ty") %>%
  add_column(`      ` = "", .after = "Thompson Okanagan_jo_ty") %>%
  select(-contains("British Columbia"))

length(unique(career_profile_regional_excel$NOC))

wb <- loadWorkbook(here("templates", list.files(here("templates"), pattern = "WorkBC_Career_Profile")))
setMissingValue(wb, value = "N/A")
writeWorksheet(wb, wbccpd, "Provincial Outlook", 4, 1, header = FALSE)
writeWorksheet(wb, career_profile_regional_excel, "Regional Outlook", 5, 1, header = FALSE)
saveWorkbook(wb, here(
  "out",
  paste0(
    "WorkBC_Career_Profile_Data_",
    today(),
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
  fuzzyjoin::stringdist_join(ind_char, by=c("Industry"="lmo_detailed_industry"))|>
  group_by(aggregate_industry) |>
  summarize(jo = sum(jo))|>
  rename(jo_ty=jo)

wbcip <- emp |>
  filter(
    NOC != "#T",
    `Geographic Area` == "British Columbia",
    Industry != "All industries"
  ) |>
  fuzzyjoin::stringdist_join(ind_char, by=c("Industry"="lmo_detailed_industry"))|>
  group_by(aggregate_industry, name, Variable) |>
  summarize(value = sum(value)) |>
  group_by(aggregate_industry, .add = FALSE) |>
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
  full_join(wbcip_jo)|>
  mutate(
    jo_ty = round(jo_ty, -1),
    across(contains("share"), ~ round(.x, 3))
    #  across(contains("employment"), ~ round(.x, -1))
  )|>
  select(
    contains("Industry"),
    jo_ty,
    contains("share"),
    contains("cagr")
  )

wb <- loadWorkbook(here("templates", list.files(here("templates"), pattern = "WorkBC_Industry_Profile")))
setMissingValue(wb, value = "N/A")
writeWorksheet(wb, wbcip, "Sheet1", 3, 1, header = FALSE)
saveWorkbook(wb, here(
  "out",
  paste0(
    "WorkBC_Industry_Profile_",
    today(),
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
  mutate(across(c(jo_ffy, jo_sfy, jo_ty, rep, exp, contains("employment"), diff), ~ round(.x, -1))) |>
  select(Region=`Geographic Area`, jo_ffy, jo_sfy, jo_ty, rep, rep_p, exp, exp_p, diff, ty_cagr, ffy_cagr, sfy_cagr)|>
  mutate(Region=fix_region_names(Region))

wbcrpd_s2 <- long |>
  filter(
    NOC != "#T",
    `Geographic Area` != "British Columbia"
  ) |>
  mutate(
    jo = map_dbl(data, get_10_jo)
  ) |>
  group_by(`Geographic Area`) |>
  slice_max(order_by = jo, n = 10, with_ties = FALSE) |>
  select(NOC, Description, jo, Region=`Geographic Area`) |>
  mutate(jo = round(jo, -1),
         Region=fix_region_names(Region))|>
  rename(jo_ty=jo)

wb <- loadWorkbook(here("templates", list.files(here("templates"), pattern = "WorkBC_Regional_Profile")))
setMissingValue(wb, value = "N/A")
writeWorksheet(wb, wbcrpd_s1, "Regional Profiles - LMO", 5, 1, header = FALSE)
writeWorksheet(wb, wbcrpd_s2, "Top Occupation", 4, 1, header = FALSE)
saveWorkbook(wb, here(
  "out",
  paste0(
    "WorkBC_Regional_Profile_Data_",
    today(),
    ".xlsx"
  )
))

#Career Transition Tool_Opportunities-------------------------

#the onet data-----------------------------------
tbbl <- tibble(file=c("Essential Skills.xlsx",
                      "Transferable Skills.xlsx",
                      "Abilities.xlsx",
                      "Knowledge.xlsx",
                      "Work Activities.xlsx"))%>%
  mutate(data=map(file, read_data))%>%
  select(-file)%>%
  unnest(data)|>
  inner_join(onet_mapping, by=c("o_net_soc_code"="onet_soc_code"))|>
  mutate(score=score*down_weight)|>
  group_by(noc_plus_title, element_name)%>%
  summarize(score=sum(score))|>
  pivot_wider(id_cols=noc_plus_title, names_from = element_name, values_from = score)|>
  column_to_rownames("noc_plus_title")

pca <- prcomp(tbbl, center = TRUE, scale. = TRUE)

first_pca <-pca[["x"]][,1:12]%>% #keep only the first 5 principal components.
  as.data.frame()

unrestricted <- data.frame(first_pca)%>%
  mutate(data=list(first_pca))%>%
  rownames_to_column(var="noc")%>%
  nest(query = starts_with("PC"))%>%
  mutate(nearest=map2(query, data, get_nn))%>%
  select(-data,-query)%>%
  unnest(nearest)%>%
  filter(noc!=nearest_neighbours)%>%
  arrange(noc)%>%
  mutate(`Current Occupation TEER`=str_sub(noc,2,2), .after="noc")%>%
  mutate(`Career Option TEER`=str_sub(nearest_neighbours,2,2), .after="nearest_neighbours")

for_workBC <- unrestricted|>
  mutate(`Current Occupation`=str_replace_all(noc,": "," - "))|>
  separate(noc, into = c("Current Occupation (NOC)", "Current Occupation Title"), sep=": ")|>
  separate(nearest_neighbours, into = c("Career Option (NOC)", "Career Option Title"), sep=": ")|>
  mutate(similarity=if_else(distance>median(distance), "medium", "high"))|>
  select(`Current Occupation (NOC)`,
         `Current Occupation Title`,
         `Current Occupation`,
         `Current Occupation TEER`,
         `Career Option (NOC)`,
         `Career Option Title`,
         `Career Option TEER`,
         distance,
         similarity)

teer_description <- read_csv(here("data","teer_description_short.csv"))|>
  mutate(teer=as.character(teer))

for_workBC_w_teer_description <- for_workBC|>
  full_join(teer_description, by=c("Current Occupation TEER"="teer"))|>
  select(-`Current Occupation TEER`)|>
  rename(`Current Occupation TEER`=data)|>
  full_join(teer_description, by=c("Career Option TEER"="teer"))|>
  select(-`Career Option TEER`)|>
  rename(`Career Option TEER`=data)|>
  select(`Current Occupation (NOC)`,
         `Current Occupation Title`,
         `Current Occupation`,
         `Current Occupation TEER`,
         `Career Option (NOC)`,
         `Career Option Title`,
         `Career Option TEER`,
         distance,
         similarity)

openxlsx::write.xlsx(for_workBC_w_teer_description,
                     here("out",
                          paste0("career_transition_tool_opportunities_",today(),".xlsx")))

# Not part of the workbc files

# occ_char|>
#   left_join(skills_df)|>
#   left_join(mapped_interests, by=c("NOC"="noc_2021"))|>
#   openxlsx::write.xlsx(here("out", paste0("occupational_characteristics_",today(),".xlsx")))
#
#



