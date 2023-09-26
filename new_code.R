library(tidyverse)
library(here)
library(vroom)
library(readxl)
library(rvest)
library(openxlsx)
library(lubridate)
library(XLConnect)
current_year <- year(today())
tyfn <- current_year+10
#functions------------------
# Function to quickly export data
write_workbook <- function(data, sheetname, startrow, startcol) {
  writeWorksheet(
    wb,
    data,
    sheetname,
    startRow = startrow,
    startCol = startcol,
    header = FALSE
  )
}

make_wide <- function(dat){
  dat |>
    mutate(row = row_number()) |>
    tidyr::pivot_wider(names_from = row,
                       values_from = c(`Industry (aggregate)`, job_openings, `%`),
                       names_vary = "slowest")
}
#READ IN THE DATA------------------------
#full time part time---------------------
ftpt <- vroom(here("data",
                   list.files(here("data"),
                              pattern = "ftpt")),
              col_types=cols(
                SYEAR = col_double(),
                FTPT = col_character(),
                NOC_5 = col_character(),
                `_COUNT_` = col_double()
              )
              )|>
  na.omit()|>
  pivot_wider(id_cols = c(SYEAR, NOC_5), names_from = FTPT, values_from = `_COUNT_`)|>
  group_by(NOC_5)|>
  summarize(`Full-time`=sum(`Full-time`, na.rm = TRUE),
            `Part-time`=sum(`Part-time`, na.rm = TRUE)
            )|>
  mutate(prop_ft=`Full-time`/(`Full-time`+`Part-time`),
         `Part-time/full-time`=if_else(prop_ft<.70,
                                       "Higher Chance of Part-Time",
                                       "Higher Chance of Full-Time"
                                       ))|>
  select(NOC_2021=NOC_5, `Part-time/full-time`)|>
  mutate(NOC_2021=paste0("#", NOC_2021))
senior_managers <- tibble(NOC_2021="#00018", `Part-time/full-time`="Higher Chance of Full-Time")
ftpt <- bind_rows(ftpt, senior_managers)
#get teer descriptions from web------------------------------------
teer_description <- read_html("https://noc.esdc.gc.ca/Training/TeerCategory")|>
html_table()
teer_description <- teer_description[[1]]
teer_description <- teer_description|>
  mutate(TEER=as.character(`TEER category`))|>
  select(-`TEER category`, TEER_description=starts_with("Nature"))
#occupation characteristics--------------------------------------
occ_char <- read_excel(here("data",
                            list.files(here("data"), pattern = "Occupational Characteristics")),
                       skip=3)

ind_char <- read_excel(here("data",
                            list.files(here("data"), pattern = "industry")))|>
  transmute(`Industry (sub-industry)`=lmo_detailed_industry,
         `Industry (aggregate)`=str_to_title(str_replace_all(aggregate_industry,"_"," ")))|>
  distinct()
#job openings----------------------------------------------
jo <- vroom(here("data", "job_openings.csv"), skip=3)|>
  janitor::remove_empty()|>
  filter(Variable=="Job Openings",
         !`Geographic Area` %in% c("North","South East"))|>
  pivot_longer(cols=starts_with("2"))|>
  group_by(NOC, Description, Industry, `Geographic Area`)|>
  summarize(jo=round(sum(value),-1))|>
  rename(NOC_2021=NOC,
         NOC_2021_Description=Description)
#high opportunity occupations------------------------------------
hoo_sheets <- head(excel_sheets(here("data","LMO 2023E HOO BC and Regions 2023-08-23.xlsx")), -1)|>
  sort()
jo_regions <- unique(jo$`Geographic Area`)|>
  sort()
#CHECK ME!!!!!(hoo_geography could be wrong)---------------------------------
hoo_geography <- tibble(name=hoo_sheets,
                        Geography=jo_regions)

hoo <- map(set_names(hoo_sheets),
           read_excel, path = here("data","LMO 2023E HOO BC and Regions 2023-08-23.xlsx"),
           range="A5:C500" #cant possibly be more than 500 hoos?
)|>
  enframe()|>
  full_join(hoo_geography)|>
  select(-name)|>
  unnest(value)|>
  na.omit()

#PROCESS THE DATA-----------------------------
#All Occupations' TEERs.xlsx------------------------------

occ_char|>
  select(NOC_2021=NOC, NOC_2021_Description=Description)|>
  mutate(TEER=str_sub(NOC_2021, 3,3))|>
  inner_join(teer_description)|>
  write.xlsx(here("out","All Occupations' TEERs.xlsx"))

#CareerDiscoveryQuizzesJobOpenings- master.csv--------------------------------

cdqjom <- occ_char |>
  select(NOC_2021=NOC,
          NOC_2021_Description=Description,
          `Job Openings 2023-2033`="LMO Job Openings 2023-2033",
          TEER,
          `Salary (calculated median salary)`=contains("Median")
          )|>
  mutate(Link=paste0("https://www.workbc.ca/Jobs-Careers/Explore-Careers/Browse-Career-Profile/",
                     str_sub(NOC_2021,2)),
         JobBank2=paste0("https://www.workbc.ca/jobs-careers/find-jobs/jobs.aspx?searchNOC=",
                         str_sub(NOC_2021,2))
                         )|>
  mutate(TEER=as.character(TEER))|>
  inner_join(teer_description)|>
  relocate(TEER_description, .after=TEER)

cdqjom%>%
  write_csv(here("out","CareerDiscoveryQuizzesJobOpenings- master.csv"))

  #Career Search Tool Job Openings.xlsx-----------------------------

cdqjom%>%
  select(-contains("Job Openings"))|>
  full_join(jo)|>
  mutate(`Industry (sub-industry)`=str_to_lower(str_replace_all(Industry," ","_")))|>
  full_join(ind_char)|>
  select(NOC_2021,
         NOC_2021_Description,
         `Industry (sub-industry)`=Industry,
         Region=`Geographic Area`,
         "Job Openings {current_year}-{tyfn}":=jo,
         TEER,
         TEER_description,
         contains("Salary"),
         Link,
         JobBank2,
         `Industry (aggregate)`
  )|>
  mutate(`Industry (aggregate)`=if_else(is.na(`Industry (aggregate)`),
                                        "All industries",
                                        `Industry (aggregate)`))|>
  filter(NOC_2021!="#T")|>
  left_join(ftpt)|>
  write.xlsx(here("out","Career Search Tool Job Openings.xlsx"))

#career_search_tool_occupation_groups_manual_update.xlsx-----------------

cstogmu_hoo <- hoo|>
  select(NOC="#NOC (2021)", Region=Geography)|>
  mutate(`Occupational category`="High opportunity occupations")

cstogmu_stem <- occ_char|>
  filter(`Occ Group: STEM`=="STEM")|>
  select(NOC)|>
  mutate(Region="British Columbia",
         `Occupational category`="STEM"
         )

cstogmu_trades <- occ_char|>
  filter(`Occ Group: Trades`=="Trades")|>
  select(NOC,
         `Occupational category`=`Occ Group: Construction Trades`
         )|>
  mutate(Region="British Columbia")

cstogmu_care <- occ_char|>
  filter(`Occ Group: Deloitte Care Economy`=="Care Economy")|>
  select(NOC)|>
  mutate(Region="British Columbia",
         `Occupational category`="Care Economy"
  )

cstogmu_manage <- occ_char|>
  select(NOC, TEER)|>
  mutate(Region="British Columbia",
         `Occupational category`=if_else(TEER==0,
                                         "Management occupations",
                                         "Non-management occupations")
  )|>
  select(-TEER)

cstogmu_all <- occ_char|>
  select(NOC)|>
  mutate(Region="British Columbia",
         `Occupational category`="All"
  )

bind_rows(cstogmu_all,
                     cstogmu_manage,
                     cstogmu_care,
                     cstogmu_trades,
                     cstogmu_stem,
                     cstogmu_hoo
                     )|>
  write.xlsx(here("out","career_search_tool_occupation_groups_manual_update.xlsx"))


#HOO BC and Region for new tool.xlsx-----------------------------

hoo|>
  select(-contains("Job Openings"))|>
  left_join(occ_char, by=c("#NOC (2021)"="NOC"))|>
  select(`Occupation Title`=Description,
         "Job Openings {current_year}-{tyfn}":=contains("Job Openings") & !contains("/"),
         `Median Annual Salary`=contains("Employment Income"),
         `#NOC`=`#NOC (2021)`,
         Geography
         )|>
  mutate(NOC=str_sub(`#NOC`,2,-1), .before=`#NOC`)|>
  mutate(TEER=str_sub(NOC,2,2), .before=Geography)|>
  write.xlsx(here("out","HOO BC and Region for new tool.xlsx"))

#HOO List.xlsx-----------------------------

hoo|>
  filter(Geography=="British Columbia")|>
  select(-contains("Job Openings"))|>
  left_join(occ_char, by=c("#NOC (2021)"="NOC"))|>
  select(`Occupation Title`=Description,
         "Job Openings {current_year}-{tyfn}":=contains("Job Openings") & !contains("/"),
         `Median Annual Salary`=contains("Employment Income"),
         NOC=`#NOC (2021)`
        )|>
  mutate(NOC=str_sub(NOC, 2,-1),
         TEER=str_sub(NOC,2,2))|>
  write.xlsx(here("out","HOO List.xlsx"))

#Job Openings by Industry_2016 Census_2023 LMO.xlsx------------------------

career_profiles <- jo|>
  filter(`Geographic Area`=="British Columbia")|>
  mutate(Industry=str_to_lower(str_replace_all(Industry, " ","_")))|>
  full_join(ind_char, by=c("Industry"="Industry (sub-industry)"))|>
  na.omit()|>
  group_by(NOC_2021, NOC_2021_Description, `Industry (aggregate)`)|>
  summarize(job_openings = sum(jo))|>
  slice_max(job_openings, n=5, with_ties = FALSE)|>
  mutate(`%`=round(job_openings/sum(job_openings), 2))|>
  group_by(NOC_2021, NOC_2021_Description)|>
  nest()|>
  mutate(data=map(data, make_wide))|>
  unnest(data)|>
  relocate(`%_1`, .before=`job_openings_1`)|>
  relocate(`%_2`, .before=`job_openings_2`)|>
  relocate(`%_3`, .before=`job_openings_3`)|>
  relocate(`%_4`, .before=`job_openings_4`)|>
  relocate(`%_5`, .before=`job_openings_5`)


`Job Openings by industry` <- jo|>
  filter(`Geographic Area`=="British Columbia",
         NOC_2021=="#T")|>
  mutate(Industry=str_to_lower(str_replace_all(Industry, " ","_")))|>
  full_join(ind_char, by=c("Industry"="Industry (sub-industry)"))|>
  na.omit()|>
  group_by(`Industry (aggregate)`)|>
  summarize(jo = sum(jo))|>
  rename("Job Openings {current_year}-{tyfn}":=jo)

wb <- loadWorkbook(here("new_templates", "Job Openings by Industry_2016 Census_2023 LMO.xlsx"))
write_workbook(career_profiles, "Career Profiles", 5, 1)
write_workbook(`Job Openings by industry`, "Job Openings by industry", 4, 1)
saveWorkbook(wb, here(
  "out",
  paste0("Job Openings by Industry_2016 Census_",
         current_year,
         "LMO.xlsx")))



