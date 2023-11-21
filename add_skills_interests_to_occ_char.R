library(readxl)
library(here)
library(tidyverse)
library(openxlsx)

get_skill_names <- function(tbbl){
  on_own <- tbbl$`Skills & Competencies`
  together <- paste(on_own, collapse = ", ")
  temp <- as_tibble(t(c(on_own, together)), .name_repair = "minimal")
  colnames(temp) <- c("Skill1", "Skill2", "Skill3", "Skills: Top 3")
  temp
}

get_interests_names <- function(tbbl){
  tbbl <- tbbl|>
    na.omit()
  together <- paste(tbbl$`Occupational Interest`, collapse = ", ")
  on_own <- str_split_1(together, pattern = ", ")
  length(on_own) <- 3 #pads out to length 3
  temp <- as_tibble(t(c(on_own, together)), .name_repair = "minimal")
  colnames(temp) <- c("Interest1", "Interest2", "Interest3", "Interests")
  temp
}

skills_df <- read_excel(here("data","Top skills by NOC2021 occupations.xlsx"))|>
  group_by(NOC=NOC2021)|>
  mutate(NOC=paste0("#",NOC))|>
  slice_max(`Importance Score`, n=3, with_ties = FALSE)|>
  nest()|>
  mutate(skills=map(data, get_skill_names))|>
  select(-data)|>
  unnest(skills)

skills_df|>
  write.xlsx(here("out","skills.xlsx"))

interests_df <- read_excel(here("data","Occupational interest by NOC2021 occupation.xlsx"))|>
  select(NOC=`NOC 2021`, `Occupational Interest`)|>
  mutate(NOC=paste0("#",NOC))|>
  group_by(NOC)|>
  nest()|>
  mutate(interests=map(data, get_interests_names))|>
  select(-data)|>
  unnest(interests)

interests_df|>
 write.xlsx(here("out",
                 paste("Occupational Interests_",
                       year(today()),
                       ".xlsx")))

occ_char <- read_excel(here("data",
                            list.files(here("data"),
                                       pattern = "Occupational Characteristics based")),
                       skip=3)|>
  left_join(skills_df)|>
  left_join(interests_df)

openxlsx::write.xlsx(occ_char, here("data","Occupational Characteristics (with skills and interests).xlsx"), asTable = TRUE)




