library(dplyr)
library(tidyverse)
library(readr)
library(readxl)
library(writexl)
library(openxlsx)

valuation <- "202312"
cohort <- 2023
initial_period <- "2022-12-31"

report_period <- "2023-12-31"

ULR_files <- read_excel("Assumption/ULR.v.28.xlsx", sheet = "ULR")
ULR_init <- ULR_files[ULR_files$`Valuation Date` == initial_period & ULR_files$`Cohort Year` == cohort,]
ULR_IC <- ULR_init
colnames(ULR_IC)[4] <- "IsShortTermIC"

NPR_files <- read_excel("Assumption/NonPerformanceRisks.v.3.xlsx")
NPR_init <- NPR_files[NPR_files$`Valuation Date` == initial_period,]
colnames(NPR_init)[5:6] <- c("IsShortTerm", "Cohort")
NPR_init <- NPR_init[,-1]

NDIC_files <- read_excel("Assumption/NDIC for RCH v.2.xlsx")
NDIC_filter <- NDIC_files[NDIC_files$`NDIC Type` != "NA",]
colnames(NDIC_filter)[4] <- "Reinsurance Code"


lob_filter <- ""

list_files <- list.files(path = paste0("input dataset/",valuation), pattern = ".xlsx")

ic_nm <- list_files[grepl(paste0("Contracts.",cohort), list_files)]
rch_tfo_nm <- list_files[grepl(paste0("Obligatory Contract Helds.",cohort), list_files)]
rch_fac_nm <- list_files[grepl(paste0("Facultative Contract Helds.",cohort), list_files)]

ic <- read_excel(paste0("input dataset/",valuation,"/",ic_nm))
coldate_ic <- colnames(ic)[grepl("Date", colnames(ic))]
ic$IsShortTerm <- ifelse(ic$`ST/LT` == "ST", 1, 0)

rch_tfo <- read_excel(paste0("input dataset/",valuation,"/",rch_tfo_nm))
coldate_rch_tfo <- colnames(rch_tfo)[grepl("Date", colnames(rch_tfo))]
rch_tfo$IsShortTerm <- ifelse(rch_tfo$`ST/LT` == "ST", 1, 0)

rch_fac <- read_excel(paste0("input dataset/",valuation,"/",rch_fac_nm))
coldate_rch_fac <- colnames(rch_fac)[grepl("Date", colnames(rch_fac))]
rch_fac$IsShortTerm <- ifelse(rch_fac$`ST/LT` == "ST", 1, 0)

ic_group <- ic %>%
  group_by(PolicyNo, PolicyTransactionNo, `Reserving Class Code`, IsShortTerm, `Transaction Issue Date`) %>%
  summarise(`Gross Premium` = sum(Premium))

ic_group <- left_join(ic_group, ULR_init, by = c("Reserving Class Code","IsShortTerm"))
ic_select <- ic_group %>%
  ungroup %>%
  select(PolicyNo, PolicyTransactionNo, `IC Transaction Issue Date` = `Transaction Issue Date`, `Gross Premium`, `Gross ULR`, `Net wo XL ULR`) %>%
  filter(`IC Transaction Issue Date` <= as.Date(report_period))

ic_select$`Expected Gross Claim` <- ic_select$`Gross Premium`*ic_select$`Gross ULR`/100
ic_select$`Expected Net wo XL Claim` <- ic_select$`Gross Premium`*ic_select$`Net wo XL ULR`/100

#Sum of TFO and FAC

tfo_prem <- rch_tfo %>%
  group_by(PolicyNo, PolicyTransactionNo, `Reinsurance Code`, `Participant Code`, `Reserving Class Code`, IsShortTerm, `IFRS Group Code`, `Transaction Issue Date`) %>%
  summarise(`Participant Premium` = sum(Premium)
            ,AGRICommission = sum(AGRICommission)
            ,OCPORCommission = sum(OCPORCommission)) %>%
  mutate(`IsShortTermIC` = ifelse(grepl("#ST", `IFRS Group Code`),1,0)) %>%
  left_join(ULR_IC, by = c("Reserving Class Code","IsShortTermIC")) %>%
  mutate(`Participant Expected Net wo XL Claim` = `Participant Premium` * `Net wo XL ULR`/100,
         `RI Type` = "Treaty",
         `Is Catastrophe XL` = 0,
         `Is Treaty Or Facultative Obligatory` = 1,
         #`Cohort` = cohort,
         `Is Initial` = 0) %>%
  left_join(NPR_init, by = c("Reserving Class Code","Is Catastrophe XL","Is Treaty Or Facultative Obligatory","IsShortTerm","Is Initial")) %>%
  filter(`Transaction Issue Date` <= as.Date(report_period))

tfo_prem_ttl <- tfo_prem %>%
  group_by(PolicyNo, PolicyTransactionNo) %>%
  summarise(`Total RI Premium` = sum(`Participant Premium`),
            `Total Expected RI Net wo XL Claim` = sum(`Participant Expected Net wo XL Claim`))

fac_prem <- rch_fac %>%
  group_by(PolicyNo, PolicyTransactionNo, `Reinsurance Code`, `Participant Code`, `Reserving Class Code`, IsShortTerm, `IFRS Group Code`, `Transaction Issue Date`) %>%
  summarise(`Participant Premium` = sum(Premium)
            ,AGRICommission = sum(AGRICommission)
            ,OCPORCommission = sum(OCPORCommission)) %>%
  mutate(`IsShortTermIC` = ifelse(grepl("#ST", `IFRS Group Code`),1,0)) %>%
  left_join(ULR_IC, by = c("Reserving Class Code","IsShortTermIC")) %>%
  mutate(`Participant Expected Net wo XL Claim` = `Participant Premium` * `Net wo XL ULR`/100,
         `RI Type` = "Facultative",
         `Is Catastrophe XL` = 0,
         `Is Treaty Or Facultative Obligatory` = 0,
         #`Cohort` = cohort,
         `Is Initial` = 0) %>%
  left_join(NPR_init, by = c("Reserving Class Code","Is Catastrophe XL","Is Treaty Or Facultative Obligatory","IsShortTerm","Is Initial")) %>%
  filter(`Transaction Issue Date` <= as.Date(report_period))
  
fac_prem_ttl <- fac_prem %>%
  group_by(PolicyNo, PolicyTransactionNo) %>%
  summarise(`Total RI Premium` = sum(`Participant Premium`),
            `Total Expected RI Net wo XL Claim` = sum(`Participant Expected Net wo XL Claim`))

#Stack and Group By the TFO and FAC
prem_ttl <- rbind(tfo_prem_ttl, fac_prem_ttl) %>%
  group_by(PolicyNo, PolicyTransactionNo) %>%
  summarise(`Total RI Premium` = sum(`Total RI Premium`),
            `Total Expected RI Net wo XL Claim` = sum(`Total Expected RI Net wo XL Claim`))

ic_new <- full_join(ic_select, prem_ttl, by = c("PolicyNo","PolicyTransactionNo"))

ic_new$`Total RI Premium` <- replace(ic_new$`Total RI Premium`, is.na(ic_new$`Total RI Premium`),0)
ic_new$`Total Expected RI Net wo XL Claim` <- replace(ic_new$`Total Expected RI Net wo XL Claim`, is.na(ic_new$`Total Expected RI Net wo XL Claim`),0)

ic_tfo <- left_join(ic_new, tfo_prem, by = c("PolicyNo","PolicyTransactionNo"))
#disini harus dihapus yang NA
ic_tfo <- ic_tfo[!is.na(ic_tfo$`RI Type`),]

ic_fac <- left_join(ic_new, fac_prem, by = c("PolicyNo","PolicyTransactionNo"))
#disini harus dihapus yang NA
ic_fac <- ic_fac[!is.na(ic_fac$`RI Type`),]

ic_final <- union(ic_tfo, ic_fac)

colnames(ic_final) <- gsub("Gross ULR.x","% Gross LR",colnames(ic_final))
colnames(ic_final) <- gsub("Net wo XL ULR.x","% Net WO XL LR",colnames(ic_final))
colnames(ic_final) <- gsub("Gross ULR.y","Participant % Gross LR",colnames(ic_final))
colnames(ic_final) <- gsub("Net wo XL ULR.y","Participant % Net WO XL LR",colnames(ic_final))

ic_final$`Net Premium` <- ic_final$`Gross Premium` - ic_final$`Total RI Premium`
ic_final$`Expected RCH`<- ic_final$`Expected Gross Claim`-(ic_final$`Expected Net wo XL Claim`-ic_final$`Total Expected RI Net wo XL Claim`)

ic_final$`IC Expected Gross Claim` <- (ic_final$`Participant Premium`/ic_final$`Total RI Premium`)*ic_final$`Expected Gross Claim`
ic_final$`IC Expected Gross Claim` <- replace(ic_final$`IC Expected Gross Claim`, is.na(ic_final$`IC Expected Gross Claim`),0)

ic_final$`IC Expected Net WO XL Claim` <- (ic_final$`Participant Premium`/ic_final$`Total RI Premium`)*ic_final$`Expected Net wo XL Claim`
ic_final$`IC Expected Net WO XL Claim` <- replace(ic_final$`IC Expected Net WO XL Claim`, is.na(ic_final$`IC Expected Net WO XL Claim`),0)

ic_final$`Participant Expected Net wo XL Claim` <- replace(ic_final$`Participant Expected Net wo XL Claim`, is.na(ic_final$`Participant Expected Net wo XL Claim`),0)
ic_final$`RCH Expected Claim` <- ic_final$`IC Expected Gross Claim`-(ic_final$`IC Expected Net WO XL Claim`-ic_final$`Participant Expected Net wo XL Claim`)

a_ic_rch_df <- ic_final[!is.na(ic_final$`IC Transaction Issue Date`) & !is.na(ic_final$`RI Type`),]
a_ic_only_df <- ic_final[!is.na(ic_final$`IC Transaction Issue Date`) & is.na(ic_final$`RI Type`),]
a_rch_only_df <- ic_final[is.na(ic_final$`IC Transaction Issue Date`),]

ic_rch_df <- a_ic_rch_df %>%
  group_by(PolicyNo, PolicyTransactionNo, `Gross Premium`, `Total RI Premium`, `Expected RCH`) %>%
  summarise(`RCH Expected Claim` = sum(`RCH Expected Claim`))

a_rch_only_df$`Gross Premium` <- 0
rch_only_df <- a_rch_only_df %>%
  group_by(PolicyNo, PolicyTransactionNo, `Gross Premium`, `Expected RCH`) %>%
  summarise(`RCH Expected Claim` = sum(`RCH Expected Claim`)) %>%
  left_join(ic_new, by = c("PolicyNo","PolicyTransactionNo")) %>%
  select(PolicyNo, PolicyTransactionNo, `Gross Premium` =`Gross Premium.x`, `Total RI Premium`, `Expected RCH`, `RCH Expected Claim`)

ic_summary <- rbind(ic_rch_df, rch_only_df)

ic_ndic <- left_join(ic_final, NDIC_filter, by = c("Reinsurance Code"))

#CEK ic_final dan ic_ndic sama atau tidak

ic_final <- ic_ndic %>%
  select(PolicyNo
         ,PolicyTransactionNo
         ,`IC Transaction Issue Date`
         ,`Gross Premium`
         ,`% Gross LR`
         ,`% Net WO XL LR`
         ,`Expected Gross Claim`
         ,`Expected Net wo XL Claim`
         ,`Total RI Premium`
         ,`Net Premium`
         ,`Total Expected RI Net wo XL Claim`
         ,`Expected RCH`
         ,`RI Type`
         ,`Reinsurance Code`
         ,`Participant Code`
         ,`RCH Transaction Issue Date` = `Transaction Issue Date`
         ,`Participant Premium`
         ,`Participant % Gross LR`
         ,`Participant % Net WO XL LR`
         ,`Participant Expected Net wo XL Claim`
         ,`Non Performance Risks`
         ,`IFRS Group Code`
         ,AGRICommission
         ,OCPORCommission
         ,`Profit Commission`=`Profit Commisson`
         ,`Reinsurer Expenses`
         )

ic_final <- unique(ic_final)

ic_final <- ic_final[!is.na(ic_final$`RI Type`),]

ic_final$`IC Expected Gross Claim` <- NA
ic_final$`IC Expected Net WO XL Claim` <- NA
ic_final$`RCH Expected Claim` <- NA
ic_final$NDIC <- NA

ic_final <- ic_final %>%
  arrange(PolicyNo, PolicyTransactionNo)
ic_summary <- ic_summary %>%
  arrange(PolicyNo, PolicyTransactionNo)

wb <- createWorkbook()

addWorksheet(wb,"Summary")
addWorksheet(wb,"Detail")

writeData(wb,sheet = "Summary", x = ic_summary, startCol = 1, startRow = 1)
writeData(wb,sheet = "Detail", x = ic_final, startCol = 1, startRow = 1)

saveWorkbook(wb, paste0("RCH Expected Claim & NDIC.",valuation,".v.1.xlsx"), overwrite = TRUE)
#write_xlsx(ic_final,"RCH Expected Claim.v.1.xlsx")
