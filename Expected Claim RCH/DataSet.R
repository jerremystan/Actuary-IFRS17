library(dplyr)
library(readr)
library(readxl)
library(writexl)

cohort <- 2023
lob_filter <- "ENGINEERING"

valuation <- "202312"

list_files <- list.files(path = paste0("input dataset/",valuation), pattern = ".xlsx")

case_ic_nm <- list_files[grepl("Case Reserve Insurance Contract", list_files)]
case_tfo_nm <- list_files[grepl("Case Reserve Non XL Treaty", list_files)]
case_fac_nm <- list_files[grepl("Case Reserve Non XL Facultative", list_files)]

paid_ic_nm <- list_files[grepl("Paid Loss Insurance Contract.20", list_files)]
paid_tfo_nm <- list_files[grepl("Paid Loss Non XL Treaty Or Facultative Obligatory.20", list_files)]
paid_fac_nm <- list_files[grepl("Paid Loss Non XL Facultative.20", list_files)]

paid_act_ic_nm <- list_files[grepl("Paid Loss Insurance Contract.Actual", list_files)]
paid_act_tfo_nm <- list_files[grepl("Paid Loss Non XL Treaty Or Facultative Obligatory.Actual", list_files)]
paid_act_fac_nm <- list_files[grepl("Paid Loss Non XL Facultative.Actual", list_files)]

rch_tfo_nm <- list_files[grepl(paste0("Obligatory Contract Helds.",cohort), list_files)]
rch_fac_nm <- list_files[grepl(paste0("Facultative Contract Helds.",cohort), list_files)]

rch_act_tfo_nm <- list_files[grepl(paste0("Obligatory Contract Helds.Actual.",cohort), list_files)]
rch_act_fac_nm <- list_files[grepl(paste0("Facultative Contract Helds.Actual.",cohort), list_files)]

ic_nm <- list_files[grepl(paste0("Contracts.",cohort), list_files)]
ic_act_nm <- list_files[grepl(paste0("ceContract.Actual.",cohort), list_files)]

case_ic <- read_excel(paste0("input dataset/",valuation,"/",case_ic_nm))
case_ic <- case_ic[case_ic$`Reserving Class` == lob_filter,]

case_tfo <- read_excel(paste0("input dataset/",valuation,"/",case_tfo_nm))
case_tfo <- case_tfo[case_tfo$`Reserving Class` == lob_filter,]

case_fac <- read_excel(paste0("input dataset/",valuation,"/",case_fac_nm))
case_fac <- case_fac[case_fac$`Reserving Class` == lob_filter,]

paid_fac <- read_excel(paste0("input dataset/",valuation,"/",paid_fac_nm))
coldate_paid_fac <- colnames(paid_fac)[grepl("Date", colnames(paid_fac))]
paid_fac_filter <- paid_fac[paid_fac$`Reserving Class` == lob_filter,]

paid_tfo <- read_excel(paste0("input dataset/",valuation,"/",paid_tfo_nm))
coldate_paid_tfo <- colnames(paid_tfo)[grepl("Date", colnames(paid_tfo))]
paid_tfo_filter <- paid_tfo[paid_tfo$`Reserving Class` == lob_filter,]

paid_ic <- read_excel(paste0("input dataset/",valuation,"/",paid_ic_nm))
coldate_paid_ic <- colnames(paid_ic)[grepl("Date", colnames(paid_ic))]
paid_ic_filter <- paid_ic[paid_ic$`Reserving Class` == lob_filter,]

paid_act_tfo <- read_excel(paste0("input dataset/",valuation,"/",paid_act_tfo_nm))
coldate_paid_act_tfo <- colnames(paid_act_tfo)[grepl("Date", colnames(paid_act_tfo))]

paid_act_fac <- read_excel(paste0("input dataset/",valuation,"/",paid_act_fac_nm))
coldate_paid_act_fac <- colnames(paid_act_fac)[grepl("Date", colnames(paid_act_fac))]

paid_act_ic <- read_excel(paste0("input dataset/",valuation,"/",paid_act_ic_nm))
coldate_paid_act_ic <- colnames(paid_act_ic)[grepl("Date", colnames(paid_act_ic))]

paid_ic_map <- data.frame(`Claim No` = paid_ic$`Claim No`
                          ,`Claim Transaction No` = paid_ic$`Claim Transaction No`
                          ,`BusinessTypeSID` = paid_ic$`Business Type Code`
                          ,`Insurer Code` = paid_ic$`Insurer Code`
                          ,`ReinsuranceTypeCode` = paid_ic$`Reinsurance Type Code`
                          ,`Reserving Class` = paid_ic$`Reserving Class`
                          # ,`IFRS Group Code` = paid_ic$`IFRS Group Code`
                          ,check.names = FALSE
                          )
paid_ic_map <- unique(paid_ic_map)
paid_ic_map$BusinessTypeSID <- as.numeric(paid_ic_map$BusinessTypeSID)

paid_tfo_map <- data.frame(`Claim No` = paid_tfo$`Claim No`
                          ,`Claim Transaction No` = paid_tfo$`Claim Transaction No`
                          ,`BusinessTypeSID` = paid_tfo$`Business Type Code`
                          ,`ReinsuranceTypeCode` = paid_tfo$`Reinsurance Type Code`
                          ,`Reinsurance Code` = paid_tfo$`Reinsurance Code`
                          ,`Participant Code` = paid_tfo$`Participant Code`
                          ,`Reserving Class` = paid_tfo$`Reserving Class`
                          # ,`IFRS Group Code` = paid_tfo$`IFRS Group Code`
                          ,check.names = FALSE
)
paid_tfo_map <- unique(paid_tfo_map)
paid_tfo_map$BusinessTypeSID <- as.numeric(paid_tfo_map$BusinessTypeSID)


paid_fac_map <- data.frame(`Claim No` = paid_fac$`Claim No`
                           ,`Claim Transaction No` = paid_fac$`Claim Transaction No`
                           ,`BusinessTypeSID` = paid_fac$`Business Type Code`
                           ,`ReinsuranceTypeCode` = paid_fac$`Reinsurance Type Code`
                           ,`Reinsurance Code` = paid_fac$`Reinsurance Code`
                           ,`Reserving Class` = paid_fac$`Reserving Class`
                           # ,`IFRS Group Code` = paid_fac$`IFRS Group Code`
                           ,check.names = FALSE
)
paid_fac_map <- unique(paid_fac_map)
paid_fac_map$BusinessTypeSID <- as.numeric(paid_fac_map$BusinessTypeSID)

paid_act_tfo_final <- left_join(paid_act_tfo, paid_tfo_map, by = c("Claim No", "Claim Transaction No", "BusinessTypeSID", "ReinsuranceTypeCode", "Reinsurance Code", "Participant Code"))

paid_act_fac_final <- left_join(paid_act_fac, paid_fac_map, by = c("Claim No", "Claim Transaction No", "BusinessTypeSID", "ReinsuranceTypeCode", "Reinsurance Code"))

paid_act_ic_final <- left_join(paid_act_ic, paid_ic_map, by = c("Claim No", "Claim Transaction No", "BusinessTypeSID", "Insurer Code", "ReinsuranceTypeCode"))

paid_act_tfo_final_filter <- paid_act_tfo_final[paid_act_tfo_final$`Reserving Class` == lob_filter & !is.na(paid_act_tfo_final$`Claim No`),]
paid_act_fac_final_filter <- paid_act_fac_final[paid_act_fac_final$`Reserving Class` == lob_filter & !is.na(paid_act_fac_final$`Claim No`),]
paid_act_ic_final_filter <- paid_act_ic_final[paid_act_ic_final$`Reserving Class` == lob_filter & !is.na(paid_act_ic_final$`Claim No`),]

rch_tfo <- read_excel(paste0("input dataset/",valuation,"/",rch_tfo_nm))
coldate_rch_tfo <- colnames(rch_tfo)[grepl("Date", colnames(rch_tfo))]

rch_fac <- read_excel(paste0("input dataset/",valuation,"/",rch_fac_nm))
coldate_rch_fac <- colnames(rch_fac)[grepl("Date", colnames(rch_fac))]

ic <- read_excel(paste0("input dataset/",valuation,"/",ic_nm))
coldate_ic <- colnames(ic)[grepl("Date", colnames(ic))]

rch_act_tfo <- read_excel(paste0("input dataset/",valuation,"/",rch_act_tfo_nm))
coldate_rch_act_tfo <- colnames(rch_act_tfo)[grepl("Date", colnames(rch_act_tfo))]
rch_act_tfo$`Business Type Code` <- as.character(rch_act_tfo$`Business Type Code`)

rch_act_fac <- read_excel(paste0("input dataset/",valuation,"/",rch_act_fac_nm))
coldate_rch_act_fac <- colnames(rch_act_fac)[grepl("Date", colnames(rch_act_fac))]
rch_act_fac$`Business Type Code` <- as.character(rch_act_fac$`Business Type Code`)

ic_act <- read_excel(paste0("input dataset/",valuation,"/",ic_act_nm))
coldate_ic_act <- colnames(ic_act)[grepl("Date", colnames(ic_act))]
ic_act$`Business Type Code` <- as.character(ic_act$`Business Type Code`)

rch_tfo_filter <- rch_tfo[rch_tfo$`Reserving Class` == lob_filter,]

rch_fac_filter <- rch_fac[rch_fac$`Reserving Class` == lob_filter,]

ic_filter <- ic[ic$`Reserving Class` == lob_filter,]

colnames(rch_act_tfo)[1:2] <- c("PolicyNo", "PolicyTransactionNo")
colnames(rch_act_fac)[1:2] <- c("PolicyNo", "PolicyTransactionNo")
colnames(ic_act)[1:2] <- c("PolicyNo", "PolicyTransactionNo")

rch_tfo_map <- data.frame(PolicyNo = rch_tfo$PolicyNo
                             ,PolicyTransactionNo = rch_tfo$PolicyTransactionNo
                             ,`Reinsurance Code` = rch_tfo$`Reinsurance Code`
                             ,`Reinsurance Type Code` = rch_tfo$`Reinsurance Type Code`
                             ,`Business Type Code` = rch_tfo$`Business Type Code`
                             ,`Participant Code` = rch_tfo$`Participant Code`
                             ,`Reserving Class` = rch_tfo$`Reserving Class`
                             #,`IFRS Group Code` = rch_tfo$`IFRS Group Code`
                      , check.names = FALSE)
rch_tfo_map <- unique(rch_tfo_map)

rch_fac_map <- data.frame(PolicyNo = rch_fac$PolicyNo
                      ,PolicyTransactionNo = rch_fac$PolicyTransactionNo
                      ,`Reinsurance Code` = rch_fac$`Reinsurance Code`
                      ,`Business Type Code` = rch_fac$`Business Type Code`
                      ,`Reserving Class` = rch_fac$`Reserving Class`
                      #,`IFRS Group Code` = rch_fac$`IFRS Group Code`
                      , check.names = FALSE)
rch_fac_map <- unique(rch_fac_map)
# 
# rch_fac_act_check <- data.frame(PolicyNo = rch_act_fac$PolicyNo
#                                 ,PolicyTransactionNo = rch_act_fac$PolicyTransactionNo
#                                 ,`Reinsurance Code` = rch_act_fac$`Reinsurance Code`
#                                 ,`Business Type Code` = rch_act_fac$`Business Type Code`
#                                 ,check.names = FALSE)
# 
# rch_fac_act_check <- rch_fac_act_check %>%
#   mutate(Map = paste0(PolicyNo, PolicyTransactionNo, `Reinsurance Code`, `Business Type Code`))
# 
# rch_fac_final_check <- rch_fac_final_check
# rch_fac_final_check <- rch_act_fac_final %>%
#   mutate(Map = paste0(PolicyNo, PolicyTransactionNo, `Reinsurance Code`, `Business Type Code`))
# 
# result <- rch_fac_act_check %>%
#   rowwise() %>%
#   mutate(count = sum(rch_fac_final_check$Map == Map))
# 
# filter_result <- result[result$count > 1,]
# 
ic_map <- unique(data.frame(PolicyNo = ic$PolicyNo
                            ,PolicyTransactionNo = ic$PolicyTransactionNo
                            ,`Reinsurance Type Code` = ic$`Reinsurance Type Code`
                            ,`Business Type Code` = ic$`Business Type Code`
                            ,`Reserving Class` = ic$`Reserving Class`
                            ,`IFRS Group Code` = ic$`IFRS Group Code`, check.names = FALSE))
ic_map <- unique(ic_map)

rch_act_tfo_final <- left_join(rch_act_tfo, rch_tfo_map, by = c("PolicyNo", "PolicyTransactionNo", "Reinsurance Code", "Reinsurance Type Code", "Business Type Code", "Participant Code"))

rch_act_fac_final <- left_join(rch_act_fac, rch_fac_map, by = c("PolicyNo", "PolicyTransactionNo", "Reinsurance Code", "Business Type Code"))

ic_act_final <- left_join(ic_act, ic_map, by = c("PolicyNo", "PolicyTransactionNo", "Reinsurance Type Code", "Business Type Code"))

rch_act_tfo_final_filter <- rch_act_tfo_final[rch_act_tfo_final$`Reserving Class` == lob_filter,]
rch_act_fac_final_filter <- rch_act_fac_final[rch_act_fac_final$`Reserving Class` == lob_filter,]
ic_act_final_filter <- ic_act_final[(ic_act_final$`Reserving Class` == lob_filter),]


for (col in coldate_paid_tfo) {
  
  paid_tfo_filter[[col]] <- as.Date(paid_tfo_filter[[col]], format = "%Y-%m-%d")
  
}

for (col in coldate_paid_fac) {
  
  paid_fac_filter[[col]] <- as.Date(paid_fac_filter[[col]], format = "%Y-%m-%d")
  
}

for (col in coldate_paid_ic) {
  
  paid_ic_filter[[col]] <-as.Date(paid_ic_filter[[col]], format = "%Y-%m-%d")
  
}

for (col in coldate_paid_act_tfo) {
  
  paid_act_tfo_final_filter[[col]] <- as.character(as.Date(paid_act_tfo_final_filter[[col]], format = "%Y-%m-%d"))
  
}

for (col in coldate_paid_act_fac) {
  
  paid_act_fac_final_filter[[col]] <- as.character(as.Date(paid_act_fac_final_filter[[col]], format = "%Y-%m-%d"))
  
}

for (col in coldate_paid_act_ic) {
  
  paid_act_ic_final_filter[[col]] <- as.character(as.Date(paid_act_ic_final_filter[[col]], format = "%Y-%m-%d"))
  
}

for (col in coldate_rch_tfo) {
  
  rch_tfo_filter[[col]] <- as.Date(rch_tfo_filter[[col]], format = "%Y-%m-%d")
  
}

for (col in coldate_rch_fac) {
  
  rch_fac_filter[[col]] <- as.Date(rch_fac_filter[[col]], format = "%Y-%m-%d")
  
}

for (col in coldate_ic) {
  
  ic_filter[[col]] <- as.Date(ic_filter[[col]], format = "%Y-%m-%d")

}

for (col in coldate_rch_act_tfo) {
  
  rch_act_tfo_final_filter[[col]] <- as.character(as.Date(rch_act_tfo_final_filter[[col]], format = "%Y-%m-%d"))

}

for (col in coldate_rch_act_fac) {
  
  rch_act_fac_final_filter[[col]] <- as.character(as.Date(rch_act_fac_final_filter[[col]], format = "%Y-%m-%d"))
  
}

for (col in coldate_ic_act) {
  
  ic_act_final_filter[[col]] <- as.character(as.Date(ic_act_final_filter[[col]], format = "%Y-%m-%d"))

}

dir.create(paste0("Output/",valuation))

paid_act_fac_final_filter <- paid_act_fac_final_filter[!is.na(paid_act_fac_final_filter$`Claim No`),]
paid_act_tfo_final_filter <- paid_act_tfo_final_filter[!is.na(paid_act_tfo_final_filter$`Claim No`),]
paid_act_ic_final_filter <- paid_act_ic_final_filter[!is.na(paid_act_ic_final_filter$`Claim No`),]

rch_act_fac_final_filter <- rch_act_fac_final_filter[!is.na(rch_act_fac_final_filter$PolicyNo),]
rch_act_tfo_final_filter <- rch_act_tfo_final_filter[!is.na(rch_act_tfo_final_filter$PolicyNo),]
ic_act_final_filter <- ic_act_final_filter[!is.na(ic_act_final_filter$PolicyNo),]

write_xlsx(case_ic, paste0("Output/",valuation,"/",lob_filter,"-",case_ic_nm,".xlsx"))
write_xlsx(case_tfo, paste0("Output/",valuation,"/",lob_filter,"-",case_tfo_nm,".xlsx"))
write_xlsx(case_fac, paste0("Output/",valuation,"/",lob_filter,"-",case_fac_nm,".xlsx"))

write_xlsx(paid_ic_filter, paste0("Output/",valuation,"/",lob_filter,"-",paid_ic_nm,".xlsx"))
write_xlsx(paid_tfo_filter, paste0("Output/",valuation,"/",lob_filter,"-",paid_tfo_nm,".xlsx"))
write_xlsx(paid_fac_filter, paste0("Output/",valuation,"/",lob_filter,"-",paid_fac_nm,".xlsx"))

write_xlsx(paid_act_ic_final_filter, paste0("Output/",valuation,"/",lob_filter,"-",paid_act_ic_nm,".xlsx"))
write_xlsx(paid_act_tfo_final_filter, paste0("Output/",valuation,"/",lob_filter,"-",paid_act_tfo_nm,".xlsx"))
write_xlsx(paid_act_fac_final_filter, paste0("Output/",valuation,"/",lob_filter,"-",paid_act_fac_nm,".xlsx"))

write_xlsx(rch_tfo_filter, paste0("Output/",valuation,"/",lob_filter,"-",rch_tfo_nm,".xlsx"))
write_xlsx(rch_fac_filter, paste0("Output/",valuation,"/",lob_filter,"-",rch_fac_nm,".xlsx"))
write_xlsx(ic_filter, paste0("Output/",valuation,"/",lob_filter,"-",ic_nm,".xlsx"))

write_xlsx(rch_act_tfo_final_filter, paste0("Output/",valuation,"/",lob_filter,"-",rch_act_tfo_nm,".xlsx"))
write_xlsx(rch_act_fac_final_filter, paste0("Output/",valuation,"/",lob_filter,"-",rch_act_fac_nm,".xlsx"))
write_xlsx(ic_act_final_filter, paste0("Output/",valuation,"/",lob_filter,"-",ic_act_nm,".xlsx"))

