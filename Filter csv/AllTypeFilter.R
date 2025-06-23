library(dplyr)
library(readr)
library(openxlsx)

valuation <- "202312"
filter <- "FIRE"
lrc_1 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_1.csv"), check.names = FALSE)
# lrc_2 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_2.csv"), check.names = FALSE)
# lrc_3 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_3.csv"), check.names = FALSE)
# lrc_4 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_4.csv"), check.names = FALSE)
# lrc_5 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_5.csv"), check.names = FALSE)
# lrc_6 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_6.csv"), check.names = FALSE)
# lrc_7 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_7.csv"), check.names = FALSE)
#lrc_8 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_8.csv"), check.names = FALSE)

ifrs_cf <- rbind(lrc_1
                 #, lrc_2, lrc_3, lrc_4, lrc_5
                 #, lrc_6
                 #, lrc_7
                 #, lrc_8
)

ifrs_cf_filt <- ifrs_cf[grepl(filter,ifrs_cf$ifrs_group_code),]

var_1 <- read.csv(paste0("input csv/", valuation,"/variable_all_runs_2023_",valuation,"_1.csv"), check.names = FALSE)
# var_2 <- read.csv(paste0("input csv/", valuation,"/variable_all_runs_2023_",valuation,"_2.csv"), check.names = FALSE)
# var_3 <- read.csv(paste0("input csv/", valuation,"/variable_all_runs_2023_",valuation,"_3.csv"), check.names = FALSE)
# var_4 <- read.csv(paste0("input csv/", valuation,"/variable_all_runs_2023_",valuation,"_4.csv"), check.names = FALSE)
# 
var_cf <- rbind(var_1
                #, var_2, var_3, var_
)

var_cf_filt <- var_cf[grepl(filter, ifrs_cf$ifrs_group_code),]

ifrs_cf_filt$custom_dim_1[ifrs_cf_filt$custom_dim_1 == ""] <- NA
var_cf_filt$custom_dim_1[var_cf_filt$custom_dim_1 == ""] <- NA

wb <- createWorkbook()

addWorksheet(wb, "IFRS")
addWorksheet(wb, "VAR")

writeData(wb, "IFRS", ifrs_cf_filt, startCol = 1, startRow = 1)
writeData(wb, "VAR", var_cf_filt, startCol = 1, startRow = 1)

saveWorkbook(wb, paste0("Cashflow-",filter,".xlsx"), overwrite = TRUE)
