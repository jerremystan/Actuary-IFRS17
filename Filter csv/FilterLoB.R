library(dplyr)
library(readr)

valuation <- "202312"

lrc_1 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_1.csv"), check.names = FALSE)
# lrc_2 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_2.csv"), check.names = FALSE)
# lrc_3 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_3.csv"), check.names = FALSE)
# lrc_4 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_4.csv"), check.names = FALSE)
# lrc_5 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_5.csv"), check.names = FALSE)
# lrc_6 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_6.csv"), check.names = FALSE)
# lrc_7 <- read.csv(paste0("input csv/", valuation,"/ifrs_group_cashflow_all_runs_2023_",valuation,"_7.csv"), check.names = FALSE)

ifrs_cf <- rbind(lrc_1
                 # , lrc_2, lrc_3, lrc_4, lrc_5
                 # , lrc_6
                 # , lrc_7
                 )

var_1 <- read.csv(paste0("input csv/", valuation,"/variable_all_runs_2023_",valuation,"_1.csv"), check.names = FALSE)
# var_2 <- read.csv("input csv/variable_all_runs_2023_202312_2.csv")
# var_3 <- read.csv("input csv/variable_all_runs_2023_202312_3.csv")
# var_4 <- read.csv("input csv/variable_all_runs_2023_202312_4.csv")
# 
var_cf <- rbind(var_1
          #, var_2, var_3, var_4
          )
# 
# disc_1 <- read.csv("input csv/ifrs_group_discount_rate_all_runs_2023_202312_1.csv")
# disc_2 <- read.csv("input csv/ifrs_group_discount_rate_all_runs_2023_202312_2.csv")
# disc_3 <- read.csv("input csv/ifrs_group_discount_rate_all_runs_2023_202312_3.csv")
# disc_4 <- read.csv("input csv/ifrs_group_discount_rate_all_runs_2023_202312_4.csv")
# 
# disc_cf <- rbind(disc_1,disc_2, disc_3, disc_4)

filter_1 <- "FAC"
filter_2 <- "ENGINEERING"

filter <- c(filter_1, filter_2)

filtered_ifrs <- ifrs_cf[grepl(filter_1,ifrs_cf$ifrs_group_code) & grepl(filter_2,ifrs_cf$ifrs_group_code),]

ifrs_group_IC <- unique(filtered_ifrs$underlying_ifrs_group_code)

filtered_IC <- ifrs_cf[ifrs_cf$ifrs_group_code %in% ifrs_group_IC,]

write_csv(filtered_IC, paste0("1.UnderlyingCF","-",filter_1,"-",filter_2,".csv"), na = "")
write_csv(filtered_ifrs, paste0("1.ReinsuranceCF","-",filter_1,"-",filter_2,".csv"), na = "")


# filtered_var <- var_cf[grepl(filter_1, var_cf$ifrs_group_code) & grepl(filter_2, var_cf$ifrs_group_code),]
filtered_var <- var_cf[grepl(filter_1,var_cf$ifrs_group_code) & grepl(filter_2,var_cf$ifrs_group_code),]
# filtered_disc <- disc_cf[grepl(filter_1, disc_cf$ifrs_group_code) & grepl(filter_2, disc_cf$ifrs_group_code),]
write_csv(filtered_var, "2.VAR.csv", na = "")
# write_csv(filtered_disc, "3.DISC.csv", na = "")
