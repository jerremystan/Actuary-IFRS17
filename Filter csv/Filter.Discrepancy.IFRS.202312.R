library(readr)
library(dplyr)
library(writexl)
library(readxl)


ifrs_group_name <- "TFO#FIRE#LT#2021#N#TMIS#RI#FIRE#ST#2023#P#IDR#F#EAQ#B1#REC#L#OTHERS"

ifrs_cf <- read.csv("Input/ifrs_group_cashflow_all_runs_2023_202312_1.csv", check.names = FALSE)
df_ifrs_cf <- ifrs_cf[ifrs_cf$ifrs_group_code == ifrs_group_name,]

direct <- unique(df_ifrs_cf$underlying_ifrs_group_code)

df_ifrs_cf <- ifrs_cf[ifrs_cf$ifrs_group_code %in% c(ifrs_group_name,direct),]

ifrs_group <- read.csv("Input/ifrs_group_2023_202312_1.csv", check.names = FALSE)
df_ifrs_group <- ifrs_group[ifrs_group$code  %in% c(ifrs_group_name,direct),]



portfolio_name <- unique(df_ifrs_group$portfolio_name)



disc_1 <- read.csv("Input/ifrs_group_discount_rate_all_runs_2023_202312_1.csv", check.names = FALSE)
disc_2 <- read.csv("Input/ifrs_group_discount_rate_all_runs_2023_202312_2.csv", check.names = FALSE)
disc_3 <- read.csv("Input/ifrs_group_discount_rate_all_runs_2023_202312_3.csv", check.names = FALSE)
disc_4 <- read.csv("Input/ifrs_group_discount_rate_all_runs_2023_202312_4.csv", check.names = FALSE)
disc_5 <- read.csv("Input/ifrs_group_discount_rate_all_runs_2023_202312_5.csv", check.names = FALSE)

disc <- rbind(disc_1
              , disc_2
              , disc_3
              , disc_4
              , disc_5
              )

disc_df <- disc[disc$ifrs_group_code  %in% c(ifrs_group_name,direct),]

variable <- read.csv("Input/variable_all_runs_2023_202312_1.csv", check.names = FALSE)
df_variable <- variable[variable$ifrs_group_code  %in% c(ifrs_group_name,direct),]

variable_cf <- read.csv("Input/variable_cashflow_mapping_1.csv", check.names = FALSE)
df_variable_mp <- variable_cf[variable_cf$portfolio_name %in% portfolio_name,]

write.csv(df_ifrs_group, "Output/ifrs_group_2023_202312_1.csv", row.names = FALSE, na = "")
write.csv(df_ifrs_cf, "Output/ifrs_group_cashflow_all_runs_2023_202312_1.csv", row.names = FALSE, na = "")
write.csv(disc_df, "Output/ifrs_group_discount_rate_all_runs_2023_202312_1.csv", row.names = FALSE, na = "")
write.csv(df_variable, "Output/variable_all_runs_2023_202312_1.csv", row.names = FALSE, na = "")
write.csv(df_variable_mp, "Output/variable_cashflow_mapping_1.csv", row.names = FALSE, na = "")
