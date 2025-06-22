library(dplyr)
library(readxl)
library(tidyr)
library(lubridate)

data_A <- read.csv("D:/Stanley/Tugas IFRS17/Reconcile Retro/Cek input RiskIntegrity/ifrs_group_cashflow_all_runs_202301.csv")
data_B <- read.csv("D:/Stanley/Tugas IFRS17/Reconcile Retro/Cek input RiskIntegrity/ifrs_group_cashflow_all_runs_202302.csv")

col_A <- colnames(data_A[,11:ncol(data_A)])

data_A_long <- pivot_longer(data = data_A,
                            cols = col_A,
                            names_to = "DueDate",
                            values_to = "AmountA")

data_A_long$DueDate <- gsub("X","",data_A_long$DueDate)
data_A_long$DueDate <- ymd(data_A_long$DueDate)

col_B <- colnames(data_B[,11:ncol(data_B)])

data_B_long <- pivot_longer(data = data_B,
                            cols = col_B,
                            names_to = "DueDate",
                            values_to = "AmountB")

data_B_long$DueDate <- gsub("X","",data_B_long$DueDate)
data_B_long$DueDate <- ymd(data_B_long$DueDate)

A_on_B <- colnames(data_A_long[,c(1,3,4,11)])

check_data <- left_join(data_A_long, data_B_long, by = A_on_B)
check_data <- check_data[,c(1,3,4,10,11,12,20)]

write.csv(check_data, "A vs B.v1.csv", row.names = FALSE)
