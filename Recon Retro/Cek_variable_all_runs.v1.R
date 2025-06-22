library(dplyr)
library(readxl)
library(tidyr)
library(lubridate)

tanggalA <- "2023-01-31"
tanggalB <- "2023-02-28"

data_A <- read.csv(paste0("D:/Stanley/Tugas IFRS17/Reconcile Retro/Cek input RiskIntegrity/variable_all_runs_",tanggalA,".csv"))
data_B <- read.csv(paste0("D:/Stanley/Tugas IFRS17/Reconcile Retro/Cek input RiskIntegrity/variable_all_runs_",tanggalB,".csv"))

col_A <- colnames(data_A[,10:ncol(data_A)])

data_A_long <- pivot_longer(data = data_A,
                            cols = col_A,
                            names_to = "CFDate",
                            values_to = "AmountA")

data_A_long$CFDate <- gsub("X","",data_A_long$CFDate)
data_A_long$CFDate <- ymd(data_A_long$CFDate)

col_B <- colnames(data_B[,10:ncol(data_B)])

data_B_long <- pivot_longer(data = data_B,
                            cols = col_B,
                            names_to = "CFDate",
                            values_to = "AmountB")

data_B_long$CFDate <- gsub("X","",data_B_long$CFDate)
data_B_long$CFDate <- ymd(data_B_long$CFDate)

filter_variable <- c("AQE_PAID_FUT", "P_PAID_FUT", "B_PAID_CUR", "E_PAID_CUR")

A_on_B <- colnames(data_A_long[,c(1,3,5,8,10)])

date_limit <- as.Date(tanggalB)

check_data <- left_join(data_A_long, data_B_long, by = A_on_B)
check_data <- check_data[check_data$variable_name %in% filter_variable & check_data$CFDate < date_limit, c(1,3,5,8,10,11,17)]

check_data$Diff <- ifelse(check_data$AmountA != check_data$AmountB, "ERROR", "NO ERROR")

write.csv(check_data, paste0(tanggalA," vs ",tanggalB,"_variable.v2.csv"), row.names = FALSE)
