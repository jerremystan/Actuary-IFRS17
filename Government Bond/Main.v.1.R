library(dplyr)
library(readxl)
library(openxlsx)

source("Fwd rate 0.5 - FOREIGN.R")
source("Fwd rate 0.5.R")

idr_data <- read.csv("InterestRate_IDR.csv")
frg_data <- read.csv("InterestRate_FOREIGN.csv")

final_result <- rbind(idr_data, frg_data)
final_result <- final_result[,-6:-7]
colnames(final_result) <- (c("Valuation Date", "Currency Code", "Tenor", "Spot", "Forward", "Spot + Illiquidity Premium", "Forward + Illiquidity Premium"))

write.xlsx(final_result, "InterestRate.v12.xlsx", sheetName = "Rate", rowNames = FALSE)
