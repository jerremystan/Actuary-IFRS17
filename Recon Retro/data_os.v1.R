library(dplyr)
library(readxl)
library(tidyr)
library(lubridate)

data_os <- read.csv("Sprint Retro Claim OS.csv")
colnames(data_os) <- c("PolicyNo", "PolicyTransNo", "ClaimNo", "CurrencyCode", "RPTMonth", "RPTDate", "LossDate", "ClaimOSAsAt")

ifrs_group <- read_excel("Sprint.Result.20240422.202301.20221231.01.xlsx")
ifrs_group <- ifrs_group[,1:3]
colnames(ifrs_group) <- c("PolicyNo", "PolicyTransNo", "IFRSGroup")

ifrs_os <- left_join(ifrs_group, data_os, by = c("PolicyNo", "PolicyTransNo"))
ifrs_os <- ifrs_os[,c(1,2,4,3,5,6,7,8,9)]

ifrs_os_xNA <- na.omit(ifrs_os)



write.csv(ifrs_os_xNA, "Retro_Claim_OS.v1.csv", row.names = FALSE)

#FULL FROM ORACLE BI OS

data_os_2023 <- read.csv("Retro Claim OS.csv")

data_os_2023 <- data_os_2023[data_os_2023$Claims.Outstanding.AsAt...Org != 0,c(1,2,4,5,6,7,8,9)]

data_os_2023$Date.of.Loss <- substr(data_os_2023$Date.of.Loss, 1, 10)

colnames(data_os_2023) <- gsub("\\.","",colnames(data_os_2023))

write.csv(data_os_2023, "Retro_ClaimOS_IssueYear_2023.v2.csv", row.names = FALSE)

data_paid_2023 <- read.csv("Retro Claim Paid.csv")

data_paid_2023 <- data_paid_2023[data_paid_2023$Claims.Paid.After.Co.Out...Org != 0, c(1,2,4,5,6,7,8)]

nama_paid <- colnames(data_paid_2023)

nama_paid <- gsub("\\.","",nama_paid)

colnames(data_paid_2023) <- nama_paid

data_paid_2023$ApproveDate <- substr(data_paid_2023$ApproveDate, 1, 10)
data_paid_2023$DateofLoss <- substr(data_paid_2023$DateofLoss, 1, 10)

write.csv(data_paid_2023,"Retro_ClaimPaid_IssueYear_2023.v2.csv", row.names = FALSE)

data_ep_2023 <- read.csv("Retro Earned Prem.csv")

data_ep_2023 <- data_ep_2023[data_ep_2023$Earned.Premium.Aft.Co.Out...Org != 0, c(1,2,4,5,6)]

colnames(data_ep_2023) <- gsub("\\.","",colnames(data_ep_2023))

write.csv(data_ep_2023,"Retro_EP_IssueYear_2023.v2.csv", row.names = FALSE)
