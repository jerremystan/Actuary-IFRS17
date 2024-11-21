library(dplyr)

library(readxl)
data <- read_excel("Input/CONFIDENTIAL FILES.xlsx", sheet = "pivot")

source("Data EP SQL.v1.R")

nama_kolom_data <- data[1,]

index <- as.numeric(which(data[,1]=="Grand Total"))

data <- filter(data[2:index-1,])

data <- data[-1,]

colnames(data) <- nama_kolom_data


data_EP <- read.csv("Output/data_EP_final_SQL.csv")

map_RC <- read_excel("Mapping/ReservingClass.xlsx")

data_EP <- left_join(data_EP, map_RC, by = "ReservingClass")

data_EP <- data_EP[,-1]

data_EP <- data_EP[,c(3,1,2)]
colnames(data_EP) <- c("Main Class", "Profit Center", "EP After Co-Out")


data_EP <- left_join(data_EP, data, by = c("Main Class","Profit Center"))

data_EP$Acq <- 100*(as.numeric(data_EP$`Sum of Opex-Acq`)/as.numeric(data_EP$`EP After Co-Out`))
data_EP$PME <- 100*(as.numeric(data_EP$`Sum of Opex-PME`)/as.numeric(data_EP$`EP After Co-Out`))
data_EP$CHE <- 100*(as.numeric(data_EP$`Sum of Opex-CHE`)/as.numeric(data_EP$`EP After Co-Out`))

data_EP <- data_EP[,-3:-6]

data_EP[,3:5][is.na(data_EP[,3:5])] <- 0

for (i in 3:5){

  data_EP[,i] <- gsub("Inf",0,data_EP[,i])

}

write.csv(data_EP,"Output/Expense_Template_SQL.csv", row.names = FALSE)
