library(readxl)
library(dplyr)


data_fo <- read_excel("D:/Stanley/Tugas IFRS17/NPR/NPR Working Paper.2022.v2.xlsx", sheet = "RI FO 2022")
data_fo <- data_fo[,c(1,6)]

map_client <- read.csv("D:/Stanley/Tugas IFRS17/NPR/Client ID Reinsurer - Edited Mapped for R.csv")
colnames(map_client) <- c("FO Company","Reinsure No",'BI Name')

data_fo <- left_join(data_fo, map_client, by = "FO Company")

data_fo <- unique(data_fo)

data_fo <- data_fo[,c(1,3,2)]
colnames(data_fo) <- c("Reinsurer Name", "Reinsurer No","NPR")

data_tr <- read_excel("D:/Stanley/Tugas IFRS17/NPR/NPR Working Paper.2022.v2.xlsx", sheet = "Treaty Map 2022")
data_tr <- data_tr[,c(3,1,4)]
colnames(data_tr) <- c("Reinsurer Name", "Reinsurer No","NPR")

data_npr <- rbind(data_tr, data_fo)

data_npr[is.na(data_npr)] <- ""

write.csv(data_npr,"D:/Stanley/Tugas IFRS17/NPR/NPR Reinsurer 2022.v1.csv", row.names = FALSE)
