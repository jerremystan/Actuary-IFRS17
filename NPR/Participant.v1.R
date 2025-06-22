library(dplyr)
library(openxlsx)
library(readxl)
library(dplyr)

data_npr0 <- read.csv("Reins No and Participant No_1999_2003.csv", colClasses = c("Participant.No." = "character","Reinsurer.No." = "character"))
data_npr1 <- read.csv("Reins No and Participant No_2004_2008.csv", colClasses = c("Participant.No." = "character","Reinsurer.No." = "character"))
data_npr2 <- read.csv("Reins No and Participant No_2009_2013.csv", colClasses = c("Participant.No." = "character","Reinsurer.No." = "character"))
data_npr3 <- read.csv("Reins No and Participant No_2014_2018.csv", colClasses = c("Participant.No." = "character","Reinsurer.No." = "character"))
data_npr4 <- read.csv("Reins No and Participant No_2019_2024.csv", colClasses = c("Participant.No." = "character","Reinsurer.No." = "character"))

data_npr <- rbind(data_npr0,data_npr1,data_npr2,data_npr3,data_npr4)

list_npr <- read_excel("List Policy Renewal no Facultative.2.xlsx")

list_npr$Policy.Renewal.No. <- paste0(list_npr$`Policy No`,"-",list_npr$`Renewal No`)

list_npr <- list_npr[,-5:-6]

colnames(list_npr) <- c("Policy.No.","Renewal.No.","Risk.No.","Reinsurer.No.","Policy.Renewal.No.")

participant <- left_join(list_npr, data_npr, by = c("Policy.Renewal.No.","Risk.No.","Reinsurer.No."))

write.csv(participant, "FacultativeNPRResult.v2.csv", row.names = FALSE)


df_01 <- read.csv("FacultativeNPRResult.v1.csv",  colClasses = c("Participant.No." = "character","Reinsurer.No." = "character"))

df_02 <- read_excel("FacultativeNPRResult.v2.xlsx")
df_03 <- read_excel("FacultativeNPRResultEdited.v1.xlsx")

df_01$Participant.Mapping <- NA

df_01$Participant.Mapping <- df_02$`Participant Mapping`[,df_01$Policy.Renewal.No.==df_02$Policy.Renewal.No. & df_01$Risk.No. == df_02$Risk.No. & df_01$Reinsurer.No.==df_02$Reinsurer.No. & df_01$Participant.No. == df_02$Participant.No.]
