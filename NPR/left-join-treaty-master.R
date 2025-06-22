library(dplyr)
library(readxl)
library(openxlsx)

treaty_code <- read_excel("RCH LoA.2022.v.1.xlsx", sheet = "NA NPR")

master <- read_excel("Stanley - Treaty Participant Master.xlsx")

colnames(master)[1] <- "Treaty Code"

join_df <- left_join(treaty_code, master, by = c("Treaty Code"))

colnames(join_df)[2:6] <- c("Reinsurer No"
                           ,"Reinsurer Name"
                           , "Participant No"
                           , "Participant Name"
                           , "Participant Percentage")

write.xlsx(join_df, "TreatyPercentage.xlsx", overwrite = TRUE)
