library(lubridate)
library(rowr)

tahun <- list("2015",
              "2016",
              "2017",
              "2018",
              "2019",
              "2020",
              "2021",
              "2022",
              "2023",
              "2024")

for (i in tahun){
 
start_year <- i

dates <- read.csv(paste0("mapping",i,".csv"))
end_of_months_df <- data.frame(Valuation.Date = dates$Valuation.Date)
end_of_months_df <- data.frame(Valuation.Date = end_of_months_df[12,1])

profit_center <-  read.csv("ProfitCenter.v2.csv", header = TRUE)

reserving_class <- read.csv("ReservingClass.csv", header = TRUE)

expense_type <- data.frame("Expense Type"=c("Acq", "PME", "CHE"))

df_1 <- cbind.fill(end_of_months_df, profit_center, fill = "")
df_2 <- cbind.fill(df_1,reserving_class, fill = "")
df_3 <- cbind.fill(df_2, expense_type, fill = "")

write.csv(df_3, file = paste0("D:/Stanley/Tugas IFRS17/Expense Template/Expense/mapping",i,".v4.csv"), row.names = FALSE)

}
