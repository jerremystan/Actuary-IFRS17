library(dplyr)
library(purrr)
library(openxlsx)
library(lubridate)
library(stringr)
library(readxl)

source("generate mapping.v4.R")


list_tahun <- list(
              "2024",
              "2023",
              "2022",
              "2021",
              "2020",
              "2019",
              "2018",
              "2017",
              "2016",
              "2015")

all_year <- data.frame()

for (i in list_tahun){


nama_input <- paste0("mapping",i,".v4.csv")

data <- read.csv(nama_input,header=TRUE, na.strings=c("","NA"))

new_df <- as.data.frame(do.call(expand.grid, data))

#new_df <- expand.grid(
#  Valuation.Date = data$Valuation.Date,
#  Profit.Center = data$Profit.Center,
#  Reserving.Class = data$Reserving.Class,
#  Partner = data$Partner,
#  Expense.Type = data$Expense.Type
#)

lookup <- read.csv("Expense amount 2024 assumed.v1.csv",header=TRUE)

#print(new_df)'

cleaned_data <- na.omit(new_df)

#cleaned_data$Profit.Center <- gsub("NANIK", as.character("NA"),cleaned_data$Profit.Center)

#nrow(cleaned_data)'

'filter_Acq <- subset(cleaned_data, Expense.Type == "Acq")
merge_Acq <- left_join(filter_Acq, lookup, by = "Valuation.Date")
result_Acq <- subset(merge_Acq, select = -c(CHE, PME))
colnames(result_Acq)[colnames(result_Acq) == "Acq"] <- "Expense.Amount"'

filter_PME <- subset(cleaned_data, Expense.Type == "PME")
merge_PME <- left_join(filter_PME, lookup, by = "Valuation.Date")
result_PME <- subset(merge_PME, select = -c(Acq))
colnames(result_PME)[colnames(result_PME) == "PME"] <- "PME Percentage"
colnames(result_PME)[colnames(result_PME) == "CHE"] <- "CHE Percentage"
result_PME$Expense.Type <- NULL

'filter_CHE <- subset(cleaned_data, Expense.Type == "CHE")
merge_CHE <- left_join(filter_CHE, lookup, by = "Valuation.Date")
result_CHE <- subset(merge_CHE, select = -c(PME, Acq))
colnames(result_CHE)[colnames(result_CHE) == "CHE"] <- "CHE Percentage"
result_CHE$Expense.Type <- NULL'

#untuk write terakhir! dibawah
#write.csv(result_PME, file = paste0("D:/Stanley/Tugas IFRS17/Expense Template/Expense/Output/",i,"/",i,"-PME-TemplateExpense.csv"), row.names = FALSE)

all_year <- rbind(all_year,result_PME)

}


new_date <- all_year
kolom_tanggal <- new_date$Valuation.Date
new_date$Valuation.Date <- as.character(as.Date(kolom_tanggal, format = "%m/%d/%Y"))

colnames(new_date) <- c("Valuation Date","Profit Center", "Reserving Class", "CHE Percentage", "PME Percentage")

prior_period <- data.frame(ValuationDate = as.character(seq(as.Date("1999-01-01"),as.Date("2015-01-01"),by = "12 month")-1))

filter_prior <- filter(new_date, new_date$`Valuation Date` == "2015-12-31")

df_prior <- data.frame()
for (i in prior_period$ValuationDate){
  new_df <- filter_prior
  new_df$`Valuation Date` <- i
  df_prior <- rbind(df_prior, new_df)
}

new_date <- rbind(new_date, df_prior)
new_date <- new_date[order(new_date$`Valuation Date`),]

baris <- as.numeric(nrow(new_date))
df_temp <- data.frame()

for(i in 1:baris){
  #i <- 1
  tanggal <- as.character(new_date[i,1])
  tahun <- as.numeric(substr(tanggal,1,4))
  issue <- data.frame(seq(1998,tahun))
  n <- as.numeric(length(issue))
  
  new_df <- data.frame(
    ValuationDate = replicate(n, new_date[i,1]),
    Profit = replicate(n, new_date[i,2]),
    Reserve = replicate(n, new_date[i,3]),
    IssueYear = issue,
    che = replicate(n, new_date[i,4]),
    pme = replicate(n, new_date[i,5])
    
  )
  colnames(new_df) <- c("ValuationDate","ProfitCenter", "ReservingClass", "IssueYear","CHE","PME")
  
  df_temp <- bind_rows(df_temp, new_df)
  
}

lock_df <- df_temp

# Tadi sampe sini
filter_year <- df_temp[substr(df_temp$`ValuationDate`,6,7) == "12" & as.numeric(substr(df_temp$`ValuationDate`,1,4)) == df_temp$`IssueYear` ,]
filter_year$`IssueYear` <- filter_year$`IssueYear`+1

df_perm <- df_temp

df_temp <- df_perm

df_temp <- rbind(df_temp, filter_year) #RUN ULANG

df_temp <- df_temp[order(df_temp$`ValuationDate`,df_temp$`ProfitCenter`, df_temp$`ReservingClass`, df_temp$`IssueYear`),]

df_temp <- na.omit(df_temp)

df_temp$ACE <- 1

df_temp <- df_temp[,c(1,2,3,4,7,6,5)]

df_temp_ST_0 <- df_temp
df_temp_ST_0$IsShortTerm <- 0

df_temp_ST_1 <- df_temp
df_temp_ST_1$IsShortTerm <- 1

df_temp <- rbind(df_temp_ST_0, df_temp_ST_1)

df_temp <- df_temp[,c(1,2,3,4,8,5,6,7)]

df_temp <- df_temp[order(df_temp$`ValuationDate`, df_temp$`ProfitCenter`, df_temp$`ReservingClass`, df_temp$`IssueYear`, df_temp$IsShortTerm),]

write.csv(df_temp, "All-Expense-Issue-Year.v3.csv", row.names = FALSE)

#replace NotApp with NA

#df_temp$`Profit Center Code` <- gsub(df_temp$`Profit Center Code`, "NotApp", "NA")

df_temp_smt <- df_temp 

df_temp_smt$`ProfitCenter` <- gsub("NotApp", "NA", df_temp_smt$`ProfitCenter`)

df_temp <- df_temp_smt

df_temp <- df_temp[order(df_temp$ValuationDate, df_temp$ProfitCenter, df_temp$ReservingClass, df_temp$IssueYear, df_temp$IsShortTerm),]


#CREATE WORKBOOK

colnames(df_actual_2023_final) <- c("Valuation Date", "Profit Center Code", "Reserving Class Code", "Issue Year", "Is Short Term", "ACE", "PME", "CHE")

filter_df_act_ve <- df_actual_2023_final[df_actual_2023_final$`Reserving Class Code` == "VO", ]
filter_df_act_ve$`Reserving Class Code` <- "VE"

df_actual_2023_final <- rbind(df_actual_2023_final, filter_df_act_ve)

df_actual_2023_final <- df_actual_2023_final[order(df_actual_2023_final$`Valuation Date`, df_actual_2023_final$`Profit Center Code`, df_actual_2023_final$`Reserving Class Code`, df_actual_2023_final$`Issue Year`, df_actual_2023_final$`Is Short Term`),]

colnames(df_temp) <- c("Valuation Date", "Profit Center Code", "Reserving Class Code", "Issue Year", "Is Short Term", "ACE", "PME", "CHE")



wb <- createWorkbook()

stack_function <- function(tahun, sheet){
  
final_result <- filter(df_temp, substr(df_temp$`Valuation Date`,1,4) == tahun)

addWorksheet(wb, tahun)
writeData(wb, sheet, x = final_result, startRow = 1, startCol = 1 )

}



#TAMBAHKANN TAHUN DISINI!!!!

tahun_1998 <- stack_function("1998",1)
tahun_1999 <- stack_function("1999",2)
tahun_2000 <- stack_function("2000",3)
tahun_2001 <- stack_function("2001",4)
tahun_2002 <- stack_function("2002",5)
tahun_2003 <- stack_function("2003",6)
tahun_2004 <- stack_function("2004",7)
tahun_2005 <- stack_function("2005",8)
tahun_2006 <- stack_function("2006",9)
tahun_2007 <- stack_function("2007",10)
tahun_2008 <- stack_function("2008",11)
tahun_2009 <- stack_function("2009",12)
tahun_2010 <- stack_function("2010",13)
tahun_2011 <- stack_function("2011",14)
tahun_2012 <- stack_function("2012",15)
tahun_2013 <- stack_function("2013",16)
tahun_2014 <- stack_function("2014",17)
tahun_2015 <- stack_function("2015",18)
tahun_2016 <- stack_function("2016",19)
tahun_2017 <- stack_function("2017",20)
tahun_2018 <- stack_function("2018",21)
tahun_2019 <- stack_function("2019",22)
tahun_2020 <- stack_function("2020",23)
tahun_2021 <- stack_function("2021",24)
tahun_2022 <- stack_function("2022",25)
tahun_2023 <- stack_function("2023",26)
tahun_2024 <- stack_function("2024",27)

saveWorkbook(wb, "ExpectedExpense.v1.xlsx")

#CATATAN PENTING!!!!
#Replace 'NotApp' dengan 'NA' di excel
#GANTI FILE PROFIT CENTER SETIAP UPDATE ACCOUNTING
#UPDATE MAPPING (generate mapping.R)
#UPDATE AMOUNT


