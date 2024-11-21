library(dplyr)
library(purrr)
library(openxlsx)
library(lubridate)
library(stringr)
library(readxl)

list_year <- list.files("Amount Expense")

all_year_actual <- data.frame()
all_tahun <- c()

for (i in list_year){
  
  tahun <- substr(i,12,15)
  all_tahun <- c(all_tahun,tahun)
  
  val_date <- paste0(tahun,"-12-31")
  
  actual_data <- read_excel(paste0("Amount Expense/",i))
  
  actual_data <- replace(actual_data, is.na(actual_data),0)

  actual_data$ACE <- actual_data$ACE*100
  actual_data$PME <- actual_data$PME*100
  actual_data$CHE <- actual_data$CHE*100
  
  map_RC <- read_excel("Mapping/mappingRC.v1.xlsx")
  
  map_PC <- read_excel("Mapping/mappingPC.v1.xlsx")
  
  colnames(map_RC)[1] <- "LoB"
  
  join_actual_data <- left_join(actual_data, map_RC, by = "LoB")
  join_actual_data$IsShortTerm <- ifelse(join_actual_data$`Coverage Term` == "Short Term", 1, 0)
  
  base_actual_data <- join_actual_data[,c(2,7,8,4,5,6)]
  
  colnames(base_actual_data) <- c("ProfitCenter", "ReservingClassCode", "IsShortTerm","ACE","PME","CHE")
  
  base_actual_data_x <- left_join(base_actual_data, map_PC, by = "ProfitCenter")
  
  base_actual_data_x <- base_actual_data_x[,c(7,2,3,4,5,6)]
  
  df_actual_data <- data.frame()
  
  
  
  for (k in 1:nrow(base_actual_data_x)){
    # k <- 1
      # i <- 1
      issue <- seq(1998,as.numeric(tahun))
      n <- length(issue)
      
      new_df <- data.frame(ValuationDate = replicate(n,val_date),
                           ProfitCenter = unlist(replicate(n,base_actual_data_x[k,1])),
                           ReserveClass = unlist(replicate(n,base_actual_data_x[k,2])),
                           IssueYear = issue,
                           IsShortTerm = unlist(replicate(n,base_actual_data_x[k,3])),
                           ACE = unlist(replicate(n,base_actual_data_x[k,4])),
                           PME = unlist(replicate(n,base_actual_data_x[k,5])),
                           CHE = unlist(replicate(n,base_actual_data_x[k,6]))
      )
      
      colnames(new_df) <- c("ValuationDate", "ProfitCenter", "ReserveClass", "IssueYear", "IsShortTerm", "ACE", "PME", "CHE")
      
      df_actual_data <- rbind(df_actual_data, new_df)
      
    
  }
  
  df_actual_data_x <- df_actual_data
  
  #df_actual_data_x$ValuationDate <- as.Date(df_actual_data_x$ValuationDate, "%m/%d/%Y")
  
  df_actual_data12 <- df_actual_data_x[substr(df_actual_data_x$ValuationDate,6,7) == "12" & df_actual_data_x$IssueYear == as.numeric(tahun),]
  df_actual_data12$IssueYear <- as.numeric(tahun)+1
  
  df_actual_data_final <- rbind(df_actual_data_x, df_actual_data12)
  df_actual_data_final <- df_actual_data_final[order(df_actual_data_final$ValuationDate, df_actual_data_final$ProfitCenter, df_actual_data_final$ReserveClass, df_actual_data_final$IssueYear, df_actual_data_final$IsShortTerm),]
  
  all_year_actual <- rbind(all_year_actual, df_actual_data_final)
  
}


#CREATE WORKBOOK

colnames(all_year_actual) <- c("Valuation Date", "Profit Center Code", "Reserving Class Code", "Issue Year", "Is Short Term", "ACE", "PME", "CHE")

filter_df_act_ve <- all_year_actual[all_year_actual$`Reserving Class Code` == "VO", ]
filter_df_act_ve$`Reserving Class Code` <- "VE"

all_year_actual <- rbind(all_year_actual, filter_df_act_ve)

all_year_actual <- all_year_actual[order(all_year_actual$`Valuation Date`, all_year_actual$`Profit Center Code`, all_year_actual$`Reserving Class Code`, all_year_actual$`Issue Year`, all_year_actual$`Is Short Term`),]

colnames(all_year_actual) <- c("Valuation Date", "Profit Center Code", "Reserving Class Code", "Issue Year", "Is Short Term", "ACE", "PME", "CHE")



wb <- createWorkbook()

stack_function <- function(tahun, sheet){
  
  
  final_result <- filter(all_year_actual, substr(all_year_actual$`Valuation Date`,1,4) == tahun)
  
  addWorksheet(wb, tahun)
  writeData(wb, sheet, x = final_result, startRow = 1, startCol = 1 )
  
}



temp_n <- 0

for (i in all_tahun){
  
  temp_n <- temp_n+1
  assign(paste0("tahun_",i), stack_function(i,temp_n))
  
}


saveWorkbook(wb, "Output/ActualExpense.v2.xlsx")




