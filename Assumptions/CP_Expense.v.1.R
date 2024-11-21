library(dplyr)
library(purrr)
library(openxlsx)
library(lubridate)
library(stringr)
library(readxl)

contract_type <- read.csv("mapping/mapCS.csv")

year <- "2023"
month <- seq(1,12)
reserve_class <- read.csv("mapping/mapRC.csv")
mapRC <- reserve_class$ReservingClassCode
profit_center <- read.csv("mapping/mapPC.csv", na.strings = c("NoA"))
mapPCBI <- profit_center$PCBI
distribution <- read.csv("mapping/mapDC.csv", na.string = c("NoN"))
mapDC <- distribution$DistributionChannelCode
category_code <- c("-", "J", "L", "NA")
isshort <- c(0,1)
#s_paid <- c("OS", "Paid")

expand_df <- expand.grid(year = year, month = month, ReservingClassCode = mapRC, PCBI = mapPCBI, DC = mapDC, Category = category_code, isshort = isshort)

expand_join <- left_join(expand_df, reserve_class, by = "ReservingClassCode")
#expand_join <- left_join(expand_join, contract_type, by = "ReservingClassCode")
expand_join <- left_join(expand_join, profit_center, by = "PCBI")

cost_driver <- expand_join[,c(1,2,3,8,4,10,9,5,6,7)]
colnames(cost_driver) <- c("Incurred Year", "Incurred Month", "Reserving Class Code","Reserving Class", "Profit Center Code", "Profit Center", "Profit Center T Code", "Distribution Channel Code", "Category Code", "Is Short Term")

cost_driver <- cost_driver[order(cost_driver$`Incurred Year`, cost_driver$`Incurred Month`, cost_driver$`Reserving Class Code`, cost_driver$`Reserving Class`, cost_driver$`Profit Center Code`, cost_driver$`Profit Center`, cost_driver$`Profit Center T Code`, cost_driver$`Distribution Channel Code`, cost_driver$`Category Code`, cost_driver$`Is Short Term`),]

#Distribution Channel T Code adalah input ke SunSystem

cost_driver_gep <- cost_driver

colnames(cost_driver_gep)[1:2] <- c("Earned Year", "Earned Month")


cost_driver$`GIC wo Salvage/Subrogation` <- ""
cost_driver_gep$`GEP w/o Cancellation` <- ""
cost_driver_gep$`GEP with Cancellation` <- ""
cost_driver_gep$`NEP w/o Cancellation` <- ""
cost_driver_gep$`NEP with Cancellation` <- ""

cost_driver$`Incurred Year` <- as.integer(year)
cost_driver_gep$`Earned Year` <- as.integer(year)

#filter_cd <- cost_driver[cost_driver$`Reserving Class Code` == "VE",]

  
wb <- createWorkbook()

stack_function <- function(stat, term, sheet){
  
  
  ifelse(term == 0, t_name <- paste0(stat," Short Term"), t_name <- paste0(stat," Long Term"))
  
  ifelse(stat == "GIC", final_result <- cost_driver[cost_driver$`Is Short Term` == term,], final_result <- final_result <- cost_driver_gep[cost_driver_gep$`Is Short Term` == term,])
  
  addWorksheet(wb, t_name)
  writeData(wb, sheet, x = final_result, startRow = 1, startCol = 1)
  
}

gic_short_term <- stack_function("GIC",0,1)
gic_long_term <- stack_function("GIC",1,2)
gep_short_term <- stack_function("GEP",0,3)
gep_long_term <- stack_function("GEP",1,4)

excel_name <- paste0("CostDriver_",year,".v2.xlsx")

saveWorkbook(wb, excel_name)



'stack_function <- function(m, sheet){

  m_name <- month.name[m]
  
  final_result <- cost_driver[cost_driver$`Incurred Month` == m,]
  
  addWorksheet(wb, m_name)
  writeData(wb, sheet, x = final_result, startRow = 1, startCol = 1)
  
  
}

Januari <- stack_function(1,1)
Februari <- stack_function(2,2)
Maret <- stack_function(3,3)
April <- stack_function(4,4)
Mei <- stack_function(5,5)
Juni <- stack_function(6,6)
Juli <- stack_function(7,7)
Agustus <- stack_function(8,8)
September <- stack_function(9,9)
Oktober <- stack_function(10,10)
November <- stack_function(11,11)
Desember <- stack_function(12,12)'

#saveWorkbook(wb, "CostDriver.v2.xlsx")

#order_cost_driver <- do.call(order, cost_driver[1:12])



