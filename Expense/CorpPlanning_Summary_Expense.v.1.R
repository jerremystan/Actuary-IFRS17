library(dplyr)
library(purrr)
library(openxlsx)
library(lubridate)
library(stringr)
library(readxl)

book <- c("T", "P")
month <- c("2023-01","2023-02","2023-03","2023-04","2023-05","2023-06","2023-07","2023-08","2023-09","2023-10","2023-11","2023-12")
group <- c("Corporate","Direct","Direct support")
categorycode <- c("J","L")
reserve_class <- read.csv("mapping/mapRC.csv")
mapRC <- unique(reserve_class$ReservingClass)
distribution <- read.csv("mapping/mapDC.csv", na.string = c("NoN"))
mapDC <- unique(distribution$DistributionChannelName)
profit_center <- read.csv("mapping/mapPC.csv", na.strings = c("NoA"))
mapPCCP <- unique(profit_center$PCCP)

expand_opx <- expand.grid(book = book, AccM = month, group = group, JL = categorycode, MainClass = mapRC, channel = mapDC, PC = mapPCCP)
colnames(expand_opx) <- c("BOOK", "Accounting Month", "Group", "JL", "Main Class", "Channel", "Profit Center")
expand_opx$`Opex-Acq` <- ""
expand_opx$`Opex-PME` <- ""
expand_opx$`Opex-CHE` <- ""
expand_opx$`Partner Name` <- ""
expense_allocated <- expand_opx[order(expand_opx$BOOK,expand_opx$`Accounting Month`,expand_opx$Group,expand_opx$JL,expand_opx$`Main Class`,expand_opx$Channel,expand_opx$`Profit Center`),]

wb <- createWorkbook()

stack_function <- function(m, sheet){
  
  m_name <- month.name[m]
  
  final_result <- expense_allocated[as.numeric(substr(expense_allocated$`Accounting Month`,6,7)) == m,]
  
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
Desember <- stack_function(12,12)

saveWorkbook(wb, "ExpenseAllocated.v1.xlsx")
