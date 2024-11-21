library(readxl)
library(lubridate)

year <- list("2023")


for (i in year){
  
  tahun <- i
  
  excel_file <- paste0("D:/Stanley/Tugas IFRS17/Exchange Rate/Exchange Rate/Input/",tahun,"-ExchangeRate.xlsx")
  
  month_map <- read.csv("D:/Stanley/Tugas IFRS17/Exchange Rate/Exchange Rate/Input/MonthMap.csv", header = TRUE)
  
  random_md_temp <- data.frame()
  
  for(j in 1:12){
    
    data <- read_excel(excel_file, sheet= month_map[j,2])
    
    random_md <- ymd(paste0(i,"-",j,"-15"))
    
    end_month <- ceiling_date(random_md, unit = "month")-1
    
    for (k in 1:31){
      
      new_row <- cbind(end_month, data[k+5,3], as.numeric(data[k+5,7]))
      random_md_temp <- rbind(random_md_temp, new_row)
      
    }#currency
    
    'new_row <- data.frame(RandomDate = random_md)
    rand_md_temp <- rbind(rand_md_temp, new_row)'

  }#bulan
  
  'rand_md_temp$RandomDate <- ceiling_date(rand_md_temp$RandomDate, unit = "month")-1
  
  colnames(rand_md_temp)[1] <- "Valuation Date"'
  
  nama_file <- paste0("D:/Stanley/Tugas IFRS17/Exchange Rate/Exchange Rate/Output/",i,"-ExchangeRate-Output.csv")
  
  colnames(random_md_temp) <- c("Valuation Date", "Currency", "Exchange Rate")
  write.csv(random_md_temp, file = nama_file, row.names = FALSE)
 
}#tahun 
