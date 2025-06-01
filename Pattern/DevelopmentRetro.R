library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)
library(stringr)
library(data.table)

#list_data <- as.list(list.files(path = "input_IBNR"))

list_data <- "IBNR 2022-12-31.xlsx"

all_data <- data.frame()

for(i in list_data){
  
  #i <- "IBNR 2015-12-31.xlsx"
  
  tanggal <- substr(i,6,15)
  
  nama_A <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Auto_Grs",
                   "Auto_Gross")
  
  data_A <- read_excel(i, sheet = nama_A)

  nama_ME <- ifelse(tanggal == "2015-12-31" |
                    tanggal == "2016-12-31" |
                    tanggal == "2017-12-31" |
                    tanggal == "2018-12-31" |
                    tanggal == "2019-03-31" |
                    tanggal == "2019-06-30" |
                    tanggal == "2019-09-30" |
                    tanggal == "2019-12-31",
                    "Marine-All_Grs",
                    "Marine-ECommerce_Grs")
  
  data_ME <- read_excel(i, sheet = nama_ME)
  
  nama_MO <- ifelse(tanggal == "2015-12-31" |
                    tanggal == "2016-12-31" |
                    tanggal == "2017-12-31" |
                    tanggal == "2018-12-31" |
                    tanggal == "2019-03-31" |
                    tanggal == "2019-06-30" |
                    tanggal == "2019-09-30" |
                    tanggal == "2019-12-31",
                    "Marine-All_Grs",
                    "Marine-Other_Grs")
  
  data_MO <- read_excel(i, sheet = nama_MO)
  
  data_F <- read_excel(i, sheet = "Fire-All_Grs")
  
  data_E <- read_excel(i, sheet = "Eng-All_Grs")
  
  nama_L <- ifelse(tanggal == "2015-12-31" |
                   tanggal == "2016-12-31" |
                   tanggal == "2017-12-31",
                   "Various-All_Grs",
                   "Liability-All_Grs")
  
  data_L <- read_excel(i, sheet = nama_L)
  
  nama_P <- ifelse(tanggal == "2015-12-31" |
                   tanggal == "2016-12-31" |
                   tanggal == "2017-12-31",
                   "Various-All_Grs", 
                   ifelse(tanggal == "2018-12-31" |
                            tanggal == "2019-03-31" |
                            tanggal == "2019-06-30" |
                            tanggal == "2019-09-30" |
                            tanggal == "2019-12-31",
                   "PA-All_Grs", "PA-Retail_Grs"))
  
  data_P <- read_excel(i, sheet = nama_P)
  
  nama_T <- ifelse(tanggal == "2015-12-31" |
                   tanggal == "2016-12-31" |
                   tanggal == "2017-12-31" |
                   tanggal == "2018-12-31" |
                   tanggal == "2019-03-31" |
                   tanggal == "2019-06-30" |
                   tanggal == "2019-09-30" |
                   tanggal == "2019-12-31",
                   "Various-All_Grs",
                   "Travel-Retail_Grs")
  
  data_T <- read_excel(i, sheet = nama_T)
  
  nama_VE <- ifelse(  tanggal == "2015-12-31" |
                      tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30" |
                      tanggal == "2019-12-31" |
                      tanggal == "2020-03-31" |
                      tanggal == "2020-06-30" |
                      tanggal == "2020-09-30" |
                      tanggal == "2020-12-31" |
                      tanggal == "2021-03-31" |
                      tanggal == "2021-06-30" |
                      tanggal == "2021-09-30" |
                      tanggal == "2021-12-31" |
                      tanggal == "2022-03-31" |
                      tanggal == "2022-06-30" |
                      tanggal == "2022-09-30" |
                      tanggal == "2022-12-31" |
                      tanggal == "2023-03-31" |
                      tanggal == "2023-06-30",
                    "Various-All_Grs",
                    "Various-Ecommerce_Grs")
  
  data_VE <- read_excel(i, sheet = nama_VE)
  
  data_VO <- read_excel(i, sheet = "Various-All_Grs")
  
  nama_C <- ifelse(  tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31" |
                     tanggal == "2018-12-31" |
                     tanggal == "2019-03-31" |
                     tanggal == "2019-06-30" |
                     tanggal == "2019-09-30" |
                     tanggal == "2019-12-31",
                     "Various-All_Grs",
                     "TradeCredit-All_Grs")
  
  data_C <- read_excel(i, sheet = nama_C)
  
  nama_S <- ifelse(  tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31" |
                     tanggal == "2018-12-31" |
                     tanggal == "2019-03-31" |
                     tanggal == "2019-06-30" |
                     tanggal == "2019-09-30" |
                     tanggal == "2019-12-31" |
                     tanggal == "2020-03-31" |
                     tanggal == "2020-06-30" |
                     tanggal == "2020-09-30" |
                     tanggal == "2020-12-31",
                     "Various-All_Grs",
                     "Health-All_Grs")
  
  data_S <- read_excel(i, sheet = nama_S)
  
  
  
  n <- 24
  
  #val_date <- data.frame(ValuationDate = replicate(n,tanggal))
  #dev_quarter <- data.frame(DevelopmentQuarter = seq(1,n))
  
  DevelopmentFunction <- function(df,RC){
  
  index <- data.frame(index = which(df[,1] == "% Developed"))
    #df <- data_A
    #RC <- "A"
    
    df_temp <- data.frame()
    
    for(k in 1:n){
      #k <- 1
      
      ifelse( k != 1, paid <- 100*as.numeric(df[index[1,1],k+2])-100*as.numeric(df[index[1,1],k+1]),
                      paid <- 100*as.numeric(df[index[1,1],k+2]))
      
      ifelse( k != 1, incur <- 100*as.numeric(df[index[2,1],k+2])-100*as.numeric(df[index[2,1],k+1]),
                      incur <- 100*as.numeric(df[index[2,1],k+2]))
      
      ifelse(tanggal == "2015-12-31", method <- df[80-k,15],
             ifelse(  tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30", 
                      method <- df[80-k,23],
                      method <- df[81-k,23]))
        
        paid_nest <- rep(paid/3, 3)
        incur_nest <- rep(incur/3, 3)
        valuation_nest <- rep(tanggal,3)
        reserving_nest <- rep(RC,3)
        dev_nest <- seq((k-1)*3+1,k*3)
        method_nest <- rep(unlist(method[1,1]),3)
        
        new_df_nest<-data.frame(ValuationDate = valuation_nest,
                                ReservingClass = reserving_nest,
                                DevelopmentQuarter = dev_nest,
                                paid_nest,
                                incur_nest,
                                method_nest
                                )
        
        colnames(new_df_nest) <- c("ValuationDate", "ReservingClass", "DevelopmentQuarter", "Increment Dev Paid", "Increment Dev Incurred", "Method Selection")
        
        df_temp <- rbind(df_temp, new_df_nest)
      
      
    }
    
    return(df_temp)
    
  }
  
  result_A <- DevelopmentFunction(data_A,"A")
  result_ME <- DevelopmentFunction(data_ME,"ME")
  result_MO <- DevelopmentFunction(data_MO,"M")
  result_F <- DevelopmentFunction(data_F,"F")
  result_E <- DevelopmentFunction(data_E,"E")
  result_L <- DevelopmentFunction(data_L,"L")
  result_P <- DevelopmentFunction(data_P,"P")
  result_T <- DevelopmentFunction(data_T,"T")
  result_VE <- DevelopmentFunction(data_VE,"VE")
  result_VO <- DevelopmentFunction(data_VO,"VO")
  result_C <- DevelopmentFunction(data_C,"C")
  result_S <- DevelopmentFunction(data_S,"S")
  
  list_of_result <- list(
    result_A,
    result_ME,
    result_MO,
    result_F,
    result_E,
    result_L,
    result_P,
    result_T,
    result_VE,
    result_VO,
    result_C,
    result_S
  )
  
  
  
  stack_data <- rbindlist(list_of_result, use.names=FALSE)
  
  #all_data <- rbind(all_data, stack_data)
  nama_output <- paste0("Retro_Development_Assumptions_",tanggal,".csv")
  write.csv(stack_data, nama_output, row.names = FALSE)
  
}



'
team_A <- data.frame(ValuationDate = as.character(seq(as.Date("1998-04-01"), as.Date("2015-10-30"), by = "3 months")-1))
team_B <- data.frame(ValuationDate = c("2016-03-31","2016-06-30","2016-09-30"))
team_C <- data.frame(ValuationDate = c("2017-03-31","2017-06-30","2017-09-30"))
team_D <- data.frame(ValuationDate = c("2018-03-31","2018-06-30","2018-09-30"))

filter_A <- filter(all_data, all_data$`ValuationDate` == "2015-12-31")
filter_B <- filter(all_data, all_data$`ValuationDate` == "2015-12-31")
filter_C <- filter(all_data, all_data$`ValuationDate` == "2016-12-31")
filter_D <- filter(all_data, all_data$`ValuationDate` == "2017-12-31")

df_A <- data.frame()
for (i in team_A$ValuationDate){
  new_df <- filter_A
  new_df$`ValuationDate` <- i
  df_A <- rbind(df_A, new_df)
}

df_B <- data.frame()
for (i in team_B$ValuationDate){
  new_df <- filter_B
  new_df$`ValuationDate` <- i
  df_B <- rbind(df_B, new_df)
}

df_C <- data.frame()
for (i in team_C$ValuationDate){
  new_df <- filter_C
  new_df$`ValuationDate` <- i
  df_C <- rbind(df_C, new_df)
}

df_D <- data.frame()
for (i in team_D$ValuationDate){
  new_df <- filter_D
  new_df$`ValuationDate` <- i
  df_D <- rbind(df_D, new_df)
}

all_data <- rbind(all_data, df_A, df_B, df_C, df_D)
all_data <- all_data[order(all_data$`ValuationDate`),]

all_data$Future <- as.numeric(all_data$Future)*100

colnames(all_data) <- c("Valuation Date", "Reserving Class", "Accident Period", "Selected Paid DF CL", "Selected Incurred DF CL", "Method Selection", "Future Paid Development")

write.xlsx(all_data, "SelectedPaidIncurred.xlsx", sheetName = "LDF", rowNames = FALSE)

'
