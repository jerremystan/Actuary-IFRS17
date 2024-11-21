library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)
library(stringr)
library(data.table)

#list_data <- as.list(list.files(path = "input_IBNR"))

list_data <- c("IBNR 2023-12-31.xlsx")

#Gross Start

all_data <- data.frame()

for(i in list_data){
  print(i)
  #i <- "IBNR 2023-12-31.xlsx"
  
  tanggal <- substr(i,6,15)
  
  nama_A <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "AAA_Grs",
                   "AAA_Gross")
  
  data_A <- read_excel(i, sheet = nama_A)
  
  nama_ME <- ifelse(tanggal == "2015-12-31" |
                      tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30" |
                      tanggal == "2019-12-31",
                    "MMM-All_Grs",
                    "MMM-ECommerce_Grs")
  
  data_ME <- read_excel(i, sheet = nama_ME)
  
  nama_MO <- ifelse(tanggal == "2015-12-31" |
                      tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30" |
                      tanggal == "2019-12-31",
                    "MMM-All_Grs",
                    "MMM-Other_Grs")
  
  data_MO <- read_excel(i, sheet = nama_MO)
  
  data_F <- read_excel(i, sheet = "FFF-All_Grs")
  
  data_E <- read_excel(i, sheet = "EEE-All_Grs")
  
  nama_L <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "VVVV-All_Grs",
                   "LLL-All_Grs")
  
  data_L <- read_excel(i, sheet = nama_L)
  
  nama_P <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "VVVV-All_Grs", 
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
                   "VVVV-All_Grs",
                   "TTT-Retail_Grs")
  
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
                      "VVVV-All_Grs",
                      "VVVV-Ecommerce_Grs")
  
  data_VE <- read_excel(i, sheet = nama_VE)
  
  data_VO <- read_excel(i, sheet = "VVVV-All_Grs")
  
  nama_C <- ifelse(  tanggal == "2015-12-31" |
                       tanggal == "2016-12-31" |
                       tanggal == "2017-12-31" |
                       tanggal == "2018-12-31" |
                       tanggal == "2019-03-31" |
                       tanggal == "2019-06-30" |
                       tanggal == "2019-09-30" |
                       tanggal == "2019-12-31",
                     "VVVV-All_Grs",
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
                     "VVVV-All_Grs",
                     "HHH-All_Grs")
  
  data_S <- read_excel(i, sheet = nama_S)
  
  index <- data.frame(index = which(data_A[,1] == "Selected"))
  
  n <- 24
  
  #val_date <- data.frame(ValuationDate = replicate(n,tanggal))
  #dev_quarter <- data.frame(DevelopmentQuarter = seq(1,n))
  
  DevelopmentFunction <- function(df,RC){
    
    #df <- data_A
    #RC <- "A"
    
    df_temp <- data.frame()
    
    df_tot <- data.frame(X1 = as.numeric(unlist(df[326:303,2])))
    
    for(j in 1:24){
      
      for(k in 1:23){
        
        row_temp <- data.frame(
          ValuationDate = tanggal,
          ReservingClass = RC,
          AccidentPeriod = j,
          DevelopmentPeriod = k,
          FuturePaid = as.numeric(df[327-j,2+k])/df_tot[j,1]
        )
        
        df_temp <- rbind(df_temp, row_temp)
        
      }
      
      
      
    }
    
    colnames(df_temp) <- c("Valuation Date", "Reserving Class", "Accident Period", "Development Period", "Future Paid Development")
    
    data_final <- df_temp
    
    return(data_final)
    
  }
  
  result_A <- DevelopmentFunction(data_A,"A")
  result_ME <- DevelopmentFunction(data_ME,"ME")
  result_MO <- DevelopmentFunction(data_MO,"MO")
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
  
  all_data <- rbind(all_data, stack_data)
  
  
}

#Gross END


#Net Start


all_data_net <- data.frame()

for(i in list_data){
  
  #i <- "IBNR 2015-12-31.xlsx"
  
  tanggal <- substr(i,6,15)
  
  nama_A <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "AAA_Net",
                   "AAA_Net")
  
  data_A <- read_excel(i, sheet = nama_A)
  
  nama_ME <- ifelse(tanggal == "2015-12-31" |
                      tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30" |
                      tanggal == "2019-12-31",
                    "MMM-All_Net",
                    "MMM-ECommerce_Net")
  
  data_ME <- read_excel(i, sheet = nama_ME)
  
  nama_MO <- ifelse(tanggal == "2015-12-31" |
                      tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30" |
                      tanggal == "2019-12-31",
                    "MMM-All_Net",
                    "MMM-Other_Net")
  
  data_MO <- read_excel(i, sheet = nama_MO)
  
  data_F <- read_excel(i, sheet = "FFF-All_Net")
  
  data_E <- read_excel(i, sheet = "EEE-All_Net")
  
  nama_L <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "VVVV-All_Net",
                   "LLL-All_Net")
  
  data_L <- read_excel(i, sheet = nama_L)
  
  nama_P <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "VVVV-All_Net", 
                   ifelse(tanggal == "2018-12-31" |
                            tanggal == "2019-03-31" |
                            tanggal == "2019-06-30" |
                            tanggal == "2019-09-30" |
                            tanggal == "2019-12-31",
                          "PA-All_Net", "PA-Retail_Net"))
  
  data_P <- read_excel(i, sheet = nama_P)
  
  nama_T <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31" |
                     tanggal == "2018-12-31" |
                     tanggal == "2019-03-31" |
                     tanggal == "2019-06-30" |
                     tanggal == "2019-09-30" |
                     tanggal == "2019-12-31",
                   "VVVV-All_Net",
                   "TTT-Retail_Net")
  
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
                      "VVVV-All_Net",
                      "VVVV-Ecommerce_Net")
  
  data_VE <- read_excel(i, sheet = nama_VE)
  
  data_VO <- read_excel(i, sheet = "VVVV-All_Net")
  
  nama_C <- ifelse(  tanggal == "2015-12-31" |
                       tanggal == "2016-12-31" |
                       tanggal == "2017-12-31" |
                       tanggal == "2018-12-31" |
                       tanggal == "2019-03-31" |
                       tanggal == "2019-06-30" |
                       tanggal == "2019-09-30" |
                       tanggal == "2019-12-31",
                     "VVVV-All_Net",
                     "TradeCredit-All_Net")
  
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
                     "VVVV-All_Net",
                     "HHH-All_Net")
  
  data_S <- read_excel(i, sheet = nama_S)
  
  index <- data.frame(index = which(data_A[,1] == "Selected"))
  
  n <- 24
  
  #val_date <- data.frame(ValuationDate = replicate(n,tanggal))
  #dev_quarter <- data.frame(DevelopmentQuarter = seq(1,n))
  
  DevelopmentFunction <- function(df,RC){
    
    #df <- data_A
    #RC <- "A"
    
    df_temp <- data.frame()
    
    df_tot <- data.frame(X1 = as.numeric(unlist(df[326:303,2])))
    
    for(j in 1:24){
      
      for(k in 1:23){
        
        row_temp <- data.frame(
          ValuationDate = tanggal,
          ReservingClass = RC,
          AccidentPeriod = j,
          DevelopmentPeriod = k,
          FuturePaid = as.numeric(df[327-j,2+k])/df_tot[j,1]
        )
        df_temp <- rbind(df_temp, row_temp)
      }
      
      
      
    }
    
    colnames(df_temp) <- c("Valuation Date", "Reserving Class", "Accident Period", "Development Period", "Future Paid Development")
    
    data_final <- df_temp
    
    return(data_final)
    
  }
  
  result_A <- DevelopmentFunction(data_A,"A")
  result_ME <- DevelopmentFunction(data_ME,"ME")
  result_MO <- DevelopmentFunction(data_MO,"MO")
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
  
  all_data_net <- rbind(all_data_net, stack_data)
  
  
}

#Net END

colnames(all_data) <- c("Valuation Date", "Reserving Class", "Accident Period", "Development Period", "Gross Future Paid Development")
colnames(all_data_net) <- c("Valuation Date", "Reserving Class", "Accident Period", "Development Period", "Net Future Paid Development")

join_data_f <- left_join(all_data, all_data_net, by = c("Valuation Date", "Reserving Class", "Accident Period", "Development Period"))

all_data <- join_data_f

team_A <- data.frame(ValuationDate = as.character(seq(as.Date("1998-04-01"), as.Date("2015-10-30"), by = "3 months")-1))
team_B <- data.frame(ValuationDate = c("2016-03-31","2016-06-30","2016-09-30"))
team_C <- data.frame(ValuationDate = c("2017-03-31","2017-06-30","2017-09-30"))
team_D <- data.frame(ValuationDate = c("2018-03-31","2018-06-30","2018-09-30"))

filter_A <- filter(all_data, all_data$`Valuation Date` == "2015-12-31")
filter_B <- filter(all_data, all_data$`Valuation Date` == "2015-12-31")
filter_C <- filter(all_data, all_data$`Valuation Date` == "2016-12-31")
filter_D <- filter(all_data, all_data$`Valuation Date` == "2017-12-31")

df_A <- data.frame()
for (i in team_A$ValuationDate){
  new_df <- filter_A
  new_df$`Valuation Date` <- i
  df_A <- rbind(df_A, new_df)
}

df_B <- data.frame()
for (i in team_B$ValuationDate){
  new_df <- filter_B
  new_df$`Valuation Date` <- i
  df_B <- rbind(df_B, new_df)
}

df_C <- data.frame()
for (i in team_C$ValuationDate){
  new_df <- filter_C
  new_df$`Valuation Date` <- i
  df_C <- rbind(df_C, new_df)
}

df_D <- data.frame()
for (i in team_D$ValuationDate){
  new_df <- filter_D
  new_df$`Valuation Date` <- i
  df_D <- rbind(df_D, new_df)
}

all_data <- rbind(all_data, df_A, df_B, df_C, df_D)
all_data <- all_data[order(all_data$`Valuation Date`),]

all_data$`Gross Future Paid Development` <- as.numeric(all_data$`Gross Future Paid Development`)*100
all_data$`Net Future Paid Development` <- as.numeric(all_data$`Net Future Paid Development`)*100

colnames(all_data) <- c("Valuation Date", "Reserving Class Code", "Accident Period", "Development Period", "Gross Projected Percentage", "Net Projected Percentage")

all_data$`Gross Projected Percentage`[is.nan(all_data$`Gross Projected Percentage`)] <- 0
all_data$`Net Projected Percentage`[is.nan(all_data$`Net Projected Percentage`)] <- 0

write.xlsx(all_data, "CashflowProjectionPattern.v1.xlsx", sheetName = "Projection", rowNames = FALSE, overwrite = TRUE)

