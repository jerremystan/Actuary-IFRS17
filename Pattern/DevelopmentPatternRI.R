library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)
library(stringr)
library(data.table)

list_data <- as.list(list.files(path = "input_IBNR"))

RI_mod <- c(
            "IBNR 2023-12-31.xlsx"
            , "IBNR 2024-12-31.xlsx"
            )

all_data <- data.frame()

#GROSS START

for(i in list_data){
  print(i)
  #i <- "IBNR 2015-12-31.xlsx"
  
  tanggal <- substr(i,6,15)
  
  nama_A <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Auto_Grs",
                   "Auto_Gross")
  
  data_A <- read_excel(paste0("input_IBNR/",i), sheet = nama_A)
  
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
  
  data_ME <- read_excel(paste0("input_IBNR/",i), sheet = nama_ME)
  
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
  
  data_MO <- read_excel(paste0("input_IBNR/",i), sheet = nama_MO)
  
  data_F <- read_excel(paste0("input_IBNR/",i), sheet = "Fire-All_Grs")
  
  data_E <- read_excel(paste0("input_IBNR/",i), sheet = "Eng-All_Grs")
  
  nama_L <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Various-All_Grs",
                   "Liability-All_Grs")
  
  data_L <- read_excel(paste0("input_IBNR/",i), sheet = nama_L)
  
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
  
  data_P <- read_excel(paste0("input_IBNR/",i), sheet = nama_P)
  
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
  
  data_T <- read_excel(paste0("input_IBNR/",i), sheet = nama_T)
  
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
  
  data_VE <- read_excel(paste0("input_IBNR/",i), sheet = nama_VE)
  
  data_VO <- read_excel(paste0("input_IBNR/",i), sheet = "Various-All_Grs")
  
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
  
  data_C <- read_excel(paste0("input_IBNR/",i), sheet = nama_C)
  
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
  
  data_S <- read_excel(paste0("input_IBNR/",i), sheet = nama_S)
  
  index <- data.frame(index = which(data_A[,1] == "Selected"))
  
  n <- 24
  
  #val_date <- data.frame(ValuationDate = replicate(n,tanggal))
  #dev_quarter <- data.frame(DevelopmentQuarter = seq(1,n))
  
  DevelopmentFunction <- function(df,RC){
    
    #df <- data_A
    #RC <- "A"
    
    df_temp <- data.frame()
    
    for(k in 1:n){
      
      paid <- df[index[1,1],k+2]
      incur <- df[index[2,1],k+2]
      new_row <- data.frame(ValuationDate = tanggal,
                            ReservingClass = RC,
                            DevelopmentQuarter = k,
                            paid,
                            incur)
      colnames(new_row) <- c("ValuationDate", "ReservingClass", "DevelopmentQuarter", "SelectedPaid", "SelectedIncurred")
      df_temp <- rbind(df_temp, new_row)
      
    }
    
    method <- data.frame()
    
    for(j in 1:n){
      
      ifelse(tanggal == "2015-12-31", new_row <- data.frame(MethodSelection = df[80-j,15]),
             ifelse(tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30", new_row <- data.frame(MethodSelection = df[80-j,23]),
                    new_row <- data.frame(MethodSelection = df[81-j,23])))
      
      method <- rbind(method, new_row)
      
    }
    colnames(method) <- "MethodSelection"
    
    future <- data.frame()
    for (m in 1:n){
      
      new_row <- data.frame(FuturePaid = df[327-m,57])
      future <- rbind(future, new_row)
      
    }
    colnames(future) <- "Future"
    
    data_final <- cbind(df_temp, method, future)
    
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
# GROSS END

#===============================================================================

#NET START

all_data_net <- data.frame()

for(i in list_data){
  
  #i <- "IBNR 2015-12-31.xlsx"
  
  tanggal <- substr(i,6,15)
  
  nama_A <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Auto_Net",
                   "Auto_Net")
  
  data_A <- read_excel(paste0("input_IBNR/",i), sheet = nama_A)
  
  nama_ME <- ifelse(tanggal == "2015-12-31" |
                      tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30" |
                      tanggal == "2019-12-31",
                    "Marine-All_Net",
                    "Marine-ECommerce_Net")
  
  data_ME <- read_excel(paste0("input_IBNR/",i), sheet = nama_ME)
  
  nama_MO <- ifelse(tanggal == "2015-12-31" |
                      tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30" |
                      tanggal == "2019-12-31",
                    "Marine-All_Net",
                    "Marine-Other_Net")
  
  data_MO <- read_excel(paste0("input_IBNR/",i), sheet = nama_MO)
  
  data_F <- read_excel(paste0("input_IBNR/",i), sheet = "Fire-All_Net")
  
  data_E <- read_excel(paste0("input_IBNR/",i), sheet = "Eng-All_Net")
  
  nama_L <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Various-All_Net",
                   "Liability-All_Net")
  
  data_L <- read_excel(paste0("input_IBNR/",i), sheet = nama_L)
  
  nama_P <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Various-All_Net", 
                   ifelse(tanggal == "2018-12-31" |
                            tanggal == "2019-03-31" |
                            tanggal == "2019-06-30" |
                            tanggal == "2019-09-30" |
                            tanggal == "2019-12-31",
                          "PA-All_Net", "PA-Retail_Net"))
  
  data_P <- read_excel(paste0("input_IBNR/",i), sheet = nama_P)
  
  nama_T <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31" |
                     tanggal == "2018-12-31" |
                     tanggal == "2019-03-31" |
                     tanggal == "2019-06-30" |
                     tanggal == "2019-09-30" |
                     tanggal == "2019-12-31",
                   "Various-All_Net",
                   "Travel-Retail_Net")
  
  data_T <- read_excel(paste0("input_IBNR/",i), sheet = nama_T)
  
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
                      "Various-All_Net",
                      "Various-Ecommerce_Net")
  
  data_VE <- read_excel(paste0("input_IBNR/",i), sheet = nama_VE)
  
  data_VO <- read_excel(paste0("input_IBNR/",i), sheet = "Various-All_Net")
  
  nama_C <- ifelse(  tanggal == "2015-12-31" |
                       tanggal == "2016-12-31" |
                       tanggal == "2017-12-31" |
                       tanggal == "2018-12-31" |
                       tanggal == "2019-03-31" |
                       tanggal == "2019-06-30" |
                       tanggal == "2019-09-30" |
                       tanggal == "2019-12-31",
                     "Various-All_Net",
                     "TradeCredit-All_Net")
  
  data_C <- read_excel(paste0("input_IBNR/",i), sheet = nama_C)
  
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
                     "Various-All_Net",
                     "Health-All_Net")
  
  data_S <- read_excel(paste0("input_IBNR/",i), sheet = nama_S)
  
  index <- data.frame(index = which(data_A[,1] == "Selected"))
  
  n <- 24
  
  #val_date <- data.frame(ValuationDate = replicate(n,tanggal))
  #dev_quarter <- data.frame(DevelopmentQuarter = seq(1,n))
  
  DevelopmentFunction <- function(df,RC){
    
    #df <- data_A
    #RC <- "A"
    
    df_temp <- data.frame()
    
    for(k in 1:n){
      
      paid <- df[index[1,1],k+2]
      incur <- df[index[2,1],k+2]
      new_row <- data.frame(ValuationDate = tanggal,
                            ReservingClass = RC,
                            DevelopmentQuarter = k,
                            paid,
                            incur)
      colnames(new_row) <- c("ValuationDate", "ReservingClass", "DevelopmentQuarter", "SelectedPaid", "SelectedIncurred")
      df_temp <- rbind(df_temp, new_row)
      
    }
    
    method <- data.frame()
    
    for(j in 1:n){
      
      ifelse(tanggal == "2015-12-31", new_row <- data.frame(MethodSelection = df[80-j,15]),
             ifelse(tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30", new_row <- data.frame(MethodSelection = df[80-j,23]),
                    new_row <- data.frame(MethodSelection = df[81-j,23])))
      
      method <- rbind(method, new_row)
      
    }
    colnames(method) <- "MethodSelection"
    
    future <- data.frame()
    for (m in 1:n){
      
      new_row <- data.frame(FuturePaid = df[327-m,57])
      future <- rbind(future, new_row)
      
    }
    colnames(future) <- "Future"
    
    data_final <- cbind(df_temp, method, future)
    
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

# NET END

#===============================================================================

#NXL START


all_data_nxl <- data.frame()

for(i in RI_mod){
  
  #i <- "IBNR 2015-12-31.xlsx"
  
  tanggal <- substr(i,6,15)
  
  nama_A <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Auto_NXL_RI",
                   "Auto_NXL_RI")
  
  data_A <- read_excel(paste0("input_IBNR/",i), sheet = nama_A)
  
  nama_ME <- ifelse(tanggal == "2015-12-31" |
                      tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30" |
                      tanggal == "2019-12-31",
                    "Marine-All_NXL_RI",
                    "Marine-ECommerce_NXL_RI")
  
  data_ME <- read_excel(paste0("input_IBNR/",i), sheet = nama_ME)
  
  nama_MO <- ifelse(tanggal == "2015-12-31" |
                      tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30" |
                      tanggal == "2019-12-31",
                    "Marine-All_NXL_RI",
                    "Marine-Other_NXL_RI")
  
  data_MO <- read_excel(paste0("input_IBNR/",i), sheet = nama_MO)
  
  data_F <- read_excel(paste0("input_IBNR/",i), sheet = "Fire-All_NXL_RI")
  
  data_E <- read_excel(paste0("input_IBNR/",i), sheet = "Eng-All_NXL_RI")
  
  nama_L <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Various-All_NXL_RI",
                   "Liability-All_NXL_RI")
  
  data_L <- read_excel(paste0("input_IBNR/",i), sheet = nama_L)
  
  nama_P <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Various-All_NXL_RI", 
                   ifelse(tanggal == "2018-12-31" |
                            tanggal == "2019-03-31" |
                            tanggal == "2019-06-30" |
                            tanggal == "2019-09-30" |
                            tanggal == "2019-12-31",
                          "PA-All_NXL_RI", "PA-Retail_NXL_RI"))
  
  data_P <- read_excel(paste0("input_IBNR/",i), sheet = nama_P)
  
  nama_T <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31" |
                     tanggal == "2018-12-31" |
                     tanggal == "2019-03-31" |
                     tanggal == "2019-06-30" |
                     tanggal == "2019-09-30" |
                     tanggal == "2019-12-31",
                   "Various-All_NXL_RI",
                   "Travel-Retail_NXL_RI")
  
  data_T <- read_excel(paste0("input_IBNR/",i), sheet = nama_T)
  
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
                      "Various-All_NXL_RI",
                      "Various-ECommerce_NXL_RI")
  
  data_VE <- read_excel(paste0("input_IBNR/",i), sheet = nama_VE)
  
  data_VO <- read_excel(paste0("input_IBNR/",i), sheet = "Various-All_NXL_RI")
  
  nama_C <- ifelse(  tanggal == "2015-12-31" |
                       tanggal == "2016-12-31" |
                       tanggal == "2017-12-31" |
                       tanggal == "2018-12-31" |
                       tanggal == "2019-03-31" |
                       tanggal == "2019-06-30" |
                       tanggal == "2019-09-30" |
                       tanggal == "2019-12-31",
                     "Various-All_NXL_RI",
                     "TradeCredit-All_NXL_RI")
  
  data_C <- read_excel(paste0("input_IBNR/",i), sheet = nama_C)
  
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
                     "Various-All_NXL_RI",
                     "Health-All_NXL_RI")
  
  data_S <- read_excel(paste0("input_IBNR/",i), sheet = nama_S)
  
  index <- data.frame(index = which(data_A[,1] == "Selected"))
  
  n <- 24
  
  #val_date <- data.frame(ValuationDate = replicate(n,tanggal))
  #dev_quarter <- data.frame(DevelopmentQuarter = seq(1,n))
  
  DevelopmentFunction <- function(df,RC){
    
    #df <- data_A
    #RC <- "A"
    
    df_temp <- data.frame()
    
    for(k in 1:n){
      
      paid <- df[index[1,1],k+2]
      incur <- df[index[2,1],k+2]
      new_row <- data.frame(ValuationDate = tanggal,
                            ReservingClass = RC,
                            DevelopmentQuarter = k,
                            paid,
                            incur)
      colnames(new_row) <- c("ValuationDate", "ReservingClass", "DevelopmentQuarter", "SelectedPaid", "SelectedIncurred")
      df_temp <- rbind(df_temp, new_row)
      
    }
    
    method <- data.frame()
    
    for(j in 1:n){
      
      ifelse(tanggal == "2015-12-31", new_row <- data.frame(MethodSelection = df[80-j,15]),
             ifelse(tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30", new_row <- data.frame(MethodSelection = df[80-j,23]),
                    new_row <- data.frame(MethodSelection = df[81-j,23])))
      
      method <- rbind(method, new_row)
      
    }
    colnames(method) <- "MethodSelection"
    
    future <- data.frame()
    for (m in 1:n){
      
      new_row <- data.frame(FuturePaid = df[327-m,57])
      future <- rbind(future, new_row)
      
    }
    colnames(future) <- "Future"
    
    data_final <- cbind(df_temp, method, future)
    
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
  
  all_data_nxl <- rbind(all_data_nxl, stack_data)
  
}

#NXL END

#===============================================================================

#XL START


all_data_xl <- data.frame()

for(i in RI_mod){
  
  #i <- "IBNR 2015-12-31.xlsx"
  
  tanggal <- substr(i,6,15)
  
  nama_A <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Auto_XL_RI",
                   "Auto_XL_RI")
  
  data_A <- read_excel(paste0("input_IBNR/",i), sheet = nama_A)
  
  nama_ME <- ifelse(tanggal == "2015-12-31" |
                      tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30" |
                      tanggal == "2019-12-31",
                    "Marine-All_XL_RI",
                    "Marine-ECommerce_XL_RI")
  
  data_ME <- read_excel(paste0("input_IBNR/",i), sheet = nama_ME)
  
  nama_MO <- ifelse(tanggal == "2015-12-31" |
                      tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30" |
                      tanggal == "2019-12-31",
                    "Marine-All_XL_RI",
                    "Marine-Other_XL_RI")
  
  data_MO <- read_excel(paste0("input_IBNR/",i), sheet = nama_MO)
  
  data_F <- read_excel(paste0("input_IBNR/",i), sheet = "Fire-All_XL_RI")
  
  data_E <- read_excel(paste0("input_IBNR/",i), sheet = "Eng-All_XL_RI")
  
  nama_L <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Various-All_XL_RI",
                   "Liability-All_XL_RI")
  
  data_L <- read_excel(paste0("input_IBNR/",i), sheet = nama_L)
  
  nama_P <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31",
                   "Various-All_XL_RI", 
                   ifelse(tanggal == "2018-12-31" |
                            tanggal == "2019-03-31" |
                            tanggal == "2019-06-30" |
                            tanggal == "2019-09-30" |
                            tanggal == "2019-12-31",
                          "PA-All_XL_RI", "PA-Retail_XL_RI"))
  
  data_P <- read_excel(paste0("input_IBNR/",i), sheet = nama_P)
  
  nama_T <- ifelse(tanggal == "2015-12-31" |
                     tanggal == "2016-12-31" |
                     tanggal == "2017-12-31" |
                     tanggal == "2018-12-31" |
                     tanggal == "2019-03-31" |
                     tanggal == "2019-06-30" |
                     tanggal == "2019-09-30" |
                     tanggal == "2019-12-31",
                   "Various-All_XL_RI",
                   "Travel-Retail_XL_RI")
  
  data_T <- read_excel(paste0("input_IBNR/",i), sheet = nama_T)
  
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
                      "Various-All_XL_RI",
                      "Various-ECommerce_XL_RI")
  
  data_VE <- read_excel(paste0("input_IBNR/",i), sheet = nama_VE)
  
  data_VO <- read_excel(paste0("input_IBNR/",i), sheet = "Various-All_XL_RI")
  
  nama_C <- ifelse(  tanggal == "2015-12-31" |
                       tanggal == "2016-12-31" |
                       tanggal == "2017-12-31" |
                       tanggal == "2018-12-31" |
                       tanggal == "2019-03-31" |
                       tanggal == "2019-06-30" |
                       tanggal == "2019-09-30" |
                       tanggal == "2019-12-31",
                     "Various-All_XL_RI",
                     "TradeCredit-All_XL_RI")
  
  data_C <- read_excel(paste0("input_IBNR/",i), sheet = nama_C)
  
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
                     "Various-All_XL_RI",
                     "Health-All_XL_RI")
  
  data_S <- read_excel(paste0("input_IBNR/",i), sheet = nama_S)
  
  index <- data.frame(index = which(data_A[,1] == "Selected"))
  
  n <- 24
  
  #val_date <- data.frame(ValuationDate = replicate(n,tanggal))
  #dev_quarter <- data.frame(DevelopmentQuarter = seq(1,n))
  
  DevelopmentFunction <- function(df,RC){
    
    #df <- data_A
    #RC <- "A"
    
    df_temp <- data.frame()
    
    for(k in 1:n){
      
      paid <- df[index[1,1],k+2]
      incur <- df[index[2,1],k+2]
      new_row <- data.frame(ValuationDate = tanggal,
                            ReservingClass = RC,
                            DevelopmentQuarter = k,
                            paid,
                            incur)
      colnames(new_row) <- c("ValuationDate", "ReservingClass", "DevelopmentQuarter", "SelectedPaid", "SelectedIncurred")
      df_temp <- rbind(df_temp, new_row)
      
    }
    
    method <- data.frame()
    
    for(j in 1:n){
      
      ifelse(tanggal == "2015-12-31", new_row <- data.frame(MethodSelection = df[80-j,15]),
             ifelse(tanggal == "2016-12-31" |
                      tanggal == "2017-12-31" |
                      tanggal == "2018-12-31" |
                      tanggal == "2019-03-31" |
                      tanggal == "2019-06-30" |
                      tanggal == "2019-09-30", new_row <- data.frame(MethodSelection = df[80-j,23]),
                    new_row <- data.frame(MethodSelection = df[81-j,23])))
      
      method <- rbind(method, new_row)
      
    }
    colnames(method) <- "MethodSelection"
    
    future <- data.frame()
    for (m in 1:n){
      
      new_row <- data.frame(FuturePaid = df[327-m,57])
      future <- rbind(future, new_row)
      
    }
    colnames(future) <- "Future"
    
    data_final <- cbind(df_temp, method, future)
    
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
  
  all_data_xl <- rbind(all_data_xl, stack_data)
  
}


#XL END

#===============================================================================
#===============================================================================

colnames(all_data) <- c("ValuationDate", "ReservingClass", "DevelopmentQuarter", "Grs_SelectedPaid", "Grs_SelectedIncurred", "Grs_MethodSelection", "Grs_Future")
colnames(all_data_net) <- c("ValuationDate", "ReservingClass", "DevelopmentQuarter", "Net_SelectedPaid", "Net_SelectedIncurred", "Net_MethodSelection", "Net_Future")
colnames(all_data_nxl) <- c("ValuationDate", "ReservingClass", "DevelopmentQuarter", "NXL_SelectedPaid", "NXL_SelectedIncurred", "NXL_MethodSelection", "NXL_Future")
colnames(all_data_xl) <- c("ValuationDate", "ReservingClass", "DevelopmentQuarter", "XL_SelectedPaid", "XL_SelectedIncurred", "XL_MethodSelection", "XL_Future")

join_data <- left_join(all_data, all_data_net, by = c("ValuationDate", "ReservingClass", "DevelopmentQuarter"))
join_data <- left_join(join_data, all_data_nxl, by = c("ValuationDate", "ReservingClass", "DevelopmentQuarter"))
join_data <- left_join(join_data, all_data_xl, by = c("ValuationDate", "ReservingClass", "DevelopmentQuarter"))

all_data <- join_data

# team_A <- data.frame(ValuationDate = as.character(seq(as.Date("1998-04-01"), as.Date("2015-10-30"), by = "3 months")-1))
# team_B <- data.frame(ValuationDate = c("2016-03-31","2016-06-30","2016-09-30"))
# team_C <- data.frame(ValuationDate = c("2017-03-31","2017-06-30","2017-09-30"))
# team_D <- data.frame(ValuationDate = c("2018-03-31","2018-06-30","2018-09-30"))
# 
# filter_A <- filter(all_data, all_data$`ValuationDate` == "2015-12-31")
# filter_B <- filter(all_data, all_data$`ValuationDate` == "2015-12-31")
# filter_C <- filter(all_data, all_data$`ValuationDate` == "2016-12-31")
# filter_D <- filter(all_data, all_data$`ValuationDate` == "2017-12-31")
# 
# df_A <- data.frame()
# for (i in team_A$ValuationDate){
#   new_df <- filter_A
#   new_df$`ValuationDate` <- i
#   df_A <- rbind(df_A, new_df)
# }
# 
# df_B <- data.frame()
# for (i in team_B$ValuationDate){
#   new_df <- filter_B
#   new_df$`ValuationDate` <- i
#   df_B <- rbind(df_B, new_df)
# }
# 
# df_C <- data.frame()
# for (i in team_C$ValuationDate){
#   new_df <- filter_C
#   new_df$`ValuationDate` <- i
#   df_C <- rbind(df_C, new_df)
# }
# 
# df_D <- data.frame()
# for (i in team_D$ValuationDate){
#   new_df <- filter_D
#   new_df$`ValuationDate` <- i
#   df_D <- rbind(df_D, new_df)
# }
# 
# all_data <- rbind(all_data, df_A, df_B, df_C, df_D)

all_data <- all_data[order(all_data$`ValuationDate`),]

all_data$Grs_Future <- as.numeric(all_data$Grs_Future)*100
all_data$Net_Future <- as.numeric(all_data$Net_Future)*100
all_data$NXL_Future <- as.numeric(all_data$NXL_Future)*100
all_data$XL_Future <- as.numeric(all_data$XL_Future)*100


colnames(all_data) <- c("Valuation Date", "Reserving Class", "Accident Period" #1,2,3
                        , "Gross Selected Paid DF CL", "Gross Selected Incurred DF CL", "Gross Method Selection", "Gross Future Paid Development" #4,5,6,7
                        , "Net Selected Paid DF CL", "Net Selected Incurred DF CL", "Net Method Selection", "Net Future Paid Development" #8,9,10,11
                        , "Non XL Reinsurance Selected Paid DF CL", "Non XL Reinsurance Selected Incurred DF CL", "Non XL Reinsurance Method Selection", "Non XL Reinsurance Future Paid Development" #12,13,14,15
                        , "XL Reinsurance Selected Paid DF CL", "XL Reinsurance Selected Incurred DF CL", "XL Reinsurance Method Selection", "XL Reinsurance Future Paid Development" #16,17,18,19
                        )



#Push formula to excel




all_data$`Gross Paid DF CL` <- "" #20
all_data$`Gross Incurred DF CL` <- "" #21
all_data$`Gross Cummulative Paid DF CL` <- "" #22
all_data$`Gross Cummulative Incurred DF CL` <- "" #23

all_data$`Net Paid DF CL` <- "" #24
all_data$`Net Incurred DF CL` <- "" #25
all_data$`Net Cummulative Paid DF CL` <- "" #26
all_data$`Net Cummulative Incurred DF CL` <- "" #27

all_data$`Non XL Reinsurance Paid DF CL` <- "" #28
all_data$`Non XL Reinsurance Incurred DF CL` <- "" #29
all_data$`Non XL Reinsurance Cummulative Paid DF CL` <- "" #30
all_data$`Non XL Reinsurance Cummulative Incurred DF CL` <- "" #31

all_data$`XL Reinsurance Paid DF CL` <- "" #32
all_data$`XL Reinsurance Incurred DF CL` <- "" #33
all_data$`XL Reinsurance Cummulative Paid DF CL` <- "" #34
all_data$`XL Reinsurance Cummulative Incurred DF CL` <- "" #35

all_data_strd <- all_data[,c(1,2,3
                             ,20,21,4,5,22,23,6,7 #Gross
                             ,24,25,8,9,26,27,10,11 #Net
                             ,28,29,12,13,30,31,14,15 #NXL
                             ,32,33,16,17,34,35,18,19 #XL
                             )]

write.xlsx(all_data_strd, "DevelopmentPaidIncurred.2024Q1.v1.xlsx", sheetName = "LDF", rowNames = FALSE)


wb <- loadWorkbook("DevelopmentPaidIncurred.2024Q1.v1.xlsx")

m <- nrow(all_data_strd)


#===============================================================================

#Gross Paid DF CL
form_A <- data.frame()
for (i in 1:m){
  #=100*(IF($C2=1,1/PRODUCT(F2:OFFSET(F2,24-$C2,0)),1/PRODUCT(F2:OFFSET(F2,24-$C2,0))-1/PRODUCT(F1:OFFSET(F1,24-$C1,0))))
  temp_form <- paste0("=100*(IF($C",1+i,"=1,1/PRODUCT(F",1+i,":OFFSET(F",1+i,",24-$C",1+i,",0)),1/PRODUCT(F",1+i,":OFFSET(F",1+i,",24-$C",1+i,",0))-1/PRODUCT(F",i,":OFFSET(F",i,",24-$C",i,",0))))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  4, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
  
}
writeFormula(wb, sheet = "LDF", startCol =  4, startRow = 2, x = form_A[,1])

#Gross Incurred DF CL
form_A <- data.frame()
for (i in 1:m){
  #=100*IF($C2=1,1/PRODUCT(G2:OFFSET(G2,24-$C2,0)),1/PRODUCT(G2:OFFSET(G2,24-$C2,0))-1/PRODUCT(G1:OFFSET(G1,24-$C1,0)))
  temp_form <- paste0("=100*IF($C",1+i,"=1,1/PRODUCT(G",1+i,":OFFSET(G",1+i,",24-$C",1+i,",0)),1/PRODUCT(G",1+i,":OFFSET(G",1+i,",24-$C",1+i,",0))-1/PRODUCT(G",i,":OFFSET(G",i,",24-$C",i,",0)))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  5, startRow = temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
  
}
writeFormula(wb, sheet = "LDF", startCol =  5, startRow = 2, x = form_A[,1])

#Gross Cummulative Paid DF CL
form_A <- data.frame()
for (i in 1:m){
  #i <- 1
  #=IF($C2=1,100/D2,100/SUMIFS(D:D,$A:$A,$A2,$B:$B,$B2,$C:$C,"<="&$C2))
  temp_form <- paste0("=IF($C",1+i,"=1,100/D",1+i,",100/SUMIFS(D:D,$A:$A,$A",1+i,",$B:$B,$B",1+i,",$C:$C,\"<=\"&$C",1+i,"))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  8, startRow = temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  8, startRow = 2, x = form_A[,1])

#Gross Cummulative Incurred DF CL
form_A <- data.frame()
for (i in 1:m){
  #=IF($C2=1,100/E2,100/SUMIFS(E:E,$A:$A,$A2,$B:$B,$B2,$C:$C,"<="&$C2))
  temp_form <- paste0("=IF($C",1+i,"=1,100/E",1+i,",100/SUMIFS(E:E,$A:$A,$A",1+i,",$B:$B,$B",1+i,",$C:$C,\"<=\"&$C",1+i,"))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  9, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  9, startRow = 2, x = form_A[,1])

#===============================================================================

#Net Paid DF CL
form_A <- data.frame()
for (i in 1:m){
  #=100*(IF($C2=1,1/PRODUCT(N2:OFFSET(N2,24-$C2,0)),1/PRODUCT(N2:OFFSET(N2,24-$C2,0))-1/PRODUCT(N1:OFFSET(N1,24-$C1,0))))
  temp_form <- paste0("=100*(IF($C",1+i,"=1,1/PRODUCT(N",1+i,":OFFSET(N",1+i,",24-$C",1+i,",0)),1/PRODUCT(N",1+i,":OFFSET(N",1+i,",24-$C",1+i,",0))-1/PRODUCT(N",i,":OFFSET(N",i,",24-$C",i,",0))))")
  #writeFormula(wb, sheet = "LDF", startCol =  12, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  12, startRow = 2, x = form_A[,1])

#Net Incurred DF CL
form_A <- data.frame()
for (i in 1:m){
  #=100*IF($C2=1,1/PRODUCT(O2:OFFSET(O2,24-$C2,0)),1/PRODUCT(O2:OFFSET(O2,24-$C2,0))-1/PRODUCT(O1:OFFSET(O1,24-$C1,0)))
  temp_form <- paste0("=100*IF($C",1+i,"=1,1/PRODUCT(O",1+i,":OFFSET(O",1+i,",24-$C",1+i,",0)),1/PRODUCT(O",1+i,":OFFSET(O",1+i,",24-$C",1+i,",0))-1/PRODUCT(O",i,":OFFSET(O",i,",24-$C",i,",0)))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  13, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  13, startRow = 2, x = form_A[,1])

#Net Cummulative Paid DF CL
form_A <- data.frame()
for (i in 1:m){
  #=IF($C2=1,100/L2,100/SUMIFS(L:L,$A:$A,$A2,$B:$B,$B2,$C:$C,"<="&$C2))
  temp_form <- paste0("=IF($C",1+i,"=1,100/L",1+i,",100/SUMIFS(L:L,$A:$A,$A",1+i,",$B:$B,$B",1+i,",$C:$C,\"<=\"&$C",1+i,"))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  16, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  16, startRow = 2, x = form_A[,1])

#Net Cummulative Incurred DF CL
form_A <- data.frame()
for (i in 1:m){
  #=IF($C2=1,100/M2,100/SUMIFS(M:M,$A:$A,$A2,$B:$B,$B2,$C:$C,"<="&$C2))
  temp_form <- paste0("=IF($C",1+i,"=1,100/M",1+i,",100/SUMIFS(M:M,$A:$A,$A",1+i,",$B:$B,$B",1+i,",$C:$C,\"<=\"&$C",1+i,"))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  17, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  17, startRow = 2, x = form_A[,1])

#===============================================================================

#NXL Paid DF CL
form_A <- data.frame()
for (i in 1:m){
  #=100*(IF($C2=1,1/PRODUCT(V2:OFFSET(V2,24-$C2,0)),1/PRODUCT(V2:OFFSET(V2,24-$C2,0))-1/PRODUCT(V1:OFFSET(V1,24-$C1,0))))
  temp_form <- paste0("=100*(IF($C",1+i,"=1,1/PRODUCT(V",1+i,":OFFSET(V",1+i,",24-$C",1+i,",0)),1/PRODUCT(V",1+i,":OFFSET(V",1+i,",24-$C",1+i,",0))-1/PRODUCT(V",i,":OFFSET(V",i,",24-$C",i,",0))))")
  #writeFormula(wb, sheet = "LDF", startCol =  12, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  20, startRow = 2, x = form_A[,1])

#NXL Incurred DF CL
form_A <- data.frame()
for (i in 1:m){
  #=100*IF($C2=1,1/PRODUCT(O2:OFFSET(O2,24-$C2,0)),1/PRODUCT(O2:OFFSET(O2,24-$C2,0))-1/PRODUCT(O1:OFFSET(O1,24-$C1,0)))
  temp_form <- paste0("=100*IF($C",1+i,"=1,1/PRODUCT(W",1+i,":OFFSET(W",1+i,",24-$C",1+i,",0)),1/PRODUCT(W",1+i,":OFFSET(W",1+i,",24-$C",1+i,",0))-1/PRODUCT(W",i,":OFFSET(W",i,",24-$C",i,",0)))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  13, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  21, startRow = 2, x = form_A[,1])

#NXL Cummulative Paid DF CL
form_A <- data.frame()
for (i in 1:m){
  #=IF($C2=1,100/L2,100/SUMIFS(L:L,$A:$A,$A2,$B:$B,$B2,$C:$C,"<="&$C2))
  temp_form <- paste0("=IF($C",1+i,"=1,100/T",1+i,",100/SUMIFS(T:T,$A:$A,$A",1+i,",$B:$B,$B",1+i,",$C:$C,\"<=\"&$C",1+i,"))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  16, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  24, startRow = 2, x = form_A[,1])

#NXL Cummulative Incurred DF CL
form_A <- data.frame()
for (i in 1:m){
  #=IF($C2=1,100/M2,100/SUMIFS(M:M,$A:$A,$A2,$B:$B,$B2,$C:$C,"<="&$C2))
  temp_form <- paste0("=IF($C",1+i,"=1,100/U",1+i,",100/SUMIFS(U:U,$A:$A,$A",1+i,",$B:$B,$B",1+i,",$C:$C,\"<=\"&$C",1+i,"))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  17, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
writeFormula(wb, sheet = "LDF", startCol =  25, startRow = 2, x = form_A[,1])
  
}

#===============================================================================


#XL Paid DF CL
form_A <- data.frame()
for (i in 1:m){
  #=100*(IF($C2=1,1/PRODUCT(N2:OFFSET(N2,24-$C2,0)),1/PRODUCT(N2:OFFSET(N2,24-$C2,0))-1/PRODUCT(N1:OFFSET(N1,24-$C1,0))))
  temp_form <- paste0("=100*(IF($C",1+i,"=1,1/PRODUCT(AD",1+i,":OFFSET(AD",1+i,",24-$C",1+i,",0)),1/PRODUCT(AD",1+i,":OFFSET(AD",1+i,",24-$C",1+i,",0))-1/PRODUCT(AD",i,":OFFSET(AD",i,",24-$C",i,",0))))")
  #writeFormula(wb, sheet = "LDF", startCol =  12, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  28, startRow = 2, x = form_A[,1])

#XL Incurred DF CL
form_A <- data.frame()
for (i in 1:m){
  #=100*IF($C2=1,1/PRODUCT(O2:OFFSET(O2,24-$C2,0)),1/PRODUCT(O2:OFFSET(O2,24-$C2,0))-1/PRODUCT(O1:OFFSET(O1,24-$C1,0)))
  temp_form <- paste0("=100*IF($C",1+i,"=1,1/PRODUCT(AE",1+i,":OFFSET(AE",1+i,",24-$C",1+i,",0)),1/PRODUCT(AE",1+i,":OFFSET(AE",1+i,",24-$C",1+i,",0))-1/PRODUCT(AE",i,":OFFSET(AE",i,",24-$C",i,",0)))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  13, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  29, startRow = 2, x = form_A[,1])

#XL Cummulative Paid DF CL
form_A <- data.frame()
for (i in 1:m){
  #=IF($C2=1,100/L2,100/SUMIFS(L:L,$A:$A,$A2,$B:$B,$B2,$C:$C,"<="&$C2))
  temp_form <- paste0("=IF($C",1+i,"=1,100/AB",1+i,",100/SUMIFS(AB:AB,$A:$A,$A",1+i,",$B:$B,$B",1+i,",$C:$C,\"<=\"&$C",1+i,"))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  16, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
}
writeFormula(wb, sheet = "LDF", startCol =  32, startRow = 2, x = form_A[,1])

#XL Cummulative Incurred DF CL
form_A <- data.frame()
for (i in 1:m){
  #=IF($C2=1,100/M2,100/SUMIFS(M:M,$A:$A,$A2,$B:$B,$B2,$C:$C,"<="&$C2))
  temp_form <- paste0("=IF($C",1+i,"=1,100/AC",1+i,",100/SUMIFS(AC:AC,$A:$A,$A",1+i,",$B:$B,$B",1+i,",$C:$C,\"<=\"&$C",1+i,"))")
  temp_row <- 1+i
  #writeFormula(wb, sheet = "LDF", startCol =  17, startRow =  temp_row, x = temp_form)
  form_A <- rbind(form_A, temp_form)
  writeFormula(wb, sheet = "LDF", startCol =  33, startRow = 2, x = form_A[,1])
  
}

#===============================================================================


saveWorkbook(wb, "DevelopmentPaidIncurred.2024Q4.v1.xlsx", overwrite = TRUE)



