library(dplyr)
library(lubridate)
library(readxl)
library(stringr)


#kolom_cek <- read_excel("data EIOPA/EIOPA_RFR_20151231_Term_Structures.xlsx", sheet = "RFR_spot_no_VA", col_names = FALSE, range = "AM2")

column_letter <- function(column_number) {
  letter <- ""
  
  while (column_number > 0) {
    remainder <- (column_number - 1) %% 26
    letter <- paste0(LETTERS[remainder + 1], letter)
    column_number <- (column_number - 1) %/% 26
  }
  
  return(letter)
}

tenor <- data.frame("Tenor" = seq(1,40))

data_list <- as.list(list.files(path = "data EIOPA"))

map_cur <- read.csv("map_country_currency.csv")

all_data <- data.frame()

for(i in data_list){
  val_date <- as.character(substr(i, 11, 18))
  
  valuation <- data.frame(ValuationDate = replicate(40, val_date))
  
  all_curr_df <- data.frame()
  
  map_cur <- map_cur[,-3]
  
  index_cou <- list(matrix(read_excel(paste0("data EIOPA/",xxx), sheet = "RFR_spot_no_VA", range="C2:CC2", col_names = FALSE)))
  colnames(index_cou) <- "X1"
  
  for (y in map_cur$Country){
    
    indices <- tryCatch({
          
             column_letter(which(sapply(index_cou, function(x) x == y))+2)
       }, error = function(e) {
             # If an error occurs, return NA
               return(NA)
         })
    print(indices)
    
    map_cur$ColumnName[map_cur$Country == y] <- indices
    
    #ifelse(length(map_cur$Country == ))
    
  }
  
  map_cur <- na.omit(map_cur)

  
  for (j in map_cur$ColumnName){
    
    currency_name <- data.frame(Currency = replicate(40, map_cur[map_cur$ColumnName == as.character(j), "CurrencyCode"]))
    
    nama_excel <- paste0("data EIOPA/", as.character(i))
    spot_EIOPA <- read_excel(nama_excel, sheet = "RFR_spot_no_VA", range = paste0(as.character(j),"10:",as.character(j),"50"), col_types = "numeric")
    colnames(spot_EIOPA) <- "Spot"
    spot_EIOPA <- spot_EIOPA*100
    
    final_spot <- cbind(valuation, tenor, spot_EIOPA)
    
    final_spot <- final_spot[order(final_spot$Tenor), ]
    
    
    df_temp <- data.frame()
    
    new_row_1 <- data.frame(
      FWD = final_spot[1,3],
      SpotAdj = final_spot[1,3]+0.5,
      FWDAdj = final_spot[1,3]+0.5
    )
    df_temp <- bind_rows(df_temp,new_row_1)
    
    for(k in 1:(length(final_spot$Tenor)-1)){
      
      forward <- 100*((((1+0.01*final_spot[k+1,3])^(final_spot[k+1,2]))/((1+0.01*final_spot[k,3])^(final_spot[k,2])))-1)
      forwardAdj <- 100*((((1+0.01*(final_spot[k+1,3]+0.5))^(final_spot[k+1,2]))/((1+0.01*(final_spot[k,3]+0.5))^(final_spot[k,2])))-1)
      
      
      new_row <- data.frame(
        FWD = forward,
        SpotAdj = final_spot[k+1,3]+0.5,
        FWDAdj = forwardAdj
        
      )
      
      df_temp <- bind_rows(df_temp, new_row)
      
      
    }
    
    df_fwd <- cbind(final_spot[,1],currency_name[,1],final_spot[,2:3],df_temp)
    
    all_curr_df <- rbind(all_curr_df, df_fwd)
    
  }
  
  all_data <- rbind(all_data, all_curr_df)
  
}



colnames(all_data) <- c("Valuation Date", "Currency Code", "Tenor", "Spot", "Forward", "Spot + Illiquidity Premium", "Forward + Illiquidity Premium")

all_data$`Valuation Date` <- as.Date(all_data$`Valuation Date`, format = "%Y%m%d")

filter_CLP <- all_data[all_data$`Currency Code` == "CLF",]
filter_CLP$`Currency Code` <- "CLP"

filter_RMB <- all_data[all_data$`Currency Code` == "CNY",]
filter_RMB$`Currency Code` <- "RMB"

filter_COU <- all_data[all_data$`Currency Code` == "COP",]
filter_COU$`Currency Code` <- "COU"

all_data <- rbind(all_data, filter_CLP, filter_RMB, filter_COU)
all_data <- all_data[order(all_data$`Currency Code`),]

write.csv(all_data,"InterestRate_FOREIGN.csv", row.names = FALSE)

  
