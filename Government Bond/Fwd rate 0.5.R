library(dplyr)
library(lubridate)
library(openxlsx)

daftar_nama <- list.files(path="D:/Stanley/Tugas IFRS17/Goverment Bond/Government Bond R/data to stack")
list_tanggal <- as.list(daftar_nama)


all_temp <- data.frame()


for(j in list_tanggal){

#j <- "20231229_YieldByTenor_GB.csv"

nama_input <- paste0("D:/Stanley/Tugas IFRS17/Goverment Bond/Government Bond R/data to stack/",j)
  
input <- read.csv(nama_input)

contoh <- data.frame(input$DATE,input$TENOR.MONTH,input$TODAY)
colnames(contoh) <- c("ValuationDate","Tenor","Spot")
contoh <- contoh[order(contoh$Tenor), ]


df_temp <- data.frame()

new_row_1 <- data.frame(
  FWD = contoh[1,3],
  SpotAdj = contoh[1,3]+0.5,
  FWDAdj = contoh[1,3]+0.5
)
df_temp <- bind_rows(df_temp,new_row_1)

# new_row_2 <- data.frame(
#   FWD = 100*(((1+0.01*contoh[2,3])/((1+0.01*contoh[1,3])^(1/12)))^(12/11)-1)
# )
# df_temp <- bind_rows(df_temp,new_row_2)

# new_row_2 <- data.frame(
#   FWD = contoh[2,3]
# )
# df_temp <- bind_rows(df_temp,new_row_2)



for(i in 1:(length(contoh$Tenor)-1)){
  
  forward <- 100*((((1+0.01*contoh[i+1,3])^(contoh[i+1,2]))/((1+0.01*contoh[i,3])^(contoh[i,2])))^(2)-1)
  forwardAdj <- 100*((((1+0.01*(contoh[i+1,3]+0.5))^((contoh[i+1,2])))/((1+0.01*(contoh[i,3]+0.5))^((contoh[i,2]))))^(2)-1)
  
  new_row <- data.frame(
    FWD = forward,
    SpotAdj = contoh[i+1,3]+0.5,
    FWDAdj = forwardAdj
    
  )
  
  df_temp <- bind_rows(df_temp, new_row)
  
  
  
}

df_fwd <- cbind(contoh,df_temp)

for(k in 1:20){
  new_row <- c(df_fwd[60,1],df_fwd[59+k,2]+0.5,df_fwd[60,3],df_fwd[60,4],df_fwd[60,5],df_fwd[60,6])
  df_fwd <- rbind(df_fwd, new_row)
}

all_temp <- rbind(all_temp, df_fwd)



#all_temp$SpotNew <- as.numeric(all_temp$Spot) + 0.5



}

'Spot_Adj <- data.frame()

for (m in 1:length(all_temp$Spot)){
  
  temp <- all_temp[m,3]+0.5
  new_row <- data.frame(SpotNew = temp)
  Spot_Adj <- rbind(Spot_Adj, new_row)
}

all_temp <- cbind(all_temp,Spot_Adj)'

all_temp$ValuationDate <- as.character(all_temp$ValuationDate)

all_temp$ValuationDate <- as.Date.character(all_temp$ValuationDate, "%Y%m%d")

all_temp$ValuationDate <- ceiling_date(all_temp$ValuationDate, unit = "month")-1

colnames(all_temp) <- c("Valuation Date", "Tenor", "Spot", "Forward", "Spot + Illiquidity Premium", "Forward + Illiquidity Premium")

currency <- data.frame (Currency = 1:length(all_temp$`Valuation Date`))

currency <- "IDR"

all_temp <- cbind("Valuation Date"=all_temp[,1],"Currency Code"=currency,all_temp[,2:6])

write.csv(all_temp,"InterestRate_IDR.csv", row.names = FALSE )

