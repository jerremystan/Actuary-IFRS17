library(dplyr)
library(lubridate)
library(stringr)
library(readxl)
library(openxlsx)

#CREATING NEW MASTER DATA

list_ibnr <- as.list(list.files(path = "IBNR"))

  
ielr_data <- read_excel("IELR_RiskMargin_DiversificationBenefits.v3.xlsx", 
                          sheet = "Actuarial")

ielr_data$`Valuation Date` <- as.character(ielr_data$`Valuation Date`)

reserve_class <- unique(ielr_data$`Reserving Class`)

for(i in list_ibnr){
  
  val_date <- substr(i,6,15)
  
  gross_sum <- read_excel(path = paste0("IBNR/",as.character(i)), sheet = "Summary - Gross (Group)")
  net_sum <- read_excel(path = paste0("IBNR/",as.character(i)), sheet = "Summary - Net (Group)")
  index <- data.frame(bottom = which(gross_sum[,1] == "Total"), top = which(gross_sum[,1] == "Class of Business"))
  index$top <- index$top
  index$bottom <- index$bottom - 1
  
  #GROSS
  gross_cl <- gross_sum[index[1,2]:index[1,1], 1:15]
  colnames(gross_cl) <- unlist(gross_cl[1,])
  gross_cln <- gross_cl[-1:-2,]
  colnames(gross_cln)[1] <- "Class of Business Gross"
  gross_pl <- gross_sum[index[2,2]:index[2,1], 1:17]
  colnames(gross_pl) <- unlist(gross_pl[1,])
  gross_pln <- gross_pl[-1:-2,]
  colnames(gross_pln)[1] <- "Class of Business Gross"
  
  #NET
  net_cl <- net_sum[index[1,2]:index[1,1], 1:15]
  colnames(net_cl) <- unlist(net_cl[1,])
  net_cln <- net_cl[-1:-2,]
  colnames(net_cln)[1] <- "Class of Business Net"
  net_pl <- net_sum[index[2,2]:index[2,1], 1:17]
  colnames(net_pl) <- unlist(net_pl[1,])
  net_pln <- net_pl[-1:-2,]
  colnames(net_pln)[1] <- "Class of Business Net"
  
  new_dfx <- data.frame(`Valuation Date` = rep(val_date,length(reserve_class)), `Reserving Class` = reserve_class)
  
  classob <- unique(net_cln$`Class of Business Net`)
  classob_grs <- unique(gross_cln$`Class of Business Gross`)
  classob <- classob[c(-6,-8,-10)]
  classob_grs <- classob_grs[c(-6,-8,-10)]
  new_dfx <- cbind(new_dfx, `Class of Business Gross` = classob_grs, `Class of Business Net` = classob)
  
  join_df <- left_join(new_dfx, gross_cln, by = "Class of Business Gross")
  join_df <- left_join(join_df, gross_pln, by = "Class of Business Gross")
  join_df <- join_df[,c(1,2,3,4,20,28,13,15)]
  colnames(join_df) <- c(colnames(join_df)[1:4],"IELR Gross", "Premium Liability", "Claim Liability", "Diversification Benefit")
  join_df <- left_join(join_df, net_pln, by = "Class of Business Net")
  join_df <- join_df[,c(1:5,10,6,7,8)]
  colnames(join_df)[6] <- "IELR Net"
  join_df$`IELR Reins` <- 0
  join_df <- join_df[,c(1,2,5,6,10,7,8,9)]
  join_df$`IELR Gross` <- as.numeric(join_df$`IELR Gross`)*100
  join_df[,4] <- as.numeric(join_df[,4])*100
  join_df[,6] <- as.numeric(join_df[,6])*100
  join_df[,7] <- as.numeric(join_df[,7])*100
  join_df[,8] <- as.numeric(join_df[,8])*100
  colnames(join_df)[1:2] <- c("Valuation Date", "Reserving Class")
  
  ielr_data <- rbind(ielr_data, join_df)
  
}



team_A <- data.frame(ValuationDate = as.character(seq(as.Date("1998-04-01"), as.Date("2015-10-30"), by = "3 months")-1))
team_B <- data.frame(ValuationDate = c("2016-03-31","2016-06-30","2016-09-30"))
team_C <- data.frame(ValuationDate = c("2017-03-31","2017-06-30","2017-09-30"))
team_D <- data.frame(ValuationDate = c("2018-03-31","2018-06-30","2018-09-30"))

filter_A <- filter(ielr_data, ielr_data$`Valuation Date` == "2015-12-31")
filter_B <- filter(ielr_data, ielr_data$`Valuation Date` == "2015-12-31")
filter_C <- filter(ielr_data, ielr_data$`Valuation Date` == "2016-12-31")
filter_D <- filter(ielr_data, ielr_data$`Valuation Date` == "2017-12-31")

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

all_data <- rbind(ielr_data, df_A, df_B, df_C, df_D)
all_data <- all_data[order(all_data$`Valuation Date`),]

df_temp <- data.frame()
baris <- as.numeric(nrow(all_data))

for(i in 1:baris){
  #i <- 1
  tanggal <- as.character(all_data[i,1])
  tahun <- as.numeric(substr(tanggal,1,4))
  issue <- data.frame(seq(1998,tahun))
  n <- as.numeric(length(issue))
  
  new_df <- data.frame(
   ValuationDate = replicate(n,all_data[i,1]),
   RC = replicate(n, all_data[i,2]),
   IssueYear = issue,
   Grs = replicate(n, all_data[i,3]),
   Net = replicate(n, all_data[i,4]),
   RI = replicate(n, all_data[i,5]),
   Prem = replicate(n, all_data[i,6]),
   Claim = replicate(n, all_data[i,7]),
   DB = replicate(n, all_data[i,8])
    
  )
  colnames(new_df) <- c("ValuationDate", "ReservingClass", "IssueYear","Gross","Net","Reins", "Prem", "Claim", "DB")
  
  df_temp <- bind_rows(df_temp, new_df)
  
}

colnames(df_temp) <- c("Valuation Date", "Reserving Class Code", "Issue Year","Gross Loss Ratio","Net Loss Ratio","Reinsurance Loss Ratio", "Premium Risk Margin","Claim Risk Margin","Diversification Benefit")

filter_year <- df_temp[substr(df_temp$`Valuation Date`,6,7) == "12" & as.numeric(substr(df_temp$`Valuation Date`,1,4)) == df_temp$`Issue Year` ,]
filter_year$`Issue Year` <- filter_year$`Issue Year`+1

df_temp <- rbind(df_temp, filter_year)

df_temp <- df_temp[order(df_temp$`Valuation Date`,df_temp$`Reserving Class Code`,df_temp$`Issue Year`),]

LR_final <- df_temp[,c(1,2,3,4,5,6)]
RM_final <- df_temp[,c(1,2,3,7,8,9)]

write.xlsx(LR_final, "LossRatio.v7.xlsx", sheetName = "IELR", rowNames = FALSE)
write.xlsx(RM_final, "RiskMargin_DiversificationBenefit.v15.xlsx", sheetName = "Actual", rowNames = FALSE)
