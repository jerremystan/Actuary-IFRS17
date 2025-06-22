library(dplyr)
library(readxl)
library(DBI)
library(RSQLite)
library(odbc)

con <- dbConnect(odbc(),
                 Driver = "SQL Server",
                 Server = "ARISTYO-NB",
                 Database = "TMI.Actuary",
                 UID = "sa",
                 PWD = rstudioapi::askForPassword(),
                 Port = 1433
)

#query_M <- paste0("")

#data_EP_M <- dbGetQuery(con, query_M)

data_group <- read_excel("Sprint.Result.20240422.202301.20221231.01.xlsx", 
                         sheet = "Policy")
data_group$PolicyTrans <- paste0(data_group$`Policy No`, data_group$`Policy Transaction No`)

group_unique <- unique(data_group$IFRSGroupCode)

for (i in group_unique){
  
  data_value <- filter(data_group,data_group$IFRSGroupCode == i)
  temp_PolicyTrans <- data_value$PolicyTrans
  temp_sql <- paste0("'",paste(temp_PolicyTrans, collapse = "','"),"'")
  
  query_temp <- paste0("SELECT *, concat(prteip.PolicyNo, prteip.PolicyTransactionNo) as PolicyNoTrans FROM dbo.PolicyRiskTransactionEarnedInwardPremium prteip WHERE concat(prteip.PolicyNo, prteip.PolicyTransactionNo) in (",temp_sql,")")
  
  EP_temp <- dbGetQuery(con, query_temp)
  
  EP_name_temp <- paste0("EP_",i)
  
  #assign(EP_name_temp, EP_temp)
  
  write.csv(EP_temp, paste0("Earned Premium/",EP_name_temp,".csv"), row.names = FALSE)
  
}



#new_group_A <- paste0("'",paste(group_A, collapse = "','"),"'")
#print(new_group_A)

#dbDisconnect(con)
