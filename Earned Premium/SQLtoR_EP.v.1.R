library(dplyr)
library(readxl)
library(DBI)
library(RSQLite)
library(odbc)

con <- dbConnect(odbc(),
                 Driver = "SQL Server",
                 Server = "CONFIDENTIAL",
                 Database = "CONFIDENTIAL",
                 UID = "sa",
                 PWD = rstudioapi::askForPassword(),
                 Port = 1433
)

tahunSQL <- 2023

query_M <- paste(
  "
SELECT y.EPMonth as 'EP - Month', pt.MajorClassCode as 'Major Class Code', pt.ContractTypeCode as 'Contract Type Code', pt.ProfitCenterCode as 'Profit Center Code', pt.ClientID as 'Client ID', SUM(y.EarnedPremium) as 'Earned Premium Aft Co-Out'
FROM dbo.PolicyTransactions pt
	INNER JOIN (
	SELECT x.PolicyNo, x.PolicyTransactionNo, x.EPMonth, SUM(x.EarnedPremium) as EarnedPremium
	FROM(
	SELECT prtep.PolicyNo, prtep.PolicyTransactionNo, FORMAT(prtep.EarnedDate, 'yyyyMM') AS EPMonth, prtep.RiskNo, SUM(prtep.Amount) AS EarnedPremium
	FROM dbo.PolicyRiskTransactionEarnedInwardPremium prtep
	WHERE YEAR(prtep.EarnedDate) =",tahunSQL,"
	GROUP BY prtep.PolicyNo, prtep.PolicyTransactionNo, prtep.RiskNo, prtep.EarnedDate
	)x
	GROUP BY x.PolicyNo, x.PolicyTransactionNo, x.EPMonth
	)y
	ON (pt.PolicyNo = y.PolicyNo AND pt.PolicyTransactionNo = y.PolicyTransactionNo)
WHERE pt.MajorClassCode = 'M'
GROUP BY y.EPMonth, pt.MajorClassCode, pt.ContractTypeCode, pt.ProfitCenterCode, pt.ClientID
ORDER BY y.EPMonth, pt.MajorClassCode, pt.ContractTypeCode, pt.ProfitCenterCode, pt.ClientID

"
)

#data_EP_M <- read.csv("Input/EP for Actual Expense - M.v2.csv", colClasses = c("character"))

data_EP_M <- dbGetQuery(con, query_M)

map_PC <- read_excel("Mapping/ProfitCenterCode.v6.xlsx") #Karena summary opex masih dari v6

colnames(map_PC)[which(colnames(map_PC)=="Profit Center Code")] <- "Profit.Center.Code"

ME_client <- list("1","2","3") #diganti karena confidential

data_EP_M$ReservingClass <- ifelse( data_EP_M$`Contract Type Code` == "XXX" | data_EP_M$`Contract Type Code` == "YYY", 
                                    ifelse(data_EP_M$`Client ID` %in% ME_client, "ZE", "ZO"), "ZO")

query_AV <- paste("
SELECT y.EPMonth as 'EP - Month', pt.MajorClassCode as 'Major Class Code', pt.ContractTypeCode as 'Contract Type Code', pt.ProfitCenterCode as 'Profit Center Code', SUM(y.EarnedPremium) as 'Earned Premium Aft Co-Out'
FROM dbo.PolicyTransactions pt
	INNER JOIN (
	SELECT x.PolicyNo, x.PolicyTransactionNo, x.EPMonth, SUM(x.EarnedPremium) as EarnedPremium
	FROM(
	SELECT prtep.PolicyNo, prtep.PolicyTransactionNo, FORMAT(prtep.EarnedDate, 'yyyyMM') AS EPMonth, prtep.RiskNo, SUM(prtep.Amount) AS EarnedPremium
	FROM dbo.PolicyRiskTransactionEarnedInwardPremium prtep
	WHERE YEAR(prtep.EarnedDate) =", tahunSQL, "
	GROUP BY prtep.PolicyNo, prtep.PolicyTransactionNo, prtep.RiskNo, prtep.EarnedDate
	)x
	GROUP BY x.PolicyNo, x.PolicyTransactionNo, x.EPMonth
	)y
	ON (pt.PolicyNo = y.PolicyNo AND pt.PolicyTransactionNo = y.PolicyTransactionNo)
WHERE pt.MajorClassCode <> 'M'
GROUP BY y.EPMonth, pt.MajorClassCode, pt.ContractTypeCode, pt.ProfitCenterCode
ORDER BY y.EPMonth, pt.MajorClassCode, pt.ContractTypeCode, pt.ProfitCenterCode

")

#data_EP_AV <- read.csv("Input/EP for Actual Expense - exclude M.v2.csv", colClasses = c("character"))

data_EP_AV <- dbGetQuery(con, query_AV)

data_EP_AV$ReservingClass <- ifelse(data_EP_AV$`Major Class Code` == "V",
                                    ifelse(data_EP_AV$`Contract Type Code` == "BYP", "VE", "VO"),
                                    ifelse(data_EP_AV$`Major Class Code` == "H","MO",data_EP_AV$`Major Class Code` ))

intersect <- intersect(names(data_EP_M),names(data_EP_AV)) # return kesamaan antara 2 vectors

data_EP <- rbind(data_EP_AV[intersect], data_EP_M[intersect])

colnames(map_PC) <- c("Profit Center Code", "Profit Center T-Code", "Profit Center Long Name")

data_EP <- left_join(data_EP,map_PC,by = "Profit Center Code")

data_EP <- data_EP[,c(1,6,8,2:5,7)]

data_EP <- data_EP[,c(-4,-5,-6,-8)]

#write.csv(data_EP, "data_EP.csv", row.names = FALSE)

data_EP_final <- data_EP %>% group_by(ReservingClass, "Profit Center Code" = data_EP[,3]) %>% 
                  summarise("EP Aft Co-Out" = sum(as.numeric(`Earned Premium Aft Co-Out`)))

'if (length(group_vars(data_EP_final)) == 0) {
  print("Data is ungrouped")
} else {
  print("Data is grouped")
}
'
write.csv(data_EP_final,"Output/data_EP_final_SQL.csv", row.names = FALSE)

dbDisconnect(con)




