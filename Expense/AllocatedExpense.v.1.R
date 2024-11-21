library(readxl)
library(dplyr)
library(openxlsx)

vec_channel <- unlist(read_excel("map_Channel.xlsx"))
vec_JL <- unlist(read_excel("map_JL.xlsx"))
vec_partner <- unlist(read_excel("map_Partner Name.xlsx"))
df_profit <- read_excel("map_Profit Center.xlsx")
vec_profit <- unique(unlist(df_profit$`Profit Center Map`))
df_main <- read_excel("map_Main Class.xlsx")
vec_main <- unlist(df_main$`Main Class Map`)

opex_all <- read_excel("SummaryOpex.2023.xlsx", sheet = "All")
opex_add <- read_excel("SummaryOpex.2023.xlsx", sheet = "Additional")

opex_all[is.na(opex_all)] <- 0
opex_add[is.na(opex_add)] <- 0


opex_all <- left_join(opex_all, df_main, by = c("Main Class"))
opex_all$`Main Class`<- opex_all$`Main Class Map`
opex_all <- opex_all[,-ncol(opex_all)]

opex_add <- left_join(opex_add, df_main, by = c("Main Class"))
opex_add$`Main Class`<- opex_add$`Main Class Map`
opex_add <- opex_add[,-ncol(opex_add)]

opex_all <- left_join(opex_all, df_profit, by = c("ProfitCenter"))
opex_all$`ProfitCenter`<- opex_all$`Profit Center Map`
opex_all <- opex_all[,-ncol(opex_all)]

opex_add <- left_join(opex_add, df_profit, by = c("ProfitCenter"))
opex_add$`ProfitCenter`<- opex_add$`Profit Center Map`
opex_add <- opex_add[,-ncol(opex_add)]


# vec_channel <- unique(unlist(df_opex_all$Channel))
# vec_JL <- unique(unlist(df_opex_all$JL))
# vec_profit <- unique(unlist(df_opex_all$ProfitCenter))
# vec_main <- unique(unlist(df_opex_all$`Main Class`))


df_opex_all <- opex_all %>% group_by(JL, `Main Class`, Channel, ProfitCenter) %>% summarise(ACE = sum(ACE), PME = sum(PME), CHE = sum(CHE))
df_opex_add <- opex_add %>% group_by(JL, `Main Class`, Channel, ProfitCenter, PartnerName) %>% summarise(ACE = sum(ACE), PME = sum(PME), CHE = sum(CHE))

df_opex_all[is.na(df_opex_all)] <- 0
df_opex_add[is.na(df_opex_add)] <- 0


expand_vec_all <- expand.grid(`JL` = vec_JL, `Main Class` = vec_main, Channel = vec_channel, `ProfitCenter` = vec_profit)

expand_vec_add <- expand.grid(`JL` = vec_JL, `Main Class` = vec_main, Channel = vec_channel, `ProfitCenter` = vec_profit, `PartnerName` = vec_partner)

df_all <- left_join(expand_vec_all, df_opex_all, by = c("JL", "Main Class", "Channel", "ProfitCenter"))
df_all[is.na(df_all)] <- 0
df_all <- df_all %>% arrange(JL, `Main Class`, Channel, ProfitCenter)
df_all$`Main Class` <- toupper(df_all$`Main Class`)
colnames(df_all) <- c("JL", "Main Class", "Channel", "Profit Center", "Acquisition Cost Expense", "Policy Maintenance Expense", "Claim Handling Expense")

df_add <- left_join(expand_vec_add, df_opex_add, by = c("JL", "Main Class", "Channel", "ProfitCenter", "PartnerName"))
df_add[is.na(df_add)] <- 0
df_add <- df_add %>% arrange(JL, `Main Class`, Channel, ProfitCenter, PartnerName)
df_add$`Main Class` <- toupper(df_add$`Main Class`)
colnames(df_add) <- c("JL", "Main Class", "Channel", "Profit Center", "Partner Name", "Acquisition Cost Expense", "Policy Maintenance Expense", "Claim Handling Expense")

wb <- createWorkbook()

addWorksheet(wb, "YTD-All")
addWorksheet(wb, "YTD-Additional")

writeData(wb, sheet = "YTD-All", df_all)
writeData(wb, sheet = "YTD-Additional", df_add)

saveWorkbook(wb, "ExpenseAllocated.YTD.2023.v.2.xlsx", overwrite = TRUE)
