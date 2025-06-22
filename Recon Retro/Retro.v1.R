library(dplyr)
library(readxl)
library(tidyr)

data <- read.csv("Input/Premium.csv")

data$IssueDate <- as.Date(data$IssueDate, format = "%m/%d/%Y")
data$DueDate <- as.Date(data$DueDate, format = "%m/%d/%Y")
data$PaidDate <- as.Date(data$PaidDate, format = "%m/%d/%Y")

cf_prem_long <- pivot_longer(data = data,
                                  cols = c("Premium", "Commission"),
                                  names_to = "Cashflow",
                                  values_to = "Amount")

#Cashflow Expected
cf_exp_prem_group <- cf_prem_long %>% group_by(Cashflow, IFRSGroup, DueDate) %>% summarise(Amount = sum(Amount))
cf_exp_prem <- cf_exp_prem_group[order(cf_exp_prem_group$IFRSGroup, cf_exp_prem_group$DueDate, desc(cf_exp_prem_group$Cashflow)),]

#Cashflow Paid
cf_paid_prem_group <- cf_prem_long %>% group_by(Cashflow, IFRSGroup, PaidDate) %>% summarise(Amount = sum(Amount))
cf_paid_prem <- cf_paid_prem_group[order(cf_paid_prem_group$IFRSGroup, cf_paid_prem_group$PaidDate, desc(cf_paid_prem_group$Cashflow)),]

