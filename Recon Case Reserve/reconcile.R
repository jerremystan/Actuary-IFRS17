library(tidyverse)
library(openxlsx)
library(readxl)

ifrs_group_cf_1 <- read_csv("C:/TMI/IFRS 17/UAT/202312/RiskIntegrity/202312.2023/ifrs_group_cashflow_all_runs_2023_202312_1.csv")
ifrs_group_cf_2 <- read_csv("C:/TMI/IFRS 17/UAT/202312/RiskIntegrity/202312.2023/ifrs_group_cashflow_all_runs_2023_202312_2.csv")
ifrs_group_cf_3 <- read_csv("C:/TMI/IFRS 17/UAT/202312/RiskIntegrity/202312.2023/ifrs_group_cashflow_all_runs_2023_202312_3.csv")
ifrs_group_cf_4 <- read_csv("C:/TMI/IFRS 17/UAT/202312/RiskIntegrity/202312.2023/ifrs_group_cashflow_all_runs_2023_202312_4.csv")
ifrs_group_cf_5 <- read_csv("C:/TMI/IFRS 17/UAT/202312/RiskIntegrity/202312.2023/ifrs_group_cashflow_all_runs_2023_202312_5.csv")
ifrs_group_cf_6 <- read_csv("C:/TMI/IFRS 17/UAT/202312/RiskIntegrity/202312.2023/ifrs_group_cashflow_all_runs_2023_202312_6.csv")
ifrs_group_cf <- rbind(ifrs_group_cf_1,ifrs_group_cf_2,ifrs_group_cf_3,ifrs_group_cf_4,ifrs_group_cf_5,ifrs_group_cf_6)

##expense_assumptions <- read_excel("C:/TMI/IFRS 17/UAT/shared/ActualExpense.v2.xlsx", sheet = "2023")
expense_assumptions2 <- read_excel("C:/TMI/IFRS 17/UAT/shared/Expense Study.2023.v.3.xlsx", sheet = "Encoding Unit Cost Driver")

# problems(ifrs_group_cf_1)
# print(problems(ifrs_group_cf_1),n = 20)
# spec(ifrs_group_cf_1)

ifrs_group_long <- ifrs_group_cf %>% 
  select(-subgroup,-custom_dim_2,-incurred_date) %>% 
  pivot_longer(cols = c(-ifrs_group_code,-underlying_ifrs_group_code,-run_number,-variable_name,-transaction_currency_code,-custom_dim_1), names_to = "cf_date", values_to = "value")

# cek <- ifrs_group_long %>% 
#   filter(is_empty(value))

total_run <- ifrs_group_long %>% 
  summarise(total_cashflow = sum(value,na.rm = TRUE),
            .by = c(run_number,variable_name,custom_dim_1))

total_run <- total_run %>% 
  arrange(run_number,variable_name,custom_dim_1)

total_cf <- total_run %>% 
  filter(run_number %in% c(20,40,84,90)) %>%
  mutate(period_cf = case_when(run_number %in% c(20,84) ~ "Current",
                               run_number %in% c(40,90) ~ "Future",
                               .default = "NA")
         ) %>% 
  summarise(total_cashflow = sum(total_cashflow, na.rm = TRUE),
            .by = c(variable_name,period_cf))

write.csv(total_run,"total_run.csv", row.names = FALSE)
write.csv(total_cf,"total_cashflow.csv", row.names = FALSE)

# transaction_data <- read.xlsx("C:/TMI/IFRS 17/UAT/Dataset/InsuranceContracts.2023.v.5.xlsx", detectDates = TRUE)
transaction_data <- read_excel("C:/TMI/IFRS 17/UAT/Dataset/InsuranceContracts.2023.v.8..xlsx")

policy_data <- transaction_data %>% 
  mutate(ReservingClassCode = case_when(`Reserving Class`=="AUTO" ~ "A",
                                        `Reserving Class`=="MARINE OTHERS" ~ "MO",
                                        `Reserving Class`=="MARINE E-COMMERCE" ~ "ME",
                                        `Reserving Class`=="FIRE" ~ "F",
                                        `Reserving Class`=="ENGINEERING" ~ "E",
                                        `Reserving Class`=="LIABILITY" ~ "L",
                                        `Reserving Class`=="PA" ~ "P",
                                        `Reserving Class`=="TRAVEL" ~ "T",
                                        `Reserving Class`=="TRADE CREDIT" ~ "C",
                                        `Reserving Class`=="HEALTH" ~ "S",
                                        `Reserving Class`=="VARIOUS OTHERS" ~ "VO",
                                        .default = "NA"),
         IsShortTerm = case_when(`ST/LT`=="ST" ~ 1,
                                 `ST/LT`=="LT" ~ 0,
                                 .default = 2),
         Transaction_year = year(`Transaction Issue Date`),
         Period = case_when(`Due Date`>as.Date('2023-12-31') ~ "Future",
                            .default = "Current"),
         Recognition_year = case_when(year(RecognitionDate)<=2023 ~ "<= 2023",
                                      .default = as.character(year(RecognitionDate))) 
         )

# cek <- policy_data %>% 
#   filter(ReservingClassCode == "NA" | IsShortTerm == 2)

ULR_assumption <- read_excel("C:/TMI/IFRS 17/UAT/shared/ULR.v10.xlsx", sheet = "ULR")
ULR_assumption <- ULR_assumption %>% 
  filter(`Valuation Date`==as.Date('2023-12-31'))

# expense_assumptions$`Valuation Date` <- as.Date(expense_assumptions$`Valuation Date`)
expense_assumptions2 <- expense_assumptions2 %>% 
  select(-`Reserving Class`,-`Profit Center`)

transaction_mapping <- policy_data %>% 
  left_join(ULR_assumption, by = c("Cohort Year"="Cohort Year","ReservingClassCode"="Reserving Class Code","IsShortTerm"="IsShortTerm")) %>% 
  left_join(expense_assumptions, by = c("Valuation Date"="Valuation Date","Profit Center Code"="Profit Center Code",
                                        "ReservingClassCode"="Reserving Class Code","Cohort Year"="Issue Year","IsShortTerm"="Is Short Term"))
  
transaction_mapping2 <- policy_data %>% 
  left_join(ULR_assumption, by = c("Cohort Year"="Cohort Year","ReservingClassCode"="Reserving Class Code","IsShortTerm"="IsShortTerm")) %>% 
  left_join(expense_assumptions2, by = c("Profit Center Code"="Profit Center TCode",
                                        "ReservingClassCode"="Reserving Class Code"))
# names(expense_assumptions2)
# cek <- transaction_mapping2 %>%
#   filter(is.na(`Gross ULR`))
# unique(transaction_mapping$Transaction_year)
# names(transaction_mapping)

transaction_calculation <- transaction_mapping %>% 
  mutate(Claim = Premium*`Gross ULR`/100,
         AQE_cf = Premium*ACE/100,
         PME_cf = Premium*PME/100,
         CHE_cf = Claim*CHE/100)

transaction_calculation2 <- transaction_mapping2 %>% 
  mutate(Claim = Premium*`Gross ULR`/100,
         AQE_cf = Premium*`%ACE`/100,
         PME_cf = Premium*`% PME`/100,
         CHE_cf = Claim*`% CHE`/100)


transaction_goc <- transaction_calculation %>% 
  summarise(Premium = sum(Premium,na.rm = TRUE),
            Commission = sum(AGRICommission,na.rm = TRUE)+sum(OCPORCommission,na.rm = TRUE),
            Policy_Cost = sum(`Policy Cost`,na.rm = TRUE),
            Claim = sum(Claim,na.rm = TRUE),
            AQE = sum(AQE_cf,na.rm = TRUE),
            PME = sum(PME_cf,na.rm = TRUE),
            CHE = sum(CHE_cf,na.rm = TRUE),
            .by = c(`IFRS Group Code`,`Method Code`,`Currency Code`,Transaction_year,Period))

transaction_goc2 <- transaction_calculation2 %>% 
  summarise(Premium = sum(Premium,na.rm = TRUE),
            Commission = sum(AGRICommission,na.rm = TRUE)+sum(OCPORCommission,na.rm = TRUE),
            Policy_Cost = sum(`Policy Cost`,na.rm = TRUE),
            Claim = sum(Claim,na.rm = TRUE),
            AQE = sum(AQE_cf,na.rm = TRUE),
            PME = sum(PME_cf,na.rm = TRUE),
            CHE = sum(CHE_cf,na.rm = TRUE),
            .by = c(`IFRS Group Code`,`Method Code`,`Currency Code`,Transaction_year,Recognition_year,Period))


transaction_summary <- transaction_goc %>% 
  pivot_longer(cols = c(Premium:CHE), names_to = "cashflow", values_to = "value")

transaction_summary2 <- transaction_goc2 %>% 
  pivot_longer(cols = c(Premium:CHE), names_to = "cashflow", values_to = "value")

transaction_total <- transaction_summary %>% 
  summarise(total_transaction = sum(value,na.rm = TRUE),
            .by = c(Transaction_year,cashflow)) %>% 
  arrange(Transaction_year,cashflow)

transaction_total2 <- transaction_summary2 %>% 
  summarise(total_transaction = sum(value,na.rm = TRUE),
            .by = c(Recognition_year,Transaction_year,cashflow)) %>% 
  arrange(Recognition_year,Transaction_year,cashflow)

write.csv(transaction_total2,"total_transaction.csv", row.names = FALSE)

################################## Check Amortization
Amortization_data <- read_csv("C:/TMI/IFRS 17/UAT/Dataset/Amortization Premium Commission Policy Cost.2023.v.1.csv")
PME_AQE_data <- read_csv("C:/TMI/IFRS 17/UAT/Dataset/PME AQE.2023.v.1.csv")

## Amortization Premium, Commission, Policy Cost
Amortization_summary <- Amortization_data %>% 
  mutate(Transaction_year = year(`Transaction Issue Date`),
         Recognition_year = year(RecognitionDate)) %>% 
  summarise(Premium = sum(Premium,na.rm = TRUE),
            Commission = sum(AGRICommission,na.rm = TRUE)+sum(OCPORCommission,na.rm = TRUE),
            Policy_Cost = sum(`Policy Cost`,na.rm = TRUE),
            .by = c(Recognition_year,Transaction_year))

Amortization_total <- Amortization_summary %>% 
  pivot_longer(cols = c(Premium:Policy_Cost), names_to = "cashflow", values_to = "value")

write.csv(Amortization_total,"total_amortization.csv", row.names = FALSE)

## Expected PME AQE
PME_AQE_summary <- PME_AQE_data %>% 
  mutate(Transaction_year = year(`Transaction Issue Date`),
         Recognition_year = year(RecognitionDate)) %>% 
  summarise(AQE = sum(`Acquisition Cost Expense`,na.rm = TRUE),
            PME = sum(`Policy Maintenance Expense`,na.rm = TRUE),
            .by = c(Recognition_year,Transaction_year))

PME_AQE_total <- PME_AQE_summary %>% 
  pivot_longer(cols = c(AQE,PME), names_to = "cashflow", values_to = "value")

write.csv(PME_AQE_total,"total_PME_AQE.csv", row.names = FALSE)

## Amortization input csv
variable_csv_1 <- read_csv("C:/TMI/IFRS 17/UAT/202312/RiskIntegrity/202312.2023/variable_all_runs_2023_202312_1.csv")
variable_csv_2 <- read_csv("C:/TMI/IFRS 17/UAT/202312/RiskIntegrity/202312.2023/variable_all_runs_2023_202312_2.csv")
variable_csv <- rbind(variable_csv_1,variable_csv_2)

variable_long <- variable_csv %>% 
  select(-subgroup,-custom_dim_2) %>% 
  pivot_longer(cols = c(-ifrs_group_code,-underlying_ifrs_group_code,-run_number,-variable_name,-transaction_currency_code,-custom_dim_1), names_to = "cf_date", values_to = "value")

total_variable <- variable_long %>% 
  summarise(total = sum(value,na.rm = TRUE),
            .by = c(run_number,variable_name,custom_dim_1))

write.csv(total_variable,"total_variable.csv", row.names = FALSE)

##### Check Premium Commission IFRS Group
transaction_ifrs_group <- transaction_summary2 %>% 
  filter(cashflow %in% c("Premium","Commission","Policy_Cost") & Transaction_year == 2023 & Recognition_year == 2023) %>% 
  summarise(total_cashflow = sum(value,na.rm = TRUE),
            .by = c(`IFRS Group Code`,Transaction_year, Recognition_year, cashflow))

ifrs_group_csv <- ifrs_group_long %>% 
  filter(run_number %in% c(20,84) & variable_name %in% c("PREMIUM","COMMISSION","POLICY COST")) %>% 
  summarise(total_variable = sum(value,na.rm = TRUE),
            .by = c(ifrs_group_code,variable_name)) %>% 
  mutate(cashflow = case_when(variable_name=="PREMIUM" ~ "Premium",
                              variable_name=="COMMISSION" ~ "Commission",
                              variable_name=="POLICY COST" ~ "Policy_Cost",
                              .default = "NA"))

# cek <- ifrs_group_csv %>% 
#   filter(is.na(cashflow))
compare_ifrs_group <- transaction_ifrs_group %>% 
  left_join(ifrs_group_csv, by = c("IFRS Group Code"="ifrs_group_code","cashflow"="cashflow")) %>% 
  mutate(difference = case_when(cashflow %in% c("Premium","Policy_Cost") ~ total_cashflow+total_variable,
                                .default = total_cashflow-total_variable)
         )

compare_ifrs_group_wide <- compare_ifrs_group %>% 
  select(-variable_name) %>% 
  pivot_wider(names_from = cashflow, values_from = c(total_cashflow,total_variable,difference))
  
compare_ifrs_group_wide <- compare_ifrs_group_wide %>% 
  arrange(desc(abs(difference_Premium)),desc(abs(difference_Commission)),desc(abs(difference_Policy_Cost)))

write.csv(compare_ifrs_group_wide,"compare_ifrs_group.csv", row.names = FALSE)

####sample csv
ifrs_group_sample <- ifrs_group_cf %>% 
  filter(ifrs_group_code=='IC#FIRE#ST#2023#O#IDR#F#IAR#J1#COI#J#OTHERS' & run_number==84)

ifrs_group_sample <- ifrs_group_cf %>% 
  filter(ifrs_group_code=='IC#FIRE#ST#2023#O#IDR#F#PAR#SM#DIR#J#OTHERS' & run_number==84) %>% 
  replace(is.na(.),0)

#IC#FIRE#ST#2023#O#IDR#F#PAR#SM#DIR#J#OTHERS
write.csv(ifrs_group_sample,"ifrs_group_sample.csv", row.names = FALSE)

### Expected claim in IDR
exchange_rate <- read_excel("C:/TMI/IFRS 17/UAT/shared/ExchangeRate.xlsx")
exchange_rate_202312 <- exchange_rate %>% 
  filter(`Valuation Year`==2023&`Valuation Month`==12) %>% 
  select(-`Valuation Year`,-`Valuation Month`)

reserving_premium <- policy_data %>% 
  summarise(Premium = sum(Premium,na.rm = TRUE),
            Commission = sum(AGRICommission,na.rm = TRUE)+sum(OCPORCommission,na.rm = TRUE),
            Policy_Cost = sum(`Policy Cost`,na.rm = TRUE),
            .by = c(Recognition_year,Transaction_year,`Cohort Year`,ReservingClassCode,IsShortTerm,`Currency Code`))
  
reserving_LR <- reserving_premium %>% 
  left_join(exchange_rate_202312, by = c("Currency Code")) %>% 
  left_join(ULR_assumption, by = c("Cohort Year"="Cohort Year","ReservingClassCode"="Reserving Class Code","IsShortTerm"="IsShortTerm")) 
  
reserving_claim <- reserving_LR %>% 
  select(Recognition_year:`Currency Code`,`Exchange Rate`,`Gross ULR`,Premium) %>% 
  arrange(desc(Recognition_year),Transaction_year,ReservingClassCode,IsShortTerm,`Currency Code`)
 


# transaction_calculation2 <- transaction_mapping2 %>% 
#   mutate(Claim = Premium*`Gross ULR`/100,
#          AQE_cf = Premium*`%ACE`/100,
#          PME_cf = Premium*`% PME`/100,
#          CHE_cf = Claim*`% CHE`/100)
# 
# 
# transaction_goc2 <- transaction_calculation2 %>% 
#   summarise(Premium = sum(Premium,na.rm = TRUE),
#             Commission = sum(AGRICommission,na.rm = TRUE)+sum(OCPORCommission,na.rm = TRUE),
#             Policy_Cost = sum(`Policy Cost`,na.rm = TRUE),
#             Claim = sum(Claim,na.rm = TRUE),
#             AQE = sum(AQE_cf,na.rm = TRUE),
#             PME = sum(PME_cf,na.rm = TRUE),
#             CHE = sum(CHE_cf,na.rm = TRUE),
#             .by = c(`IFRS Group Code`,`Method Code`,`Currency Code`,Transaction_year,Recognition_year,Period))
# 
# 
# transaction_summary2 <- transaction_goc2 %>% 
#   pivot_longer(cols = c(Premium:CHE), names_to = "cashflow", values_to = "value")
# 
# transaction_total2 <- transaction_summary2 %>% 
#   summarise(total_transaction = sum(value,na.rm = TRUE),
#             .by = c(Recognition_year,Transaction_year,cashflow)) %>% 
#   arrange(Recognition_year,Transaction_year,cashflow)

write.csv(reserving_claim,"expected_claim.csv", row.names = FALSE)


