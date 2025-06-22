library(tidyverse)
library(openxlsx)
library(readxl)

ifrs_group_cf_1 <- read_csv("ifrs_group_cashflow_all_runs_2023_202312_1.csv")
ifrs_group_cf_2 <- read_csv("ifrs_group_cashflow_all_runs_2023_202312_2.csv")
ifrs_group_cf_3 <- read_csv("ifrs_group_cashflow_all_runs_2023_202312_3.csv")
ifrs_group_cf_4 <- read_csv("ifrs_group_cashflow_all_runs_2023_202312_4.csv")
ifrs_group_cf_5 <- read_csv("ifrs_group_cashflow_all_runs_2023_202312_5.csv")
ifrs_group_cf_6 <- read_csv("ifrs_group_cashflow_all_runs_2023_202312_6.csv")
ifrs_group_cf <- rbind(ifrs_group_cf_1,ifrs_group_cf_2,ifrs_group_cf_3,ifrs_group_cf_4,ifrs_group_cf_5,ifrs_group_cf_6)

##expense_assumptions <- read_excel("C:/TMI/IFRS 17/UAT/shared/ActualExpense.v2.xlsx", sheet = "2023")
#expense_assumptions2 <- read_excel("C:/TMI/IFRS 17/UAT/shared/Expense Study.2023.v.3.xlsx", sheet = "Encoding Unit Cost Driver")

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
            .by = c(run_number,variable_name,custom_dim_1, transaction_currency_code))

total_run <- total_run %>% 
  arrange(run_number,variable_name,custom_dim_1)

total_cf <- total_run %>% 
  filter(run_number %in% c(20,40,84,90)) %>%
  mutate(period_cf = case_when(run_number %in% c(20,84) ~ "Current",
                               run_number %in% c(40,90) ~ "Future",
                               .default = "NA")
  ) %>% 
  summarise(total_cashflow = sum(total_cashflow, na.rm = TRUE),
            .by = c(variable_name,period_cf, transaction_currency_code))

write.csv(total_run,"total_run.csv", row.names = FALSE)
write.csv(total_cf,"total_cashflow.csv", row.names = FALSE)
