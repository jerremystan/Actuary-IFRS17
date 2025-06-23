library(readxl)
library(writexl)
library(readr)

lob_filter <- "ENGINEERING"

Amortization <- read_csv("input dataset/202312/Amortization Insurance Contract.2023.202312.v.4.csv")
#Amortization <- read_csv("input dataset/202312/Amortization Non XL Facultative.2023.202312.v.2.xlsx")
# View(Amortization_Insurance_Contract_2023_v_3)

filter_df <- Amortization[Amortization$`Reserving Class` == lob_filter,]

write_xlsx(filter_df,paste0(lob_filter,"-IC-AMRT.xlsx"))
