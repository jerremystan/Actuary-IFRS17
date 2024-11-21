library(readxl)
library(dplyr)
library(tidyr)
library(lubridate)
library(purrr)
library(openxlsx)
library(stringr)

mapping_RC <- read_excel("mappingRc.xlsx")

summary_opex <- read_excel("OpexAmount.2023.v2.xlsx")

colnames(summary_opex)[5] <- "Reserving Class"

join_opex <- left_join(summary_opex, mapping_RC, by = "Reserving Class")

#join_opex$`Valuation Date` <- ymd("2023-12-31")

eom_date <- seq(ceiling_date(as.Date("2023-01-01"), unit = "month"),
                ceiling_date(as.Date("2023-12-01"), unit = "month"), by = "month") - 1
eom_date <- as.character.Date(eom_date)

group_opex <- join_opex[,c(1,2,4,11,6,7,8,9,10)]

group_opex <- replace(group_opex, is.na(group_opex), 0)

colnames(group_opex)[4] <- "Reserving Class Code"

opex <- group_opex %>% group_by(Book, `Valuation Date`, `Category Code`, `Reserving Class Code`, `Distribution Channel`, `Profit Center`) %>% summarise(ACE = sum(ACE), PME = sum(PME), CHE = sum(CHE))

opex <- replace(opex, is.na(opex),0)

opex <- opex[order(opex$Book,opex$`Valuation Date`,opex$`Category Code`,opex$`Reserving Class Code`,opex$`Distribution Channel`,opex$`Profit Center`),]

df_opex <- data.frame()

'for (i in eom_date){
  
  new_df <- opex
  new_df$`Valuation Date` <- as.character(i)
  
  df_opex <- rbind(df_opex, new_df)
  
}'

df_opex <- opex
df_opex$`Valuation Date` <- "2022-12-31"

colnames(df_opex) <- c("Book", "Valuation Date", "Category Code", "Reserving Class Code", "Distribution Channel Name", "Profit Center Description", "ACE", "PME", "CHE")

write.csv(df_opex,"SummaryOpex.csv", row.names = FALSE)

wb <- createWorkbook()

stack_function <- function(book, sheet){
  
  final_result <- df_opex[df_opex$Book == book,]
  
  addWorksheet(wb, book)
  writeData(wb, sheet, x = final_result, startRow = 1, startCol = 1)
}

Book_TMI <- stack_function("TMI", 1)
Book_PT <- stack_function("Parolamas", 2)

saveWorkbook(wb, "SummaryOpex.2022.v2.xlsx")

