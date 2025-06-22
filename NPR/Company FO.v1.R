library(readxl)

data_FO <- "D:/Stanley/Tugas IFRS17/NPR/FO Company List 2023.xlsx"

sheet_name <- excel_sheets(data_FO)
sheet_dn <- sheet_name[grep("- DN", sheet_name)]
sheet_ln <- sheet_name[grep("- LN", sheet_name)]

company_dn <- data.frame()

for (i in sheet_dn){
  #i <- "HB - DN"
  data_temp <- read_excel(data_FO, sheet = i, range = "E16:E200", col_names = FALSE)
  data_temp <- na.omit (data_temp)
  colnames(data_temp) <- "Company"
  
  company_dn <- rbind(company_dn, data_temp)
  
}

company_dn <- unique(company_dn)

company_dn$`L or O` <- "L"

write.csv(company_dn, "D:/Stanley/Tugas IFRS17/NPR/FO Comp 2023 - DN.csv", row.names = FALSE)


company_ln <- data.frame()

for (i in sheet_ln){
  #i <- "HB - DN"
  data_temp <- read_excel(data_FO, sheet = i, range = "E16:F200", col_names = FALSE)
  colnames(data_temp) <- c("Company","Rating")
  data_temp <- data_temp[!is.na(data_temp$Company),]
  
  company_ln <- rbind(company_ln, data_temp)
  
}

company_ln$`L or O` <- "O"

company_ln <- unique(company_ln)

write.csv(company_ln, "D:/Stanley/Tugas IFRS17/NPR/FO Comp 2023 - LN.csv", row.names = FALSE)
