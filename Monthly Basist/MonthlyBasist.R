library(readxl)
library(lubridate)
library(dplyr)
library(writexl)
library(openxlsx)

list_files <- list.files("Input")

ULR <- read_excel(paste0("Input/", list_files[grepl("ULR",list_files)]))
LDF <- read_excel(paste0("Input/", list_files[grepl("Development",list_files)]))
NPR <- read_excel(paste0("Input/", list_files[grepl("NonPer",list_files)]))
RM <- read_excel(paste0("Input/", list_files[grepl("RiskMargin",list_files)]))
IELR <- read_excel(paste0("Input/", list_files[grepl("Initial",list_files)]))

year_function <- function(df,year,beginM,lastM){
  
  ifelse(lastM == 12, lastM <- 11, lastM <- lastM)
  
  df_temp <- df
  df_filter <- df_temp[df_temp$`Valuation Date` == paste0((year-1),"-12-31"),]
  #monthly <- sprintf("%02d", beginM:lastM)
  new_df <- df
  for (i in beginM:lastM){
    new_row <- df_filter
    month <- ceiling_date(make_date(year,i,1), "month") - days(1)
    new_row$`Valuation Date` <- as.character(month)
    
    new_df <- rbind(new_df, new_row)
  }
  new_df <- new_df[order(new_df$`Valuation Date`),]
  
  return(new_df)
  
}

quarter_function <- function(df,year,beginM,lastM){
  
  ifelse(lastM %in% c(3,6,9,12), lastM <- lastM - 1, lastM)
  
  df_temp <- df
  range <- beginM:lastM
  range <- range[!(range %in% c(3,6,9,12))]
  new_df <- df
  for (i in range){
    
    if (i %in% c(1,2)) {
      new_row <- df %>% filter(`Valuation Date` == (ceiling_date(make_date((year-1), 12, 1), "month") - days(1)))
    } else if (i %in% c(4,5)) {
      new_row <- df %>% filter(`Valuation Date` == (ceiling_date(make_date(year, 3, 1), "month") - days(1)))
    } else if (i %in% c(7,8)) {
      new_row <- df %>% filter(`Valuation Date` == (ceiling_date(make_date(year, 6, 1), "month") - days(1)))
    } else if (i %in% c(10,11)) {
      new_row <- df %>% filter(`Valuation Date` == (ceiling_date(make_date(year, 9, 1), "month") - days(1)))
    } else {
      new_row <- data.frame()
    }
    
    month <- ceiling_date(make_date(year, i, 1), "month") - days(1)
    
    new_row$`Valuation Date` <- as.character(month)
    
    new_df <- rbind(new_df,new_row)
  }
  
  new_df <- new_df[order(new_df$`Valuation Date`),]
  return(new_df)
  
}

numeric_function <- function(df) {
  df_temp <- df
  numeric_columns <- names(df_temp)[sapply(df_temp, is.numeric)]
  
  df_temp[numeric_columns] <- lapply(df_temp[numeric_columns], function(x) round(x, 22))
  
  return(df_temp)
}


#ULR
ULR_final <- quarter_function(ULR,2024,1,11)
ULR_final <- quarter_function(ULR_final,2025,1,2)

#LDF
LDF_final <- quarter_function(LDF,2024,1,11)
LDF_final <- quarter_function(LDF_final,2025,1,2)

#NPR
NPR_final <- year_function(NPR,2024,1,11)
NPR_final <- year_function(NPR_final,2025,1,2)

#RM
RM_final <- year_function(RM,2024,1,11)
RM_final <- year_function(RM_final,2025,1,2)

#IELR
IELR_final <- quarter_function(IELR,2024,1,11)
IELR_final <- quarter_function(IELR_final,2025,1,2)

write.xlsx(ULR_final, paste0("Output/", gsub(".xlsx","-new22dec.xlsx",list_files[grepl("ULR",list_files)])), sheetName = "Sheet1", rowNames = FALSE, overwrite = TRUE)
write.xlsx(LDF_final, paste0("Output/", gsub(".xlsx","-new22dec.xlsx",list_files[grepl("Development",list_files)])), sheetName = "LDF", rowNames = FALSE, overwrite = TRUE)
write.xlsx(NPR_final, paste0("Output/", gsub(".xlsx","-new22dec.xlsx",list_files[grepl("NonPerform",list_files)])), sheetName = "Sheet1", rowNames = FALSE, overwrite = TRUE)
write.xlsx(RM_final, paste0("Output/", gsub(".xlsx","-new.xlsx",list_files[grepl("RiskMargin",list_files)])), sheetName = "Actual", rowNames = FALSE, overwrite = TRUE)
write.xlsx(IELR_final, paste0("Output/", gsub(".xlsx","-new.xlsx",list_files[grepl("Initial",list_files)])), sheetName = "IELR", rowNames = FALSE, overwrite = TRUE)

