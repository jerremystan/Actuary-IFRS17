library(dplyr)
library(readxl)
library(stringr)

currency_list <- unlist(read.csv("Ubah List Currency disini!.csv"), recursive = FALSE)

currency_list <- as.list(currency_list)

#list_nama <- list.files("D:/Stanley/Tugas IFRS17/Exchange Rate/Exchange Rate/Input BI")

#list_nama <- as.list(list_nama)

final_data <- data.frame()

for (i in currency_list){

nama_ccy <- i
i <- ifelse(i == "RMB", "CNY", i)
i <- ifelse(i == "WON", "KRW", i)
  
data <- read_excel(paste0("D:/Stanley/Tugas IFRS17/Exchange Rate/Exchange Rate/Input BI/Kurs Transaksi ",i,"  .xlsx"))

colnames(data) <- c("No","Nilai","KursJual","KursBeli","Tanggal")

#menghilangkan 12 AM

data <- data %>% mutate(Tanggal = gsub(" 12:00:00 AM","",Tanggal))

data$Tanggal <- as.Date(data$Tanggal, format = "%m/%d/%Y")

clean_data <- data[-1:-2,-1]

clean_data <- clean_data %>% mutate(Month = format(Tanggal, "%m"), Year = format(Tanggal, "%Y"))

filtered_data <- clean_data %>% group_by(Month, Year) %>% filter(Tanggal == max(Tanggal))

filtered_data$KursTengah <- (as.numeric(filtered_data$KursJual)+as.numeric(filtered_data$KursBeli))/2

if_A <- data.frame(Month = c("12","11", "10", "09", "08", "07", "06", "05", "04", "03", "02", "01"),
                   NewMonth = c("01","12","11", "10", "09", "08", "07", "06", "05", "04", "03", "02"))

filtered_data <- left_join(filtered_data, if_A, by = "Month")

filtered_data$NewYear <- ifelse(filtered_data$Month == "12", as.numeric(filtered_data$Year)+1, as.numeric(filtered_data$Year))

filtered_data <- filtered_data[,-5:-6]

filtered_data$Currency <- rep(nama_ccy, nrow(filtered_data)) #ganti AUD dng i

filtered_data <- filtered_data %>% select(Currency, Nilai, KursJual, KursBeli, Tanggal, KursTengah, NewYear, NewMonth)

colnames(filtered_data) <- c("Mata Uang", "Nilai", "Kurs Jual", "Kurs Beli", "Tanggal", "Kurs Tengah", "Valuation Year", "Valuation Month")

final_data <- rbind(final_data, filtered_data)

}

write.csv(final_data,"Output BI/Exchange Rate.csv", row.names = FALSE)

