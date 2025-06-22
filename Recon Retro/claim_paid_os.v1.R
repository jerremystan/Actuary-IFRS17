library(dplyr)

claim_os <- read.csv("Input/Retro Claim OS v2.csv")
nama_os <- read.delim("clipboard", header = FALSE)
nama_os <- unlist(nama_os[1,])

claim_paid <- read.csv("Input/Retro Claim Paid v2.csv")
nama_paid <- read.delim("clipboard", header = FALSE)
nama_paid <- unlist(nama_paid[1,])

claim_paid <- claim_paid[,-1]
colnames(claim_paid) <- nama_paid
