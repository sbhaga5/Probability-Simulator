library(tidyverse)
setwd(getwd())
Data <- read_csv('2022 Data.csv', col_names = FALSE)

for (i in 1:nrow(Data)){
  for (j in 1:ncol(Data)){
    if(is.na(Data[i,j])){
      Data[i,j] <- Data[j,i]
    }
  }
}

write_excel_csv(as.data.frame(Data), 'Data Modified.csv',col_names = FALSE)
