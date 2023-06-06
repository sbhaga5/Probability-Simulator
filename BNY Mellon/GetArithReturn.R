GeomArithReturn <- function(ArithReturn, Variance){
  GeomArithReturn <- ArithReturn - ((Variance*Variance)/(2*(1+ArithReturn)))
}

#use a bisection method to solve for arithmetic return
GetArithReturn <- function(Variance, Target){
  Signal <- 'Add'
  Increment <- 0.1
  ArithReturn <- 0.1
  Check <- 'Fail'
  Difference <- Target - GeomArithReturn(ArithReturn,Variance)
  
  while (Check != 'Pass'){
    if((Difference < 0) && (Signal == 'Add')){
      Increment <- -0.5*Increment
      Signal <- 'Subtract'
    } else if ((Difference > 0) && (Signal == 'Subtract')) {
      Increment <- -0.5*Increment
      Signal <- 'Add'
    }
    
    ArithReturn <- ArithReturn + Increment
    Difference <- Target - GeomArithReturn(ArithReturn,Variance)
    if(abs(Difference) < 1e-6){
      Check <-'Pass'
    }
  }
  
  return(ArithReturn)
}

library("readxl")
library(tidyverse)
setwd(getwd())
TemplateData <- read_excel('Book1.xlsx')

for(i in 1:nrow(TemplateData)){
  Variance <- TemplateData[i,4]
  Target <- TemplateData[i,2]
  TemplateData[i,3] <- GetArithReturn(Variance,Target)
}
write_excel_csv(as.data.frame(TemplateData), 'Book2.csv',col_names = FALSE)
