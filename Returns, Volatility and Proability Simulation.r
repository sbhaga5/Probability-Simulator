start_time <- Sys.time()
library("readxl")
library(tidyverse)
setwd(getwd())
#
#user can change these values
FileName <- "Return and Correlation Data 2020 TRS.xlsx"
NumberofYears <- 20
NumberofSimulations <- 10000

Percentiles <- c(5,10,25,50,75,90,95)
Probabilities <- c(9,8,7.5,7,6.5,6,5)
FundARR <- 0.075
#
#Field List has list of asset classes for each bank to look up for returns and correlations
TemplateData <- read_excel(FileName, sheet = "FieldList")
#Each Bank has its own tab, get tab list
Tabs <- excel_sheets(FileName)

#Set default return matrix for which the avg is calculated
ReturnMatrix <- TemplateData[,1:4]
#clean it up
ReturnMatrix[,3:4] <- as.double(0)

#Set default correlation matrix for average
CorrelationMatrix <- matrix(0, nrow = nrow(TemplateData), ncol = nrow(TemplateData))
#
#Initialize Tabs
#initialize the summary tab
SummaryTab <- c('Bank Name','20-year geometric avg return (mean)',
                '20-year geometric avg return (median)',
                'Standard deviation of 20-year avg returns',
                'Annual arithmetic avg return',
                'Standard deviation of annual returns')

#Initialize Percentile Tab based on start,end and gap
PercentileTab <- c('Percentile')
for (i in 1:length(Percentiles)){
  textlabel <- c(paste(as.character(Percentiles[i]),"th percentile"))
  PercentileTab <- rbind(PercentileTab,textlabel)
}

#Initialize Probability Tab based on start,end and gap
ProbabilityTab <- c('Return')
for (i in 1:length(Probabilities)){
  textlabel <- c(paste(as.character(Probabilities[i]),"%"))
  ProbabilityTab <- rbind(ProbabilityTab,textlabel)
}
#
#We need to calculate Arithmetic return from assumed geometric return
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
#
#Go through each tab. Each tab is each bank
#First Tab is FieldList, so start at 2
for (i in 2:(length(Tabs))){
  TabData <- read_excel(FileName, sheet = Tabs[i])
  
  #For each bank, get field list and tab data
  #For fund details and Historical, dont do the correlation matrix calculation
  if((Tabs[i] != 'Fund Details') && (Tabs[i] != 'Historical')){
    #initialize the return matrix and get fieldlist for each bank
    FieldList <- TemplateData[,grep(Tabs[i],colnames(TemplateData))]
    ReturnMatrix[,1] <- FieldList
    
    for (j in 1:nrow(FieldList)){
      #Get row index for asset class
      RowIndex <- which(TabData[,1] == as.character(FieldList[j,1]))
      #Volatility
      ReturnMatrix[j,3] <- as.double(TabData[RowIndex,4])
      #Returns
      ReturnMatrix[j,4] <- as.double(TabData[RowIndex,3])
      
      #Because it's a square matrix, you can iterate through nrow(FielList) for the column part
      for (k in 1:nrow(FieldList)){
        ColIndex <- which(colnames(TabData) == as.character(FieldList[k,1]))
        CorrelationMatrix[j,k] <- as.double(TabData[RowIndex,ColIndex])
      }
    }
  }
  
  #Do this for each bank and for each component of the output - summary, percentile, probabilities
  #For Fund details, the volatility for the whole is given, no need for calc
  if(Tabs[i] == 'Fund Details'){
    AnnualVolatility <- HorizonVol
    AnnualMeanReturn <- GetArithReturn(HorizonVol, FundARR)
  } else if(Tabs[i] == 'Historical'){
    AnnualVolatility <- sd(as.matrix(TabData[,2]))
    AnnualMeanReturn <- mean(as.matrix(TabData[,2]))
  } else {
    AnnualMeanReturn <- sum(ReturnMatrix[,2]*ReturnMatrix[,4])
    #Convert to Matrix for later operations
    WeightVolatility <- as.matrix(ReturnMatrix[,2]*ReturnMatrix[,3])
    Matrix1 <- CorrelationMatrix %*% WeightVolatility
    AnnualVariance <- t(WeightVolatility) %*% Matrix1
    AnnualVolatility <- sqrt(AnnualVariance)
    
    if(Tabs[i] == 'Horizon10'){
      HorizonVol <- AnnualVolatility
    }
  }
  
  #set.seed is for fixing the simulation to avoid too much variance
  set.seed(1234)
  georeturn <- function(return) {
    georeturn <- prod(1+return)^(1/length(return)) - 1
    return(georeturn)
  }
  
  #simulate the returns
  randomreturn <- rnorm(NumberofYears, mean = AnnualMeanReturn, sd = AnnualVolatility)
  georeturnsim <- replicate(n = NumberofSimulations,georeturn(rnorm(NumberofYears, mean = AnnualMeanReturn, sd = AnnualVolatility)))
  
  #Summary Tab
  BankSummaryData <- c(Tabs[i],mean(georeturnsim),
                       median(georeturnsim),
                       sd(georeturnsim),
                       AnnualMeanReturn,
                       AnnualVolatility)
  SummaryTab <- rbind(SummaryTab,BankSummaryData)
  
  #Percentile Tab
  BankPercentile <- c(Tabs[i])
  for (p in 1:length(Percentiles)){
    ReturnPerc <- quantile(georeturnsim, (Percentiles[p]/100))
    BankPercentile <- rbind(BankPercentile,ReturnPerc)
  }
  PercentileTab <- cbind(PercentileTab,BankPercentile)
  
  #Probability Tab
  BankProbability <- c(Tabs[i])
  for (p in 1:length(Probabilities)){
    ReturnPercent <- sum(georeturnsim >= (Probabilities[p]/100)) / NumberofSimulations
    BankProbability <- rbind(BankProbability,ReturnPercent)
  }
  ProbabilityTab <- cbind(ProbabilityTab,BankProbability)
}
SummaryTab <- t(SummaryTab)

#Write to csv file
write_excel_csv(as.data.frame(SummaryTab), 'Summary.csv',col_names = FALSE)
write_excel_csv(as.data.frame(PercentileTab), 'Percentiles.csv',col_names = FALSE)
write_excel_csv(as.data.frame(ProbabilityTab), 'Return Probabilities.csv',col_names = FALSE)

#Cal run time
end_time <- Sys.time()
print(end_time - start_time)
