start_time <- Sys.time()
library("readxl")
library(tidyverse)
setwd(getwd())
#
#user can change these values
FileName <- "Return and Correlation Data.xlsx"
NumberofYears <- 20
NumberofSimulations <- 10000

PercentileStart <- 5
PercentileEnd <- 95
PercentileGap <- 5

ProbabilityStart <- 3
ProbabilityEnd <- 11
ProbabilityGap <- 0.25
#
#Field List has list of asset classes for each bank to look up for returns and correlations
TemplateData <- read_excel(FileName, sheet = "FieldList")
#Each Bank has its own tab, get tab list
Tabs <- excel_sheets(FileName)

#Set default return matrix for which the avg is calculated
ReturnMatrix <- TemplateData[,1:4]
#clean it up
ReturnMatrix[,3:4] <- as.double(0)
colnames(ReturnMatrix)[3] <- "Annual Volatility (??n)"
colnames(ReturnMatrix)[4] <- "Annual Volatility (Expected returns)"

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
for (i in seq(PercentileStart,PercentileEnd,PercentileGap)){
    textlabel <- c(paste(as.character(i),"th percentile"))
    PercentileTab <- rbind(PercentileTab,textlabel)
}

#Initialize Probability Tab based on start,end and gap
ProbabilityTab <- c('Return')
for (i in seq(ProbabilityStart,ProbabilityEnd,ProbabilityGap)){
  textlabel <- c(paste(as.character(i),"%"))
  ProbabilityTab <- rbind(ProbabilityTab,textlabel)
}
#
#Go through each tab. Each tab is each bank
for (i in 2:length(Tabs)){
  #For each bank, get field list and tab data
  FieldList <- TemplateData[,grep(Tabs[i],colnames(TemplateData))]
  TabData <- read_excel(FileName, sheet = Tabs[i])
  
  #initialize the return matrix and correlation matrix for each bank
  TempReturnMatrix <- ReturnMatrix
  TempReturnMatrix[,1] <- FieldList
  TempCorrelationMatrix <- CorrelationMatrix
  
  for (j in 1:nrow(FieldList)){
      #Get row index for asset class
      RowIndex <- which(TabData[,1] == as.character(FieldList[j,1]))
      #Volatility
      TempReturnMatrix[j,3] <- as.double(TabData[RowIndex,4])
      #Returns
      TempReturnMatrix[j,4] <- as.double(TabData[RowIndex,3])
      
      #Add individual return matrix for each bank to total
      ReturnMatrix[j,3] <- ReturnMatrix[j,3] + TempReturnMatrix[j,3]
      ReturnMatrix[j,4] <- ReturnMatrix[j,4] + TempReturnMatrix[j,4]
      
      #Because it's a square matrix, you can iterate through nrow(FielList) for the column part
      for (k in 1:nrow(FieldList)){
        ColIndex <- which(colnames(TabData) == as.character(FieldList[k,1]))
        TempCorrelationMatrix[j,k] <- as.double(TabData[RowIndex,ColIndex])
        CorrelationMatrix[j,k] <- CorrelationMatrix[j,k] + TempCorrelationMatrix[j,k]
      }
  }
  
  #Do this for each bank and for each component of the output - summary, percentile, probabilities
  #Convert to Matrix for later operations
  WeightVolatility <- as.matrix(TempReturnMatrix[,2]*TempReturnMatrix[,3])
  AnnualMeanReturn <- sum(TempReturnMatrix[,2]*TempReturnMatrix[,4])
  Matrix1 <- TempCorrelationMatrix %*% WeightVolatility
  AnnualVariance <- t(WeightVolatility) %*% Matrix1
  AnnualVolatility <- sqrt(AnnualVariance)
  
  georeturn <- function(return) {
    georeturn <- prod(1+return)^(1/length(return)) - 1
    return(georeturn)
  }
  
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
  for (p in seq(PercentileStart,PercentileEnd,PercentileGap)){
    ReturnPerc <- quantile(georeturnsim,p/100)
    BankPercentile <- rbind(BankPercentile,ReturnPerc)
  }
  PercentileTab <- cbind(PercentileTab,BankPercentile)
  
  #Probability Tab
  BankProbability <- c(Tabs[i])
  for (p in seq(ProbabilityStart,ProbabilityEnd,ProbabilityGap)){
    ReturnPercent <- sum(georeturnsim >= p/100) / NumberofSimulations
    BankProbability <- rbind(BankProbability,ReturnPercent)
  }
  ProbabilityTab <- cbind(ProbabilityTab,BankProbability)
}
#
#Do the average for all banks
#Divide to get average, first Tab is FieldList so we exclude it
ReturnMatrix[,3:4] <- ReturnMatrix[,3:4] / (length(Tabs) - 1)
CorrelationMatrix <- CorrelationMatrix / (length(Tabs) - 1)

#Convert to Matrix for later operations
WeightVolatility <- as.matrix(ReturnMatrix[,2]*ReturnMatrix[,3])
AnnualMeanReturn <- sum(ReturnMatrix[,2]*ReturnMatrix[,4])
Matrix1 <- CorrelationMatrix %*% WeightVolatility
AnnualVariance <- t(WeightVolatility) %*% Matrix1
AnnualVolatility <- sqrt(AnnualVariance)

georeturn <- function(return) {
  georeturn <- prod(1+return)^(1/length(return)) - 1
  return(georeturn)
}

randomreturn <- rnorm(NumberofYears, mean = AnnualMeanReturn, sd = AnnualVolatility)
georeturnsim <- replicate(n = NumberofSimulations,georeturn(rnorm(NumberofYears, mean = AnnualMeanReturn, sd = AnnualVolatility)))

#Summary Tab
BankSummaryData <- c('Average',mean(georeturnsim),
                     median(georeturnsim),
                     sd(georeturnsim),
                     AnnualMeanReturn,
                     AnnualVolatility)
SummaryTab <- rbind(SummaryTab,BankSummaryData)
SummaryTab <- t(SummaryTab)

#Percentile Tab
BankPercentile <- c('Average')
for (p in seq(PercentileStart,PercentileEnd,PercentileGap)){
  ReturnPercentile <- quantile(georeturnsim,p/100)
  BankPercentile <- rbind(BankPercentile,ReturnPercentile)
}
PercentileTab <- cbind(PercentileTab,BankPercentile)

#Probability Tab
BankProbability <- c('Average')
for (p in seq(ProbabilityStart,ProbabilityEnd,ProbabilityGap)){
  ReturnPercent <- sum(georeturnsim >= p/100) / NumberofSimulations
  BankProbability <- rbind(BankProbability,ReturnPercent)
}
ProbabilityTab <- cbind(ProbabilityTab,BankProbability)

#Write to csv file
write_excel_csv(as.data.frame(SummaryTab), 'Summary.csv',col_names = FALSE)
write_excel_csv(as.data.frame(PercentileTab), 'Percentiles.csv',col_names = FALSE)
write_excel_csv(as.data.frame(ProbabilityTab), 'Return Probabilities.csv',col_names = FALSE)

#Cal run time
end_time <- Sys.time()
print(end_time - start_time)
