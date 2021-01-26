library("readxl")
library(tidyverse)
setwd(getwd())
FileName <- 'Model Inputs.xlsx'
#
#
model_inputs <- read_excel(FileName, sheet = 'Main')
MeritIncreases <- model_inputs[,5]
payroll_growth <- as.double(model_inputs[which(model_inputs[,1] == 'Payroll Growth Rate (Basic)'),2])
StartingSalary <- as.double(model_inputs[which(model_inputs[,1] == 'Starting Salary'),2])
Vesting <- as.double(model_inputs[which(model_inputs[,1] == 'Vesting (years)'),2])
HiringAge <- as.double(model_inputs[which(model_inputs[,1] == 'Hiring Age'),2])
YOS <- model_inputs[,4]
ER_Contrib <- as.double(model_inputs[which(model_inputs[,1] == 'ER Contribution Rate'),2])
EE_Contrib <- as.double(model_inputs[which(model_inputs[,1] == 'EE Contribution Rate'),2])
ARR <- as.double(model_inputs[which(model_inputs[,1] == 'ARR/DR'),2])
Interest <- as.double(model_inputs[which(model_inputs[,1] == 'Credited Interest'),2])
COLA <- as.double(model_inputs[which(model_inputs[,1] == 'COLA'),2])
COLACompound <- as.double(model_inputs[which(model_inputs[,1] == 'COLA Compounds (1=Yes)'),2])
#
#
#Linear Interpolation Function. Use for later
LinearInterpolation <- function(Data){
  StartValue <- Data[1,1]
  StartIndex <- 2
  count <- 2
  while(count <= nrow(Data)){
    if(!is.na(Data[count,1])){
      EndValue <- Data[count,1]
      EndIndex <- count - 1
      
      for(i in StartIndex:EndIndex){
        Data[i,1] <- Data[i-1,1] - (StartValue - EndValue)/(EndIndex - StartIndex + 2)
      }
      
      StartValue <- Data[count,1]
      StartIndex <- count + 1
    } else if((count == nrow(Data))) {
      EndIndex <- count
      for(i in StartIndex:EndIndex){
        Data[i,1] <- Data[i-1,1] - (StartValue - EndValue)/(EndIndex - StartIndex + 2)
      }
    }
    count <- count + 1
  }
  
  return(Data)
}
#
#
#Salary increases and other
MeritIncreases <- LinearInterpolation(MeritIncreases)
Salary <- MeritIncreases
colnames(Salary) <- 'Salary Growth'

Salary[1,1] <- StartingSalary
FinalAvgSalary <- MeritIncreases
colnames(FinalAvgSalary) <- 'Final Average Salary Growth'
for(i in 1:nrow(MeritIncreases)){
  if(i > 1){
    Salary[i,1] <- Salary[i-1,1]*(1 + (MeritIncreases[i-1,1] + payroll_growth)) 
  }
  
  FinalAvgSalary[i,1] <- 0
  if(YOS[i,1] >= Vesting){
    FinalAvgSalary[i,1] <- sum(Salary[(i-Vesting):(i-1),1])/Vesting
  }
}
#
#
#Employee Contribution
Employee_Contrib <- Salary*EE_Contrib
colnames(Employee_Contrib) <- 'Employee Contribution'
Employer_Contrib <- Salary*ER_Contrib
colnames(Employer_Contrib) <- 'Employer Contribution'

Employee_Contrib_CI <- Salary*EE_Contrib
colnames(Employee_Contrib) <- 'Employee Contribution (with Compounded Interest)'
PV_Employee_Contrib_CI <- Employee_Contrib_CI
colnames(PV_Employee_Contrib_CI) <- 'PV of Employee Contribution (with Compounded Interest)'

for(i in 1:nrow(PV_Employee_Contrib_CI)){
  if((YOS[i,1] > 0) && (i > 1)){
    Employee_Contrib_CI[i,1] <- (Employee_Contrib_CI[i-1,1]*(1+Interest) + Employee_Contrib[i-1,1])
    PV_Employee_Contrib_CI[i,1] <- Employee_Contrib_CI[i,1]/((1+ARR)^(as.double(YOS[i,1])))
  } else {
    Employee_Contrib_CI[i,1] <- 0
    PV_Employee_Contrib_CI[i,1] <- 0
  }
}
#
#
#Mortality Rates
#SurvivalModel <- model_inputs[which(model_inputs[,1] == 'Survival Model'),2]
MaleMortality <- read_excel(FileName, sheet = 'MP-2017_Male')
FemaleMortality <- read_excel(FileName, sheet = 'MP-2017_Female')
SurvivalRates <- read_excel(FileName, sheet = 'Mortality Rates')

Age <- SurvivalRates[,1]
Pub2010NonDisMaleTeacher <- SurvivalRates[,2]
Pub2010DisMaleTeacher <- SurvivalRates[,3]
Pub2010MaleTeacher <- SurvivalRates[,4]
Pub2010NonDisFemaleTeacher <- SurvivalRates[,5]
Pub2010DisFemaleTeacher <- SurvivalRates[,6]
Pub2010FemaleTeacher <- SurvivalRates[,7]

RP2014Male <- SurvivalRates[,8]
RP2014HealthyMale <- SurvivalRates[,9]
RP2014DisMale <- SurvivalRates[,10]
RP2014Female <- SurvivalRates[,11]
RP2014HealthyFemale <- SurvivalRates[,12]
RP2014DisFemale <- SurvivalRates[,13]

Pub2010NonDisMaleTeacher <- LinearInterpolation(Pub2010NonDisMaleTeacher)
Pub2010DisMaleTeacher <- LinearInterpolation(Pub2010DisMaleTeacher)
Pub2010MaleTeacher <- LinearInterpolation(Pub2010MaleTeacher)
Pub2010NonDisFemaleTeacher <- LinearInterpolation(Pub2010NonDisFemaleTeacher)
Pub2010DisFemaleTeacher <- LinearInterpolation(Pub2010DisFemaleTeacher)
Pub2010FemaleTeacher <- LinearInterpolation(Pub2010FemaleTeacher)

YearCol <- (2129 - 2009 + 1)
AgeRow <- (120 - 25 + 1)
SurvivalMaleMatrix <- matrix(0, nrow = AgeRow, ncol = YearCol)
SurvivalFemaleMatrix <- matrix(0, nrow = AgeRow, ncol = YearCol)

for(i in 25:125){
  RowCount <- i - 25 + 1
  
  if(i <= 80){
    SurvivalMaleMatrix[RowCount,2] <- as.double(Pub2010MaleTeacher[which(SurvivalRates[,1] == i),1])
    SurvivalFemaleMatrix[RowCount,2] <- as.double(Pub2010FemaleTeacher[which(SurvivalRates[,1] == i),1])
  } else {
    SurvivalMaleMatrix[RowCount,2] <- as.double(RP2014HealthyMale[which(SurvivalRates[,1] == i),1])
    SurvivalFemaleMatrix[RowCount,2] <- as.double(RP2014HealthyFemale[which(SurvivalRates[,1] == i),1])
  }
  
  for(j in 2011:2129){
    ColCount <- j - 2009 + 1
    if(j <= 2033){
      YearRef <- j
    } else {
      YearRef <- 2033
    }
    
    RowIndex <- which(MaleMortality[,1] == i)
    ColIndex <- which(colnames(MaleMortality) == YearRef)
    SurvivalMaleMatrix[RowCount,ColCount] <- SurvivalMaleMatrix[RowCount,ColCount-1]*(1 - as.double(MaleMortality[RowIndex,ColIndex]))
    
    RowIndex <- which(FemaleMortality[,1] == i)
    ColIndex <- which(colnames(FemaleMortality) == YearRef)
    SurvivalFemaleMatrix[RowCount,ColCount] <- SurvivalFemaleMatrix[RowCount,ColCount-1]*(1 - as.double(FemaleMortality[RowIndex,ColIndex]))
    
    if(j == 2020){
      SurvivalMaleMatrix[RowCount,1] <- SurvivalMaleMatrix[RowCount,12]
      SurvivalFemaleMatrix[RowCount,1] <- SurvivalFemaleMatrix[RowCount,12]
    }
  }
}
#
#
#RP2014 and Adjusted Mortality Rates
#RP Rates and Adjusted Mortality
RPRatesMale <- RP2014Male
RPRatesFemale <- RP2014Female
RPRatesAvg <- RPRatesMale

AdjustedMale <- RP2014Male
AdjustedFemale <- RP2014Female
AdjustedAvg <- RPRatesMale

for(i in 1:nrow(RPRatesMale)){
  if(is.na(RP2014Male[i,1])){
    RPRatesMale[i,1] <- RP2014HealthyMale[i,1]
  } else {
    RPRatesMale[i,1] <- RP2014Male[i,1]
  }
  
  if(is.na(RP2014Female[i,1])){
    RPRatesFemale[i,1] <- RP2014HealthyFemale[i,1]
  } else {
    RPRatesFemale[i,1] <- RP2014Female[i,1]
  }
  
  RPRatesAvg[i,1] <- (RPRatesMale[i,1] + RPRatesFemale[i,1]) /2
  
  if(Age[i,1] < HiringAge){
    AdjustedMale[i,1] <- SurvivalMaleMatrix[i,11]
    AdjustedFemale[i,1] <- SurvivalFemaleMatrix[i,11]
  } else {
    AdjustedMale[i,1] <- SurvivalMaleMatrix[i,11+i-1]
    AdjustedFemale[i,1] <- SurvivalFemaleMatrix[i,11+i-1]
    AdjustedAvg[i,1] <- (AdjustedMale[i,1] + AdjustedFemale[i,1]) / 2
  }
}
#
#
#Survival Probability and Annuity Factor
SurvivalRates <- AdjustedAvg

ProbNextYear <- SurvivalRates
colnames(ProbNextYear) <- 'Probability of Survival to Next Year'
ProbNextYear[1,1] <- 1

DiscProbNextYear <- ProbNextYear
colnames(DiscProbNextYear) <- 'Discount Probability of Survival to Next Year'

LifeExpectancy <- ProbNextYear
colnames(LifeExpectancy) <- 'Life Expectancy'

for(i in 1:nrow(ProbNextYear)){
  if(i == 1){
    ProbNextYear[i,1] <- (1-SurvivalRates[i,1]) * 1
  } else {
    ProbNextYear[i,1] <- (1-SurvivalRates[i-1,1]) * ProbNextYear[i-1,1]
  }
  DiscProbNextYear[i,1] <- ProbNextYear[i,1]/(1+ARR)^(i-1)
}

for(i in 1:nrow(ProbNextYear)){
  LifeExpectancy[i,1] <- sum(ProbNextYear[i:nrow(ProbNextYear),]) / ProbNextYear[i,1]
}

AnnuityFactor <- ProbNextYear
colnames(AnnuityFactor) <- 'Annuity Factor'
AnnuityFactor[1:20,1] <- 1

for(j in 45:115){
  sum <- 0
  RowIndex <- j - 25 + 1
  for(i in RowIndex:nrow(ProbNextYear)-5){
    if(as.double(Age[i,1]) == j){
      sum <- sum + ProbNextYear[i,1]
      TempValue <- as.double(ProbNextYear[(RowIndex + 5),1])
    } else {
      if(COLACompound == 1){
        sum <- sum + as.double(ProbNextYear[i,1])*((1+COLA)^(as.double(Age[i,1]) - j))
        TempValue <- as.double(ProbNextYear[(RowIndex + 5),1])*((1+COLA)^(as.double(Age[(RowIndex + 5),1]) - j))
      } else {
        sum <- sum + as.double(ProbNextYear[i,1])*(1 +(COLA)*(as.double(Age[i,1]) - j))
        TempValue <- as.double(ProbNextYear[(RowIndex + 5),1])*(1 +(COLA)*(as.double(Age[(RowIndex + 5),1]) - j))
      }
    }
    
    if(i == (RowIndex + 5)){
      FactorDiv <- TempValue
    }
  }
  AnnuityFactor[RowIndex,1] <- sum / FactorDiv
}