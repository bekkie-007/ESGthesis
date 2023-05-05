######################

# Set working directory 

setwd("~/Desktop/RB schrijft een THESIS/Programming")

######################

# Install and library packages needed

#install.packages("Ecdat")
#install.packages("quadprog")
#install.packages("readxl")
#install.packages("Matrix")
#install.packages("dplyr")
#install.packages("matrixcalc")
#install.packages("linprog")
library(Ecdat)
library(quadprog)
library(readxl)
library(Matrix)
library(dplyr)
library(matrixcalc)
library(linprog)

######################

# Input variables, constants that can be changed for stress testing #

mufree = 4.8/253 # input value of risk-free interest rate, based on 1 year US treasury yield

#choose between STOXX50 or DOWJONES
portfolio = "STOXX50"
portfolio_w  = "STOXX50_W"
portfolio_g = "STOXX50_G"
portfolio_esg = "STOXX50_ESG_AVG"
portfolio_esg_w = "STOXX50_ESG_AVG_W"
esg_weapons = "EU"
numberweaponstocks = 12

#portfolio = "DOWJONES"
#portfolio_w = "DOWJONES_W"
#portfolio_g = "DOWJONES_G"
#portfolio_esg = "DOWJONES_ESG_AVG"
#portfolio_esg_w = "DOWJONES_ESG_AVG_W"
#esg_weapons = "US"
#numberweaponstocks = 23

######################

# Define starting values 

Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio,skip=3)
Data_esg = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio_esg,skip=3)
max = max(Data$Number)+1 #number of monthly portfolios
v1 <- seq(0.05, 0.001, -0.005) #upperbounds for weapon stocks
constraints = length(v1)

# Matrices and vectors to store output

returnvector = matrix(0,max,1)
herfindahl = matrix(0,max,1)
maxdrawdownvector = matrix(0,max,1) #vector to store the monthly maximum drawdown
sharperatiovector = matrix(0,max,1) #vector to store the monthly sharpe ratio
esgratingvector = matrix(0,max,1) #vector to store monthly esg rating
percentageweaponstocksvector = matrix(0,max,1) #vector to store percentage weapon stocks
counter = 0 
returnvector_w = matrix(0,max,1)
herfindahl_w = matrix(0,max,1)
maxdrawdownvector_w = matrix(0,max,1) #vector to store the monthly maximum drawdown
sharperatiovector_w = matrix(0,max,1) #vector to store the monthly sharpe ratio
esgratingvector_w = matrix(0,max,1) #vector to store monthly esg rating
counter_w = 0
counter2 = 0
herfmatrix = matrix(0,max,constraints) 
maxdrawdownmatrix = matrix(0,max,constraints) #vector to store the monthly maximum drawdown
sharperatiomatrix = matrix(0,max,constraints) #vector to store the monthly sharpe ratio
esgratingmatrix = matrix(0,max,constraints) #vector to store monthly esg rating
percentageweaponstocksmatrix = matrix(0,max,constraints) #vector to store percentage weapon stocks
#maxdrawdownvector_g = matrix(0,max,1) #vector to store the monthly maximum drawdown
#sharperatiovector_g = matrix(0,max,1) #vector to store the monthly sharpe ratio
#counter_g = 0 
######################

# Method reshuffle
counter=0
for(i in 1:max){
  counter=counter+1
  Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio,skip=3)
  Data = Data %>% filter(Number == i-1)
  Data_date = subset(Data, select = -c(Date, Month, Number))
  mean_vect = apply(Data_date,2,mean)
  capvector = rep(-1,length(mean_vect))
  weapvector = rep(0,length(v1))
  herf3d = rep(0,length(v1))
  cc = 0
  for(j in v1){
    for(m in 1:numberweaponstocks){
      capvector[m]=(-j)
    }
    cov_mat = cov(Data_date)
    cc = cc+1
    if(is.positive.definite(cov_mat)==FALSE){
      cov_mat = nearPD(cov_mat)$mat
    }
    sd_vect = sqrt(diag(cov_mat))
    Amat = cbind(rep(1,length(mean_vect)),mean_vect, diag(-1, nrow = length(mean_vect)), diag(1, nrow = length(mean_vect)))  # set the constraints matrix
    muP = seq(min(mean_vect)+.001,max(mean_vect)-.001,length=300)  # set of 300 possible target values
    sdP = muP # set up storage for std dev's of portfolio returns
    weights = matrix(0,nrow=300,ncol=dim(Data_date)[2]) # storage for portfolio weights
    for (i in 1:length(muP))  # find the optimal portfolios for each target expected return
    {
      bvec = c(1,muP[i], capvector, rep(-0.05,length(mean_vect)))  # constraint vector
      result = solve.QP(Dmat=2*cov_mat,dvec=rep.int(0,dim(Data_date)[2]),Amat=Amat,bvec=bvec,meq=2)
      sdP[i] = sqrt(result$value)
      weights[i,] = result$solution
    }
    
    sharpe =( muP-mufree)/sdP # compute Sharpe's ratios
    ind = (sharpe == max(sharpe)) # Find maximum Sharpe's ratio
    ind2 = (sdP == min(sdP)) # find the minimum variance portfolio
    weights_minvar = c(weights[ind2,])
    returns = t(weights_minvar) %*% t(as.matrix(Data_date))
    maxdrawdown = max(returns)-min(returns)
    sharperatio = max(sharpe)
    percentageweaponstocks = sum(weights_minvar[1:numberweaponstocks])*100
    
    herfmatrix[counter,cc] = 1/sum(weights_minvar^2)
    maxdrawdownmatrix[counter,cc] = maxdrawdown #vector to store the monthly maximum drawdown
    sharperatiomatrix[counter,cc] = sharperatio #vector to store the monthly sharpe ratio
    percentageweaponstocksmatrix[counter,cc] = percentageweaponstocks #vector to store percentage weapon stocks
    
  }
}

#function to get mimimum variance portfolio with a specific cap (upperbound) on weapon stocks

Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio,skip=3)
Data_esg = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio_esg,skip=3)
Data_esg = subset(Data_esg, select = -c(Date, Month, Number))
Data_date = subset(Data, select = -c(Date, Month, Number))
mean_vect = apply(Data_date,2,mean)
cov_mat = cov(Data_date)
capvector = rep(-1,length(mean_vect))
weapvector = rep(0,length(v1))
herf3d = rep(0,length(v1))
compoundreturns = 1
cc = 0
fc <- colorRampPalette(c("red", "blue")) #definecolours
for(j in v1){
  for(m in 1:numberweaponstocks){
    capvector[m]=(-j)
  }
  cc = cc+1
  if(is.positive.definite(cov_mat)==FALSE){
    cov_mat = nearPD(cov_mat)$mat
  }
  sd_vect = sqrt(diag(cov_mat))
  Amat = cbind(rep(1,length(mean_vect)),mean_vect, diag(-1, nrow = length(mean_vect)), diag(1, nrow = length(mean_vect)))  # set the constraints matrix
  muP = seq(min(mean_vect)+.001,max(mean_vect)-.001,length=300)  # set of 300 possible target values
  # for the expect portfolio return
  sdP = muP # set up storage for std dev's of portfolio returns
  weights = matrix(0,nrow=300,ncol=dim(Data_date)[2]) # storage for portfolio weights
  for (i in 1:length(muP))  # find the optimal portfolios for each target expected return
  {
    bvec = c(1,muP[i], capvector, rep(-0.01,length(mean_vect)))  # constraint vector
    result = solve.QP(Dmat=2*cov_mat,dvec=rep.int(0,dim(Data_date)[2]),Amat=Amat,bvec=bvec,meq=2)
    sdP[i] = sqrt(result$value)
    weights[i,] = result$solution
  }
  sharpe =( muP-mufree)/sdP # compute Sharpe's ratios
  ind = (sharpe == max(sharpe)) # Find maximum Sharpe's ratio
  ind2 = (sdP == min(sdP)) # find the minimum variance portfolio
  weights_minvar = c(weights[ind2,])
  print(sum(weights_minvar[1:numberweaponstocks]))
  returns = t(weights_minvar) %*% t(as.matrix(Data_date))
  print(mean((cumprod(returns/100+1)-1)))
  if(cc==1){
    plot(Data$Date,(cumprod(returns/100+1)-1),type="l",lty=1,col=fc(10)[cc],main="Cumulative returns",xlab="",ylab="Cumulative return")  #  plot
  }
  lines(Data$Date,(cumprod(returns/100+1)-1),type="l",lty=1,col=fc(10)[cc])
  weapvector[cc] = round(percentageweaponstocks*100,2)
  herf3d[cc] = 1/sum(weights_minvar^2)
}
