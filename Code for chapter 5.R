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
#portfolio = "STOXX50"
#portfolio_w  = "STOXX50_W"
#portfolio_g = "STOXX50_G"
#portfolio_esg = "STOXX50_ESG_AVG"
#portfolio_esg_w = "STOXX50_ESG_AVG_W"
#esg_weapons = "EU"
#numberweaponstocks = 12

portfolio = "DOWJONES"
portfolio_w = "DOWJONES_W"
portfolio_g = "DOWJONES_G"
portfolio_esg = "DOWJONES_ESG_AVG"
portfolio_esg_w = "DOWJONES_ESG_AVG_W"
esg_weapons = "US"
numberweaponstocks = 23

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

Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio,skip=3)
Data_esg = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio_esg,skip=3)
Data_date = subset(Data, select = -c(Date, Month, Number))
Data_esg = subset(Data_esg, select = -c(Date, Month, Number))
mean_vect = apply(Data_date,2,mean)
cov_mat = cov(Data_date)
if(is.positive.definite(cov_mat)==FALSE){
  cov_mat = nearPD(cov_mat)$mat
}
sd_vect = sqrt(diag(cov_mat))
Amat = cbind(rep(1,length(mean_vect)),mean_vect, diag(1, nrow = length(mean_vect)))  # set the constraints matrix
muP = seq(min(mean_vect)+.001,max(mean_vect)-.001,length=300)  # set of 300 possible target values
# for the expect portfolio return
sdP = muP # set up storage for std dev's of portfolio returns
weights = matrix(0,nrow=300,ncol=dim(Data_date)[2]) # storage for portfolio weights
for (i in 1:length(muP))  # find the optimal portfolios for each target expected return
{
  bvec = c(1,muP[i], rep(0,length(mean_vect)))  # constraint vector
  result = solve.QP(Dmat=2*cov_mat,dvec=rep.int(0,dim(Data_date)[2]),Amat=Amat,bvec=bvec,meq=2)
  sdP[i] = sqrt(result$value)
  weights[i,] = result$solution
}

sharpe =( muP-mufree)/sdP # compute Sharpe's ratios
ind = (sharpe == max(sharpe)) # Find maximum Sharpe's ratio
ind2 = (sdP == min(sdP)) # find the minimum variance portfolio
weights_minvar = c(weights[ind2,])
for(j in 1:length(weights_minvar)){ # as I use an optimization algorithm, this is a double check to see if all of the weights are indeed>0
  if(weights_minvar[j]< 0){
    weights_minvar[j] = 0
  }
}

sharpe =( muP-mufree)/sdP # compute Sharpe's ratios
ind = (sharpe == max(sharpe)) # Find maximum Sharpe's ratio
ind2 = (sdP == min(sdP)) # find the minimum variance portfolio
weights_minvar = c(weights[ind2,])
esgratingminvar = mean(t(weights_minvar) %*% t(as.matrix(Data_esg)))
esgratingminvar
percentageweaponstocks = sum(weights_minvar[1:numberweaponstocks])
percentageweaponstocks
returns = t(weights_minvar) %*% t(as.matrix(Data_date))

Data_w = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio_w,skip=3)
Data_esg_w = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio_esg_w,skip=3)
Data_date_w = subset(Data_w, select = -c(Date, Month, Number))
Data_esg_w = subset(Data_esg_w, select = -c(Date, Month, Number))
mean_vect_w = apply(Data_date_w,2,mean)
cov_mat_w = cov(Data_date_w)
if(is.positive.definite(cov_mat_w)==FALSE){
  cov_mat_w = nearPD(cov_mat_w)$mat
}
sd_vect_w = sqrt(diag(cov_mat_w))
Amat_w = cbind(rep(1,length(mean_vect_w)),mean_vect_w, diag(1, nrow = length(mean_vect_w)))  # set the constraints matrix
muP_w = seq(min(mean_vect_w)+.001,max(mean_vect_w)-.001,length=300)  # set of 300 possible target values
# for the expect portfolio return
sdP_w = muP_w # set up storage for std dev's of portfolio returns
weights_w = matrix(0,nrow=300,ncol=dim(Data_date_w)[2]) # storage for portfolio weights
for (i in 1:length(muP_w))  # find the optimal portfolios for each target expected return
{
  bvec_w = c(1,muP_w[i], rep(0,length(mean_vect_w)))  # constraint vector
  result_w = solve.QP(Dmat=2*cov_mat_w,dvec=rep.int(0,dim(Data_date_w)[2]),Amat=Amat_w,bvec=bvec_w,meq=2)
  sdP_w[i] = sqrt(result_w$value)
  weights_w[i,] = result_w$solution
  for(i in 1:length(weights_w)){
    if(weights_w[i]< 0){
      weights_w[i] = 0
    }
  }
}

sharpe_w =( muP_w-mufree)/sdP_w # compute Sharpe's ratios
ind_w= (sharpe_w == max(sharpe_w)) # Find maximum Sharpe's ratio
ind2_w = (sdP_w == min(sdP_w)) # find the minimum variance portfolio
weights_minvar_w = c(weights_w[ind2_w,])
esgratingminvar_w = mean(t(weights_minvar_w) %*% t(as.matrix(Data_esg_w)))
esgratingminvar_w
returns_w = t(weights_minvar_w) %*% t(as.matrix(Data_date_w))

#function to get minimum variance portfolio for each month (for portfolio with weapons)

for(i in 1:max){
  counter = counter+1
  Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio,skip=3)
  Data = Data %>% filter(Number == i-1)
  Data_esg = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio_esg,skip=3)
  Data_esg = Data_esg %>% filter(Number == i-1) #to get the minimum variance portfolio for each month
  Data_esg = subset(Data_esg, select = -c(Date, Month, Number))
  Data_date = subset(Data, select = -c(Date, Month, Number))
  mean_vect = apply(Data_date,2,mean)
  cov_mat = cov(Data_date)
  if(is.positive.definite(cov_mat)==FALSE){
    cov_mat = nearPD(cov_mat)$mat
  }
  sd_vect = sqrt(diag(cov_mat))
  Amat = cbind(rep(1,length(mean_vect)),mean_vect, diag(1, nrow = length(mean_vect)))  # set the constraints matrix
  muP = seq(min(mean_vect)+.001,max(mean_vect)-.001,length=300)  # set of 300 possible target values
  # for the expect portfolio return
  sdP = muP # set up storage for std dev's of portfolio returns
  weights = matrix(0,nrow=300,ncol=dim(Data_date)[2]) # storage for portfolio weights
  for (i in 1:length(muP))  # find the optimal portfolios for each target expected return
  {
    bvec = c(1,muP[i], rep(0,length(mean_vect)))  # constraint vector
    result = solve.QP(Dmat=2*cov_mat,dvec=rep.int(0,dim(Data_date)[2]),Amat=Amat,bvec=bvec,meq=2)
    sdP[i] = sqrt(result$value)
    weights[i,] = result$solution
  }
  
  sharpe =( muP-mufree)/sdP # compute Sharpe's ratios
  ind= (sharpe == max(sharpe)) # Find maximum Sharpe's ratio
  ind2 = (sdP == min(sdP)) # find the minimum variance portfolio
  weightstan = c(weights[ind,])
  weightsminvar = c(weights[ind2,])
  #weights = weightstan
  weights = weightsminvar
  returns = t(weights) %*% t(as.matrix(Data_date))
  cumureturns = cumprod((returns/100)+1)
  percentageweaponstocks = sum(weights[1:numberweaponstocks])*100
  #maxdrawdown = min(cumureturns)/max(cumureturns[1:which.min(cumureturns)])-1
  maxdrawdown = (max(returns)-min(returns))
  sharperatio = max(sharpe)
  herf = 1 / sum(weights^2)
  esgrating = mean(t(weights) %*% t(as.matrix(Data_esg)))
  herfindahl[counter] = herf
  maxdrawdownvector[counter] = maxdrawdown
  sharperatiovector[counter] = sharperatio
  esgratingvector[counter] = esgrating
  percentageweaponstocksvector[counter] = percentageweaponstocks
}



#I do the exact same thing for the portfolio without weapons
#function to get minimum variance portfolio for each month (for portfolio without weapons)
for(i in 1:max){
  counter_w = counter_w+1
  Data_w = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio_w,skip=3)
  Data_w = Data_w %>% filter(Number == i-1)
  Data_esg_w = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio_esg_w,skip=3)
  Data_esg_w = Data_esg_w %>% filter(Number == i-1) #to get the minimum variance portfolio for each month
  Data_esg_w = subset(Data_esg_w, select = -c(Date, Month, Number))
  Data_date_w = subset(Data_w, select = -c(Date, Month, Number))
  mean_vect_w = apply(Data_date_w,2,mean)
  cov_mat_w = cov(Data_date_w)
  if(is.positive.definite(cov_mat_w)==FALSE){
    cov_mat_w = nearPD(cov_mat_w)$mat
  }
  sd_vect_w = sqrt(diag(cov_mat_w))
  Amat_w = cbind(rep(1,length(mean_vect_w)),mean_vect_w, diag(1, nrow = length(mean_vect_w)))  # set the constraints matrix
  muP_w = seq(min(mean_vect_w)+.001,max(mean_vect_w)-.001,length=300)  # set of 300 possible target values
  # for the expect portfolio return
  sdP_w = muP_w # set up storage for std dev's of portfolio returns
  weights_w = matrix(0,nrow=300,ncol=dim(Data_date_w)[2]) # storage for portfolio weights
  for (i in 1:length(muP_w))  # find the optimal portfolios for each target expected return
  {
    bvec_w = c(1,muP_w[i], rep(0,length(mean_vect_w)))  # constraint vector
    result_w = solve.QP(Dmat=2*cov_mat_w,dvec=rep.int(0,dim(Data_date_w)[2]),Amat=Amat_w,bvec=bvec_w,meq=2)
    sdP_w[i] = sqrt(result_w$value)
    weights_w[i,] = result_w$solution
  }
  
  sharpe_w =( muP_w-mufree)/sdP_w # compute Sharpe's ratios
  ind_w= (sharpe_w == max(sharpe_w)) # Find maximum Sharpe's ratio
  ind2_w = (sdP_w == min(sdP_w)) # find the minimum variance portfolio
  for(i in 1:length(weights_w)){
    if(weights_w[i]< 0){
      weights_w[i] = 0
    }
  }
  weightstan_w = c(weights_w[ind,])
  weightsminvar_w = c(weights_w[ind2,])
  weights_w = weightsminvar_w
  returns_w = t(weights_w) %*% t(as.matrix(Data_date_w))
  cumreturns_w = cumprod(returns_w/100+1)
  herf_w = 1 / sum(weights_w^2)
  maxdrawdown_w = (max(returns_w)-min(returns_w))
  sharperatio_w = max(sharpe_w)
  esgrating_w = mean(t(weights_w) %*% t(as.matrix(Data_esg_w)))
  herfindahl_w[counter_w] = herf_w
  returnvector_w[counter_w] = mean(returns)
  maxdrawdownvector_w[counter_w] = maxdrawdown_w
  sharperatiovector_w[counter_w] = sharperatio_w
  esgratingvector_w[counter_w] = esgrating_w
}