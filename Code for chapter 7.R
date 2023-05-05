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
portfolio = "ETF_STOXX"
w = 0.2933227
#portfolio2 = "STOXX_W"

#portfolio2 = "DOWJONES_W"
#portfolio = "ETF_DOWJONES"
#w = 0.3437114


######################

# Define starting values 

Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio)
Data2 = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio2,skip=3)
max = max(Data$Number)+1 #number of monthly portfolios
max2 = max(Data2$Number)+1 #number of monthly portfolios
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
cumureturnsvector = matrix(0,dim(Data)[1],1)
cumureturnsvector2 = matrix(0,dim(Data2)[1],1)
returncounter = 0
#maxdrawdownvector_g = matrix(0,max,1) #vector to store the monthly maximum drawdown
#sharperatiovector_g = matrix(0,max,1) #vector to store the monthly sharpe ratio
#counter_g = 0 
######################

Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio)
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



#function to get minimum variance portfolio for each month 

for(i in 1:max){
  counter = counter+1
  Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio)
  Data = Data %>% filter(Number == i-1)
  Data_date = subset(Data, select = -c(Date, Month, Number))
  mean_vect = apply(Data_date,2,mean)
  cov_mat = cov(Data_date)
  if(is.positive.definite(cov_mat)==FALSE){
    cov_mat = nearPD(cov_mat)$mat
  }
  sd_vect = sqrt(diag(cov_mat))
  Amat = cbind(rep(1,length(mean_vect)),mean_vect, diag(1, nrow = length(mean_vect)))  # set the constraints matrix
  muP = seq(min(mean_vect)+.0001,max(mean_vect)-.0001,length=300)  # set of 300 possible target values
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
  w1 = w1 + weights[1]
  w2 = w2 + weights[2]
  returns = t(weights) %*% t(as.matrix(Data_date))
  cumureturns_etf = returns
  #percentageweaponstocks = sum(weights[1:numberweaponstocks])*100
  #print(weights)
  #maxdrawdown = min(cumureturns)/max(cumureturns[1:which.min(cumureturns)])-1
  maxdrawdown = (max(returns)-min(returns))
  sharperatio = max(sharpe)
  maxdrawdownvector[counter] = maxdrawdown
  sharperatiovector[counter] = sharperatio
  for(i in 1:length(cumureturns_etf)){
    cumureturnsvector[i+returncounter] = returns[i]
  
  }
  returncounter = returncounter + length(cumureturns_etf)
}

counter = 0
returncounter = 0


for(i in 1:max2){
  counter = counter+1
  Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio2,skip=3)
  Data = Data %>% filter(Number == i-1)
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
  w1 = w1 + weights[1]
  w2 = w2 + weights[2]
  returns = t(weights) %*% t(as.matrix(Data_date))
  cumureturns_etf = returns
  #percentageweaponstocks = sum(weights[1:numberweaponstocks])*100
  #print(weights)
  #maxdrawdown = min(cumureturns)/max(cumureturns[1:which.min(cumureturns)])-1
  maxdrawdown = (max(returns)-min(returns))
  sharperatio = max(sharpe)
  maxdrawdownvector[counter] = maxdrawdown
  sharperatiovector[counter] = sharperatio
  for(i in 1:length(cumureturns_etf)){
    cumureturnsvector2[i+returncounter] = returns[i]
    
  }
  returncounter = returncounter + length(cumureturns_etf)
}

cumureturnsvector2 = cumureturnsvector2[1:2008]
Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio)
Data2 = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio2,skip=3)
plot(Data$Date,(cumprod(cumureturnsvector/100+1)-1),type="l",lty=1,col="blue",xlab="",ylab="Cumulative return")  #  plot
#lines(Data2$Date,(cumprod(cumureturnsvector2/100+1)-1),type="l",lty=1,col="blue")  #  plot

legend("topleft", legend=c("First strategy", "Second strategy", "Index without weapons"),
       col=c("red", "blue","green"), cex=1, lty=1)

returns_x  = t(c(0,1)) %*% t(as.matrix(subset(Data, select = -c(Date, Month, Number))))
returns_p = t(c(w,1-w)) %*% t(as.matrix(subset(Data, select = -c(Date, Month, Number))))
lines(Data$Date,(cumprod(returns_x/100+1)-1),type="l",lty=1,col="green")
lines(Data$Date,(cumprod(returns_p/100+1)-1),type="l",lty=1,col="red")
#lines(Data$Date,(cumprod(returns/100+1)-1),type="l",lty=1,col="brown")
#lines(Data$Date,(cumprod(returns_w/100+1)-1),type="l",lty=1,col="blue")
