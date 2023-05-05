#Plot results 

par(mfrow=c(1,2))
plot(maxdrawdownvector,type="l",main="Maximum drawdown and sharpe ratio",ylim=c(0,max(maxdrawdownvector)),col="blue", xaxt='n',xlab="Time",yaxt='n',ylab="Value")  #  plot
axis(1, at=c(0,65,85,90), labels=c(2014,"03-'20",2022,"03-'22"))
lines(sharperatiovector,type="l", col="red")  #  plot
axis(2,at=c(seq(0, 10, 1)), labels=c(seq(0, 10, 1)))
#lines(esgratingvector,type="l", col="dark green")  #  plot
legend(1, 10, legend=c("Sharpe Ratio", "Max Drawdown"),
       col=c("red", "blue"), cex=1, lty=1)
plot(maxdrawdownvector_w,type="l",main="Maximum drawdown and sharpe ratio without weapons",ylim=c(0,max(maxdrawdownvector_w)),col="blue", xaxt='n',xlab="Time",yaxt='n',ylab="Value")  #  plot
axis(1, at=c(0,65,85,90), labels=c(2014,"03-'20",2022,"03-'22"))
lines(sharperatiovector_w,type="l", col="red")  #  plot
axis(2,at=c(seq(0, 15, 1)), labels=c(seq(0, 15, 1)))
#lines(esgratingvector_w,type="l", col="dark green")  #  plot
legend(1, 10, legend=c("Sharpe Ratio", "Max Drawdown"),
       col=c("red", "blue"), cex=1, lty=1)
par(mfrow=c(2,2))
plot(esgratingvector,type="l",main="Average ESG Rating Minimum Variance Portfolio with weapons",ylim=c(min(esgratingvector)-1,max(esgratingvector)+1),col="dark green", xaxt='n',xlab="Time",yaxt='n',ylab="Value")  #  plot
axis(2,at=c(seq(50, 80, 1)), labels=c(seq(50, 80, 1)))
x = (1:length(esgratingvector))
lines(predict(lm(esgratingvector~x)),col='red') #trend line
axis(1, at=c(0,65,85,90), labels=c(2014,"03-'20",2022,"03-'22"))
plot(esgratingvector_w,type="l",main="Average ESG Rating Minimum Variance Portfolio without weapons",ylim=c(min(esgratingvector_w)-1,max(esgratingvector_w)+1),col="dark green", xaxt='n',xlab="Time",yaxt='n',ylab="Value")  #  plot
axis(2,at=c(seq(50, 80, 1)), labels=c(seq(50, 80, 1)))
x = (1:length(esgratingvector_w))
lines(predict(lm(esgratingvector_w~x)),col='red') #trend line
axis(1, at=c(0,65,85,90), labels=c(2014,"03-'20",2022,"03-'22"))

plot(percentageweaponstocksvector,type="l",main="Percentage weapon stocks",ylim=c(min(percentageweaponstocksvector)-1,max(percentageweaponstocksvector)+1),col="brown", xaxt='n',xlab="Time",yaxt='n',ylab="Percentage")  #  plot
x = (1:length(percentageweaponstocksvector))
lines(predict(lm(percentageweaponstocksvector~x)),col='red') #trend line
axis(1, at=c(0,65,85,90), labels=c(2014,"03-'20",2022,"03-'22"))
axis(2,at=c(seq(0, 100, 5)), labels=c(seq(0, 100, 5)))

#outlierplot
plot(sdP,muP,type="l",lty=3,col="brown",main="Efficient frontier outlier")  #  plot
points(0,mufree,cex=4,pch="*")  # show risk-free asset
sharpe =( muP-mufree)/sdP # compute Sharpe's ratios
ind = (sharpe == max(sharpe)) # Find maximum Sharpe's ratio
lines(c(0,2),mufree+c(0,2)*(muP[ind]-mufree)/sdP[ind],lwd=4,lty=1, col = "brown")
# show line of optimal portfolios
points(sdP[ind],muP[ind],cex=4,pch="*",col = "brown") # show tangency portfolio
ind2 = (sdP == min(sdP)) # find the minimum variance portfolio
points(sdP[ind2],muP[ind2],cex=2,pch="+",col="brown") # show min var portfolio
ind3 = (muP > muP[ind2])
lines(sdP[ind3],muP[ind3],type="l",xlim=c(0,.25),
      ylim=c(0,.3),lwd=3, col = "brown")

par(mfrow=c(1,1))
#Plot
plot(sdP,muP,type="l",xlim=c(0,2),ylim=c(-.2,.2),lty=3,col="brown",main="Efficient frontier")  #  plot
legend(0.1, 0.2, legend=c("With weapons", "Without weapons"),
       col=c("brown", "blue"), cex=1.5, lty=1)
lines(sdP_w,muP_w,type="l",lty=3, col="blue")  #  plot
#lines(sdP_g,muP_g,type="l",lty=3, col="dark green")  #  plot
Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet="ETF_STOXX")
plot(Data$Date,(cumprod(cumureturnsvector/100+1)-1),type="l",lty=1,col="green",main="Cumulative returns",xlab="",ylab="Cumulative return")  #  plot
legend("topleft", legend=c("First passive strategy", "Second passive strategy"),
       col=c("red", "green"), cex=1.5, lty=1)
returns_p = t(c(w,1-w)) %*% t(as.matrix(subset(Data, select = -c(Date, Month, Number))))
lines(Data$Date,(cumprod(returns_p/100+1)-1),type="l",lty=1,col="red")
Data = read_excel("~/Desktop/RB schrijft een THESIS/Programming/STOXXweapons.xlsx",sheet=portfolio,skip=3)
lines(Data$Date,(cumprod(returns/100+1)-1),type="l",lty=1,col="brown",main="Cumulative returns",xlab="",ylab="Cumulative return")  #  plot
#legend(0.1, 0.2, legend=c("With weapons", "Without weapons"),
#       col=c("brown", "blue"), cex=1.5, lty=1)
lines(Data_w$Date,(cumprod(returns_w/100+1)-1),type="l",lty=1,col="blue")

# the efficient frontier (and inefficient portfolios
# below the min var portfolio)
points(0,mufree,cex=4,pch="*")  # show risk-free asset
sharpe =( muP-mufree)/sdP # compute Sharpe's ratios
ind = (sharpe == max(sharpe)) # Find maximum Sharpe's ratio
weights[ind,] #  print the weights of the tangency portfolio
lines(c(0,2),mufree+c(0,2)*(muP[ind]-mufree)/sdP[ind],lwd=4,lty=1, col = "brown")
# show line of optimal portfolios
points(sdP[ind],muP[ind],cex=4,pch="*",col = "brown") # show tangency portfolio
ind2 = (sdP == min(sdP)) # find the minimum variance portfolio
points(sdP[ind2],muP[ind2],cex=2,pch="+",col="brown") # show min var portfolio
ind3 = (muP > muP[ind2])
lines(sdP[ind3],muP[ind3],type="l",xlim=c(0,.25),
      ylim=c(0,.3),lwd=3, col = "brown")  #  plot the efficient frontier
#blue
points(0,mufree,cex=4,pch="*",col="blue")  # show risk-free asset
sharpe_w =( muP_w-mufree)/sdP_w # compute Sharpe's ratios
ind_w = (sharpe_w == max(sharpe_w)) # Find maximum Sharpe's ratio
weights_w[ind_w,] #  print the weights of the tangency portfolio
lines(c(0,2),mufree+c(0,2)*(muP[ind_w]-mufree)/sdP_w[ind_w],lwd=4,lty=1, col = "blue")
# show line of optimal portfolios
points(sdP_w[ind_w],muP_w[ind_w],cex=4,pch="*",col="blue") # show tangency portfolio
ind2_w = (sdP_w == min(sdP_w)) # find the minimum variance portfolio
points(sdP_w[ind2_w],muP_w[ind2_w],cex=2,pch="+",col="blue") # show min var portfolio
ind3_w = (muP_w > muP_w[ind2_w])
lines(sdP_w[ind3_w],muP_w[ind3_w],type="l",xlim=c(0,.25),
      ylim=c(0,.3),lwd=3, col = "blue")  #  plot the efficient frontier
