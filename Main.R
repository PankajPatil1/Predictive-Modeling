library(readxl)
library(e1071)
require(forecast)
require(xlsx)

#Read the data
macro <- read.csv("Econometric data.csv")
main <- read_excel("Shipment Data.xlsx",sheet=1)
actual.data <- read_excel("Actual Data for Evaluation.xlsx",sheet=2)
colnames(macro)[1] <- "QTR"
main.merged <- merge(main,macro,by="QTR")


# Aggregate production by quarter for each product
x <- aggregate(INDUSTRY_UNITS~QTR+PRODUCT,main,sum)

macro.train <- macro[c(-24,-23),]
macro.test <- macro[c(23,24),]
actual.data <- actual.data[c(-27,-28),]
actual.data$Predicted_INDUSTRY_UNITS <- rep(0,each=26)

#Make new data frame in which products are encoded in numerical format (like 1 = 2 Door Bottom Mount,3=Cooktops etc.)
#For actual mapping of data, please check solution sheet / "actual.data" dataframe 
main2 <- x
main2$PRODUCT <- as.numeric(as.factor(x$PRODUCT))
main2 = merge(main2,macro.train,by="QTR")
#For each product, isolate it and make it into quarterly time series and use HoltWinters function on the timeseries
#Use forecast.Holtwinters to find the final predictions
#This code does not include model selection which included splitting the data into training and test data
#This code also does not include Variable selection since HoltWinters does not require external regressors
for(i in 1:13)
{
        # i=1
        main2 = merge(main2,macro.train,by="QTR")
       tmp <- main2[main2$PRODUCT==i,]
       ck <- ts(tmp[,3],frequency=4,start=c(2010,1))
       hwck <- HoltWinters(ck,seasonal="additive")
       ckforecast <- forecast.HoltWinters(hwck,h=2)
       actual.data[c(i,i+13),4]<- ceiling(ckforecast$mean)
}
#Find absolute errors and write the data into .xlsx spreadsheet
actual.data$Predicted_INDUSTRY_UNITS <- as.numeric(actual.data$Predicted_INDUSTRY_UNITS)
actual.data$Error <- abs(actual.data$`Actual INDUSTRY_UNITS` - actual.data$Predicted_INDUSTRY_UNITS)/actual.data$`Actual INDUSTRY_UNITS`
write.xlsx(actual.data,"Solution.xlsx",sheet="Solution")
