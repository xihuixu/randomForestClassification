getwd()
setwd("/Users/xx/Desktop/Apotex case package/Data Sets/")
load("yoursession.RData")
library("readxl")
library("tidyverse")
library("caret")
library("rpart")
library("openxlsx")
library("parallel")
library("doParallel")
library("randomForest")
library("dplyr")
library("MLeval")

data = read_excel("Bid Data.xlsx")

#data = read.csv("Bid Data.csv")
data = as.data.frame(data)
dim(data)
#客人
length(unique(data[,1]))
#material
length(unique(data[,2]))
#families
length(unique(data[,3]))

colnames(data) = gsub("\r\n|\\s*\\s|\\s*\n", "", colnames(data))
colnames(data)


sapply(data, class)
#######DATA ClEANING
data[,c(1,2,3,8)] = lapply(data[,c(1,2,3,8)],factor)
data[,6]= factor(toupper(data[,6]))
weird_index = which(grepl("/", data[,7]))
weird_dates = substr(data[weird_index,7], start = 2, stop = length(data[weird_index,7]))
data[,7] = ifelse(grepl("/", data[,7]), NA, as.character(convertToDateTime(data[,7],origin = "1899-12-30")))
data[weird_index,7] = format(as.Date(weird_dates,format = "%m/%d/%Y"),"%Y-%m-%d")


# target: 100% successful at its first attempt at its bidding, so removing not first attempt
data = data %>% filter(data[,10] == "Approved")
data[data[,9] == "2017",12] = NA 

#data[,4] = gsub(" ", "", data[,4], fixed = TRUE)
#data[,4] = as.numeric(factor(data[,4]))

data[,11] <- factor(data[,11],levels = c("Won", "Lost"))


data = data[,-c(10)]
# total
table(data$`CustomerResponse`)/nrow(data)  


colSums(is.na(data))


data = data %>% add_count(data$ProductFamily) %>%
  mutate(famAppearance = ifelse(n <50, "<50", 
                                ifelse(n <100, "50-100", 
                                       ifelse( n <150, "100-150", 
                                               ifelse( n <200, "150-200",
                                                       ifelse( n <300, "200-300",
                                                               ifelse( n <500, "300-500", "500-600")))))))



data = data[,-c(12,13)]

data2 = data %>% add_count(data$CustomerID)
hist(data2$n, main="Count for each Customer", xlab='Customer Frequency', freq=F, col="rosybrown")
data = data %>% add_count(data$CustomerID) %>% 
  mutate(patron =   
           ifelse(n < 100, "<100", 
                  ifelse( n <200, "100-200",
                          ifelse( n <300, "200-300",
                                  ifelse( n < 400, "300-400",
                                          ifelse( n < 500, "400-500",
                                                  ifelse( n <600, "500-600", ">600")))))))

data = data[,-c(13,14)]

data3 = data %>% add_count(data$MaterialNumber)
hist(data3$n, main="Count for each Material", xlab='Material Frequency', freq=F)
data3 = data3%>%
  mutate(MaterAppearance = ifelse(n <10, "<10", 
                                ifelse(n <20, "10-20", 
                                       ifelse( n <30, "20-30", 
                                               ifelse( n <40, "30-40",
                                                       ifelse( n <50, "40-50",
                                                               ifelse( n <60, "50-60", ">60")))))))


data$MaterialAppearance = factor(data3$MaterAppearance)
data[,c(12,13)] = lapply(data[,c(12,13)] ,factor)
data = data.frame(data)
data$FailuretoSupplyRate.FTS.perunit = as.numeric(data$FailuretoSupplyRate.FTS.perunit)

data$CustomerSupplyStartdate = format(as.Date(data$CustomerSupplyStartdate, "%Y-%m-%d"),'%d-%m-%Y')
ps = paste0('01','-',data[,8],'-',data[,9])

data$BidSubmittedMY = format(as.Date(ps,"%d-%b-%Y"),'%d-%m-%Y')
datetimes = lapply(data.frame(bidSub = data$BidSubmittedMY,supply=data$CustomerSupplyStartdate), strptime, c(format = '%d-%m-%Y'))
datetimes = data.frame(datetimes)
datetimes$diff_in_month = floor(as.double(difftime(datetimes[,2], datetimes[,1], units = "days"))/365*12) # weeks

data$diff_in_month = datetimes$diff_in_month
data[,c(7,14)] = lapply(data[,c(7,14)], factor)


new_data = data
data[,c(4,5)] = new_data[,c(4,5)]
##################################################################################################################结束cleaning



###########imputation

impute_variables <- c('CustomerID','MaterialNumber','Monthlyunits', 'ContractStatus','BidSubmittedMonth',
                      'FailuretoSupplyRate.FTS.perunit','CustomerResponse')
mice_model <- mice(data[,impute_variables], method='rf')
completeData <- complete(mice_model,1)
data$FailuretoSupplyRate.FTS.perunit = completeData$FailuretoSupplyRate.FTS.perunit
###########imputation ends


#####################################Data Exploration

# correlation
sub_cor = data.frame(lapply(data[,c(1:5,11)],as.numeric))

correlationMatrix <- cor(sub_cor)
highlyCorrelated <- findCorrelation(correlationMatrix, cutoff=0.5)
print(highlyCorrelated)
library(corrplot)
corrplot(cor(sub_cor), order = "hclust")

sub = data[,c(1-10)]
mod = glm(CustomerResponse ~ ContractStatus + MaterialNumber + famAppearance +
            patron + MaterialAppearance + Monthlyunits* diff_in_month + FailuretoSupplyRate.FTS.perunit, data=sub,family=binomial(link='logit'))
logis = summary(mod)

## ANOVA
#customer significant diff
subdataCID = data %>% group_by(CustomerID) %>% mutate(res_cus = mean(CustomerResponse=="Won"))
subdataCID = subdataCID[,c(1,6,17)]
ano1 = aov(res_cus ~ CustomerID + ContractStatus, data =subdataCID)
summary(ano1)
#fam not sig
subdataCID = data %>% group_by(ProductFamily) %>% mutate(res_fam = mean(CustomerResponse=="Won"))
subdataCID = subdataCID[,c(3,17)]
ano2 = aov(res_fam ~  ProductFamily, data =subdataCID)
summary(ano2)
# material sig
subdataCID = data %>% group_by(MaterialNumber) %>% mutate(res_material = mean(CustomerResponse=="Won"))
subdataCID = subdataCID[,c(2,17)]
ano3 = aov(res_material ~  MaterialNumber, data =subdataCID)
summary(ano3)

#####################################

x = names(train[,])
x
myControl <- trainControl(method = "cv",
                          number = 10,
                          summaryFunction = twoClassSummary,
                          classProb = TRUE,
                          savePredictions = TRUE)

tgrid <- expand.grid(
  .mtry = c(2,5,10,20,30),
  .splitrule = c("gini"),
  .min.node.size = 5
)

# divide into training and validate set
set.seed(5)
tr <- sample(nrow(data), round(nrow(data) * 0.8))
train <- data[tr, ]
test <- data[-tr, ]
train = na.omit(train)
nrow(train)
nrow(test)
table(train$`CustomerResponse`)/nrow(train)  

##############KNN
knn_model <- train(CustomerResponse ~ ContractStatus * famAppearance * patron * Monthlyunits + diff_in_month * FailuretoSupplyRate.FTS.perunit,
                   data = train,
                   metric = "ROC",
                   method = "knn",
                   tuneLength = 20,
                   trControl = myControl)
print(knn_model)

############## Naive Bayesian 
nb_model <- train(CustomerResponse ~ ContractStatus * famAppearance * patron * Monthlyunits + diff_in_month * FailuretoSupplyRate.FTS.perunit,
                  data = train,
                  metric = "ROC",
                  method = "naive_bayes",
                  trControl = myControl)
plot(nb_model)
testclassi <- predict(nb_model, newdata = test)
confusionMatrix(data = testclassi, test$CustomerResponse)


##############RF
#rf_model2 <- train(train$`Customer Response` ~ .,train, method = "ranger")
# 无日期，无2017，+ Failure 73%  CustomerResponse ~ Monthlyunits+ContractStatus*famAppearance*patron + FailuretoSupplyRate.FTS.perunit
# 最高80% ： 无failure 全部 CustomerResponse ~ Monthlyunits+ContractStatus*famAppearance*patron + BidSubmittedMonth + BidSubmittedYear
# 78% : Gap+FTC - 日期 -2017 : CustomerResponse ~ Monthlyunits+ContractStatus*famAppearance*patron  + diff_in_month + FailuretoSupplyRate.FTS.perunit

rf_model1 =  train(CustomerResponse ~ Monthlyunits*ContractStatus + famAppearance*patron + diff_in_month + FailuretoSupplyRate.FTS.perunit * BidSubmittedYear,
                   data = train, method = "ranger",
                   trControl = myControl,
                   metric = "ROC",
                   tuneGrid = tgrid,
                   importance = "impurity")
  
testclassi1 <- predict(rf_model1, newdata = test)
confusionMatrix(data = testclassi1, test$CustomerResponse)


formula = CustomerResponse ~ ContractStatus * famAppearance * patron * Monthlyunits + diff_in_month * FailuretoSupplyRate.FTS.perunit


#87%
rf_model2 <- train(formula,
                   data = train, method = "ranger",
                   trControl = myControl,
                   metric = "ROC",
                   tuneGrid = tgrid,
                   importance = "impurity")
testclassi2 <- predict(rf_model2, newdata = test)
confusionMatrix(data = testclassi2, test$CustomerResponse)
pred_prob <- predict(rf_model2, test, type="prob")

plot(rf_model2, main="Roc Curve")
#ROC plot
ggplot(rf_model2) + theme_bw() + labs(title="Roc Curve")
print(rf_model2)

x <- evalm(rf_model2)
importantVar= varImp(rf_model2, scale = FALSE)

rownames(importantVar$importance) = gsub("ContractStatusN", "N", rownames(importantVar$importance))
rownames(importantVar$importance) = gsub("FailuretoSupplyRate.FTS.perunit", "FTS", rownames(importantVar$importance))

#important plot
plot(importantVar,main="Random Forest - Variable Importance",top=10)
importantVar$importance = importantVar$importance[1:5,]


pred_prob <- predict(rf_model2, test, type="prob")

# prob prediction
pred_prob <- predict(rf_model2, test, type="prob")
test[1:3,c(1,2,3,4,5,11,12,13,16)]
test$"Predicted Win Rate" = pred_prob$Won

################使low变高
ProbToWin = data.frame(CustomerID=test[c(1067, 462, 1295,735,641),"CustomerID"])
testlow = data.frame(test[c(1067, 462, 1295,735,641),])
testlow
pred_prob_low <- predict(rf_model2, testlow, type="prob")
pred_prob_low
ProbToWin$"Original Win Rate" = pred_prob_low$Won
#
testlow$Monthlyunits = mean(data$Monthlyunits)
testlow
pred_prob_low <- predict(rf_model2, testlow, type="prob")
pred_prob_low
ProbToWin$"Increase Monthly Unit to Average" = pred_prob_low$Won
#change interval to 10 month
testlow$patron = "500-600"
testlow
pred_prob_low <- predict(rf_model2, testlow, type="prob")
pred_prob_low
ProbToWin$"Increase Customer Frequency to 500-600" = pred_prob_low$Won
# diff in month
testlow$diff_in_month = 2
testlow
pred_prob_low <- predict(rf_model2, testlow, type="prob")
pred_prob_low
ProbToWin$"Increase Interval to 2 Month" = pred_prob_low$Won

ProbToWin$"Increase By(Factor)" = ProbToWin[,5]/ProbToWin[,2]
ProbToWin

ProbToWin_full = ProbToWin
# try plot a single tree
iris.tree = train(CustomerResponse ~ ContractStatus + patron + Monthlyunits + diff_in_month + FailuretoSupplyRate.FTS.perunit,
                  data = train,
                  method="rpart", 
                  trControl = trainControl(method = "cv"))

suppressMessages(library(rattle))

rattle::fancyRpartPlot(iris.tree$finalModel)

dim(ProbToWin_full[which(ProbToWin_full[,6]<1),])


dim(ProbToWin_full[which(ProbToWin_full[,6]<1 & ProbToWin_full[,2] < 0.3 ),])

dim(ProbToWin_full)


#**************************************

#75%
rf_model3 =  train(CustomerResponse ~ Monthlyunits*ContractStatus + famAppearance * MaterialAppearance + patron + diff_in_month + FailuretoSupplyRate.FTS.perunit * BidSubmittedYear,
                   data = train, method = "ranger",
                   trControl = myControl,
                   metric = "ROC",
                   tuneGrid = tgrid,
                   importance = "impurity")

testclassi3 <- predict(rf_model3, newdata = test)
confusionMatrix(data = testclassi3, test$CustomerResponse)












#
myControl2 <- trainControl(method = "cv",
                          number = 10,
                          classProb = TRUE,
                          savePredictions = TRUE)
rf_model4 <- train(CustomerResponse ~ ContractStatus * famAppearance * patron * Monthlyunits * MaterialAppearance + diff_in_month * FailuretoSupplyRate.FTS.perunit,
                   data = train, method = "ranger",
                   trControl = myControl2,
                   tuneGrid = tgrid,
                   importance = "impurity")
testclassi4 <- predict(rf_model4, newdata = test)
confusionMatrix(data = testclassi4, test$CustomerResponse)

ggplot(rf_model4)+ theme_bw() + labs(title="Accuracy Curve")
plot(varImp(object=rf_model4),main="RF - Variable Importance")

save.image("yoursession.RData")

####logist table

table = summary(mod)$coe
rownames(table) = gsub("famAppearance","Family",rownames(table))
rownames(table) = gsub("patron","Customer",rownames(table))
rownames(table) = gsub("MaterialAppearance","Material",rownames(table))
rownames(table) = gsub("diff_in_month","Time Interval",rownames(table))
rownames(table) = gsub("FailuretoSupplyRate.FTS.perunit","FTS",rownames(table))
rownames(table) = gsub("ContractStatus","",rownames(table))
table = table[,c(1,2,4)]
kable(table,caption = "Feature Selection")
xax = rownames(table)[3:8]
y = table[,1]
data.frame[xax=rownames(table)[3:8]]




rf_model2 <- train(formula,
                   data = train, method = "ranger",
                   trControl = myControl,
                   metric = "ROC",
                   tuneGrid = tgrid,
                   importance = "impurity")
pred_Win_Rate <- predict(rf_model2, newdata = test, type="prob") # Gives Predicted Win Rate
pred_classification <- predict(rf_model2, test, type="raw") # Gives Classification
prediction = test[,c(1,3,4)]
head(prediction)
prediction$"Win Rate" = pred_Win_Rate$Won
prediction$"Classification" = pred_classification
head(prediction)



actual = test[,c(1,10)]
actual$prediction = pred_classification
View(actual)
actual[1241:1251,]
