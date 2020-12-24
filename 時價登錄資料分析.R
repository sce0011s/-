options(scipen=4)
library(xlsx)
library(openxlsx)
library(readxl)
library(dplyr)
library(stringr)
library(lubridate)
library(car)
library(GGally)
library(rpart)
library(rpart.plot)
library(randomForest)
library(e1071)
library(MASS)
library(ggplot2)
library(reshape2)
library(car)
library(nlme)
library(partykit)
library(caret)
# library(foreach)
# library(doParallel)
library(lmtest)
library(parallel)
library(snow)
# #DATA STEP
# filename <- list.files("C:\\Users\\steven\\Desktop\\時價登錄資料")
# filepos <- paste0("C:\\Users\\steven\\Desktop\\時價登錄資料\\",filename)
# name <- c("town","objtran","houseno","landtranarea","urbanlandz","nurbanlandz","nurbanlandc","trandate","transactions","shiftedarea","floors","buildingtype","mainused","buildingmaterials","buildingcomdate","buildingtransferarea","room","lroom","broom","partition","communityservice","price","unitprice","parkingspacetype","parkingspacearea","parkingspaceprice","remark","no")
# org <- list()
# for(i in 1:length(filepos)){
#   org[[i]] <- read_xls(filepos[i],sheet = 3,skip = 1)
#   if(i==8){org[[i]] <- org[[i]][,1:28]}else{org[[i]] = org[[i]]}
#   names(org[[i]]) <- name
# }
# data <- do.call(rbind,org) %>% as.data.frame()
# openxlsx::write.xlsx(data,file = "C:\\Users\\steven\\Desktop\\時價登錄資料\\data.xlsx")
# 
data <- read.xlsx("C:\\Users\\steven\\Desktop\\時價登錄資料\\data.xlsx") #讀取資料

data[,3] <- str_replace(data[,3],paste0("臺北市",data[,1]),"")
data[,3] <- str_replace(data[,3],paste0("台北市",data[,1]),"")
data[,1] <- factor(data[,1],levels = c("士林區","大同區","大安區","中山區","中正區","內湖區","文山區","北投區","松山區","信義區","南港區","萬華區"),labels = c(1:12))
data[,2] <- factor(data[,2],levels = c("土地","車位","房地(土地+建物)","房地(土地+建物)+車位","建物"),labels=c(1:5))
data[,3] <- as.factor(data[,3])
data[,5] <- factor(data[,5],levels = c("工","住","商","農","其他"),labels=c(1:5))
data[,12] <- factor(data[,12],levels = c("工廠","公寓(5樓含以下無電梯)","住宅大樓(11層含以上有電梯)","店面(店鋪)","倉庫","套房(1房1廳1衛)","透天厝","華廈(10層含以下有電梯)","農舍","廠辦","辦公商業大樓","其他"),labels=c(1:12))
data[,13] <- ifelse(is.na(data[,13]),"其他",data[,13])
data[,13] <- factor(data[,13],levels = c("工商用","工業用","住工用","住家用","住商用","停車空間","商業用","國民住宅","農舍","見使用執照","見其他登記事項","其他"),labels=c(1:10,10,10))
data[,14] <- ifelse(is.na(data[,14]),"其他",data[,14])
data[,14] <- factor(data[,14],levels = c("土木造","土造","土磚石混合造","木造","加強磚造","石造","混凝土造","預力混凝土造","壁式預鑄鋼筋混凝土造","磚造","鋼骨混凝土造","鋼骨鋼筋混凝土造","鋼造","鋼筋混凝土加強磚造","鋼筋混凝土造","見使用執照","見其他登記事項","其他"),labels=c(1:16,16,16))
data[,20] <- factor(data[,20],levels = c("無","有"),labels=c(0,1))
data[,21] <- factor(data[,21],levels = c("無","有"),labels=c(0,1))
data[,24] <- ifelse(is.na(data[,24]),"無車位",data[,24])
data[,24] <- factor(data[,24],levels = c("無車位","一樓平面","升降平面","升降機械","坡道平面","坡道機械","塔式車位","其他"),labels=c(0:7))

data$is_top_floor <- str_detect(data$floors,data$shiftedarea)
data$is_top_floor <- ifelse(is.na(data$is_top_floor),FALSE,data$is_top_floor)
data$is_top_floor <- factor(data$is_top_floor,levels=c(TRUE,FALSE),labels = c(1,0)) #是否為頂樓
#buildingcomdate用EXCEL手動清除不符合規格資料
#計算屋齡
data$age <- ifelse(nchar(data$trandate)==7 & nchar(data$buildingcomdate)==7 & !is.na(data$buildingcomdate),as.numeric(str_sub(data$trandate,1,3))-as.numeric(str_sub(data$buildingcomdate,1,3)),
                   ifelse(nchar(data$trandate)==6 & nchar(data$buildingcomdate)==7 & !is.na(data$buildingcomdate),as.numeric(str_sub(data$trandate,1,2))-as.numeric(str_sub(data$buildingcomdate,1,3)),NA))

data$land <- str_sub(str_extract(data$transactions,"土地\\d{1,}"),3) %>% as.numeric()
data$build <- str_sub(str_extract(data$transactions,"建物\\d{1,}"),3) %>% as.numeric()
data$parking <- str_sub(str_extract(data$transactions,"車位\\d{1,}"),3) %>% as.numeric()
data$parking <- ifelse(data[,24] %in% c(1:7) & data$parking==0,1,data[,24])

data$room <- as.numeric(data$room)
data$lroom <- as.numeric(data$lroom)
data$broom <- as.numeric(data$broom)
#age 遺失值填補:
avg_age <- mean(data$age , na.rm = T) %>% trunc() #去除小數位
data$age <- ifelse(data$age<0 & !is.na(data$age),NA,data$age) #負值改為NA
group_age <- data[!is.na(data$age),c(3,30)] %>% group_by(houseno) %>% summarise(avg_age=trunc(mean(age))) #以同路段屋齡平均填補
data <- left_join(data,group_age,"houseno")
data$age <- ifelse(is.na(data$age),data$avg_age,data$age) 
data <- data[,-34]
data$age <- ifelse(is.na(data$age),avg_age,data$age)
# ################
#選取分析用資料
data_analysis <- data[(str_detect(data$remark,"親友")==0 | is.na(str_detect(data$remark,"親友"))) & data$objtran %in% c(3,4,5) & data$buildingtype %in% c(2,3,6,7,8,12),c(1,2,4,12,13,14,16,17,18,19,20,21,23,24,25,29,30,31,32,33)] #篩選資料及變數
data_analysis <- data_analysis[data_analysis$unitprice!=0,]
summary(data_analysis)
# ggcorr(data = data_analysis,palette = "RdYlGn",label = TRUE,label_color = "black") #相關係數圖
data_analysis <- data_analysis[(data_analysis$unitprice %in% unique(boxplot(data_analysis$unitprice,plot = F)$out))==0,] #去除price離群值資料(小於Q1-1.5IQR或大於Q3+1.5IQR)
# boxplot(data_analysis$unitprice)$stats
# openxlsx::write.xlsx(data,file = "C:\\Users\\steven\\Desktop\\時價登錄資料\\data_1.xlsx")
set.seed(830527)
index<-sample(1:nrow(data_analysis),size=trunc(nrow(data_analysis)*0.2))
training_data <- data_analysis[-index,] #訓練資料
test_data <- data_analysis[index,] #測試資料

#RMSE function
########################################################################################################################################
rmse <- function(x){
  predict <- sqrt(sum((test_data$unitprice-predict(x,test_data))^2)/length(test_data$unitprice))
  training <- sqrt(sum((training_data$unitprice-predict(x,training_data))^2)/length(training_data$unitprice))
  return(c("predict"=predict,"training"=training))
}
#迴歸模型
########################################################################################################################################
lm_model <- lm(unitprice~.,training_data)
qqnorm(lm_model$residuals);qqline(lm_model$residuals)
summary(lm_model)
# test_data$lm_price <- predict(lm_model,test_data)
rmse(lm_model) #test:38873.11,train:39091.54
t <- boxcox(lm_model)$x[which.max(boxcox(lm_model)$y)]
boxcox_lm_model <- lm((unitprice^t-1)/t~.,training_data)
# ncvTest(boxcox_lm_model) #同質性
# dwtest(boxcox_lm_model) #獨立性(不滿足)
qqnorm(boxcox_lm_model$residuals);qqline(boxcox_lm_model$residuals)
summary(boxcox_lm_model)
sqrt(sum((test_data$unitprice-(predict(boxcox_lm_model,test_data)*t+1)^(1/t))^2)/length(test_data$unitprice)) #38762.98
sqrt(sum((training_data$unitprice-(predict(boxcox_lm_model,training_data)*t+1)^(1/t))^2)/length(training_data$unitprice)) #38998.77
#######################################################################################################################################
#rpart迴歸樹
# tree_model <- rpart(unitprice~.,data = training_data,method = "anova",cp=0.0001)
# tree_model$cptable[which.min(tree_model$cptable[,"xerror"]),"CP"]  #cp=0.0001670665
# tree_model_2 <- rpart(unitprice~.,data = training_data,method = "anova",cp=0.0001670665)
# rmse(tree_model)
# rmse(tree_model_2)
cpu.cores <- detectCores()
cl = makeCluster(cpu.cores,type='SOCK')
clusterExport(cl,c("rpart","training_data","model"))
tree_model = parSapply(cl,X=seq(0.1,0.001,-0.001),FUN=function(cp){list(rpart(unitprice~.,data = training_data,method = "anova",cp=cp))})
stopCluster(cl)

# tree_model1 <- rpart(unitprice~.,data = training_data,method = "anova",cp=0)
# rpart.plot(tree_model1)
# tree_model2 <- rpart(unitprice~.,data = training_data,method = "anova",cp=0.01)
# rpart.plot(tree_model2)
# tree_model3 <- rpart(unitprice~.,data = training_data,method = "anova",cp=0.1)
# rpart.plot(tree_model3)

ggplot()+
  geom_line(data=tree_rmse,aes(x=CP,y=predict,colour="predict"),size=1)+
  geom_line(data=tree_rmse,aes(x=CP,y=training,colour="training"),size=1)+
  scale_x_reverse()+
  scale_color_manual(name="",values=c("predict"="red","training"="blue"))+
  scale_y_continuous("rmse")
#######################################################################################################################################
#randomforest
# forest_model <- randomForest(unitprice~.,data=training_data) #預設mtry=p/3=6 ntree=150
# tuneRF(test_data[,-13],test_data[,13]) #選擇最佳mtry數目6(隨機抽取的變數數量)，類似python中max_features
# plot(forest_model) #選擇ntree最佳數目(樹的數量) 約150，類似python中n_estimators
# forest_mse_training <- training_mse(forest_model)
# forest_mse_predict <- predict_mse(forest_model)

cpu.cores <- detectCores()
cl = makeCluster(cpu.cores,type='SOCK')
clusterExport(cl,c("randomForest","training_data","rmse","test_data"))
forest_model <- parSapply(cl,X=seq(10,500,10),FUN=function(ntree){list(randomForest(unitprice~.,data=training_data,ntree=ntree))})
forest_rmse <- parSapply(cl,forest_model,FUN=function(x){list(rmse(x))})
stopCluster(cl)
forest_rmse <- do.call(rbind,forest_rmse) %>% as.data.frame()
forest_rmse$ntree <- seq(10,500,10)

ggplot()+
  geom_line(data=forest_rmse,aes(x=ntree,y=predict,colour="predict"),size=1)+
  geom_line(data=forest_rmse,aes(x=ntree,y=training,colour="training"),size=1)+
  scale_color_manual(name="",values=c("predict"="red","training"="blue"))+
  scale_y_continuous("rmse")

#ntree>100時rmse收斂，取ntree=100

cpu.cores <- detectCores()
cl = makeCluster(cpu.cores,type='SOCK')
clusterExport(cl,c("randomForest","training_data","rmse","test_data"))
forest_model <- parSapply(cl,X=seq(4,15,1),FUN=function(mtry){list(randomForest(unitprice~.,data=training_data,ntree=100,mtry=mtry))})
forest_rmse <- parSapply(cl,forest_model,FUN=function(x){list(rmse(x))})
stopCluster(cl)
forest_rmse <- do.call(rbind,forest_rmse) %>% as.data.frame()
forest_rmse$mtry <- seq(4,15,1)

ggplot()+
  geom_line(data=forest_rmse,aes(x=mtry,y=predict,colour="predict"),size=1)+
  geom_line(data=forest_rmse,aes(x=mtry,y=training,colour="training"),size=1)+
  scale_color_manual(name="",values=c("predict"="red","training"="blue"))+
  scale_y_continuous("rmse")

#mtry=7時rmse最小，取mtry=7

forest_model_final <- randomForest(unitprice~.,data=training_data,ntree=100,mtry=7)
rmse(forest_model_final) #test:29923.38,train:14937.33 
#######################################################################################################################################
#svm(cost, epsilon)
#cost:越大則容錯度越少,也越少support vectors,越容易overfitting,預設1
#epsilon:越大則support vectors數量越少,預設0.1
#gamma:越大越容易overfitting
# svm_model <- svm(unitprice~.,type="eps-regression",data=training_data,cost=0.01) #45586.79  45529.98
# # tune.svm_model <- tune(svm,unitprice~.,data=data_analysis,ranges = list(cost=seq(0.1,10,0.1),gamma=seq(0.01,0.1,0.01)))
# svm_model2 <- svm(unitprice~.,type="eps-regression",data=training_data,cost=0.1) #38628.72  38639.47
# plot(svm_model2,data=training_data,color.palette = topo.colors)

cpu.cores <- detectCores()
cl = makeCluster(cpu.cores,type='SOCK')
clusterExport(cl,c("svm","training_data","rmse","test_data"))
svm_model <- parSapply(cl,X=c(2^-5,2^-1,2^3,2^7,2^11,2^15),FUN=function(cost){list(svm(unitprice~.,data=training_data,cost=cost))})
svm_rmse <- parSapply(cl,svm_model,FUN=function(x){list(rmse(x))})
stopCluster(cl)
svm_rmse <- do.call(rbind,svm_rmse) %>% as.data.frame()
svm_rmse$cost <- c(2^-5,2^-1,2^3,2^7,2^11,2^15)

ggplot()+
  geom_line(data=svm_rmse,aes(x=cost,y=predict,colour="predict"),size=1)+
  geom_line(data=svm_rmse,aes(x=cost,y=training,colour="training"),size=1)+
  scale_color_manual(name="",values=c("predict"="red","training"="blue"))+
  scale_y_continuous("rmse")

#cost=2^7=128時對測試資料有最小rmse34413.14，取cost=128

cpu.cores <- detectCores()
cl = makeCluster(cpu.cores,type='SOCK')
clusterExport(cl,c("svm","training_data","rmse","test_data"))
svm_model <- parSapply(cl,X=c(0.01,0.1,0.5,1),FUN=function(epsilon){list(svm(unitprice~.,data=training_data,cost=128,epsilon=epsilon))})
svm_rmse <- parSapply(cl,svm_model,FUN=function(x){list(rmse(x))})
stopCluster(cl)
svm_rmse <- do.call(rbind,svm_rmse) %>% as.data.frame()
svm_rmse$epsilon <- c(0.01,0.1,0.5,1)

ggplot()+
  geom_line(data=svm_rmse,aes(x=epsilon,y=predict,colour="predict"),size=1)+
  geom_line(data=svm_rmse,aes(x=epsilon,y=training,colour="training"),size=1)+
  scale_color_manual(name="",values=c("predict"="red","training"="blue"))+
  scale_y_continuous("rmse")

#epsilon=0.5時對測試資料有最小rmse，取epsilon=0.5

svm_model_final <- svm(unitprice~.,data=training_data,cost=128,epsilon=0.5)

