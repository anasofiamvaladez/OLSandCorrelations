#This script aims to choose the best indicators 
#that explain movements on the demand of credit for Big enterprises

# Magdiel Urbina Solalinde & Ana Sofía Muñoz Valadez 25.2.2019

# Data is from main database assembled by Manuel Alejandro Bustos

# Ammends:
# 1. Polishing notation, removing unused objects
# 2. Use y as dep variable*

#----------------------------------
# (0) Clear the environment and call the packages: 
#-----------------------------------

#install.packages("XLConnect")
#install.packages("XLConnectJars")


library(XLConnect) 
library(xlsx)
library(stringr)
library(ggplot2)
#install.packages("stringr") 
#install.packages("devtools")
library(devtools)
#install_github("arnejohannesholmin/cpplot3d")
library(arnejohannesholmin/cpplot3d)



rm(list = ls()) #Clears variable environment
cat("\014")     #Clears console

setwd("H:/Proyectos Especiales/Proyectos/GAM/EnBan/Automatizacion Factores Junio 2019/Modulo 1/Demanda");
direccion <- "H:/Proyectos Especiales/Proyectos/GAM/EnBan/Automatizacion Factores Junio 2019/Modulo 1/Demanda"

# The database contains quarterly data to compute correlations between:
# credit demand; explicative variables;

#-----------------------------------
# (1) Prepare data
#-----------------------------------

data <- read.xlsx("data_factors_octubre.xlsx", sheetName = "factores (2)", header = T)
data1 <- read.xlsx("data_factors_octubre.xlsx", sheetName = "factores (3)", header = T)
date_obs <- data$date; # Dates for observed data

#correlation matrix
columns <- ncol(data)
output <- matrix(0, nrow= columns, ncol=columns)
output <- cor(data)

# Compute variables
observations <- nrow(data)

# Define number of combinations & way of combining

nc <- 5
operation <- "sum"
factor.list <- list(matrix(0, 16, 1), matrix(0, 16, 1), matrix(0, 16, 1),matrix(0, 16, 1), matrix(0, 16, 1))


#*************************************
#For every factor elected
#*************************************

###Function for number of combinations

#combination_function<-function(nc, operation){
  
######Create space######

factors <- c(colnames(data))
matrix_combinations <- combn(factors, nc)  
n <- ncol(matrix_combinations)
results <- as.data.frame (matrix(0, nrow= 7+5, ncol=n))


 for (i in 1:n) {
   for (j in 1:nc){
     
    
# j es el renglon de la matriz de combinaciones
    
     numberfactor <- matrix_combinations[j,i]
     names(numberfactor) <- paste("factor", j, sep = "")
     x <- data[numberfactor]
     factor.list[[j]] <- x
     
   }
 
   

df<-data.frame(y=data1[,2],x1=as.vector(as.matrix(as.matrix(factor.list[[1]]))),x2=as.vector(as.matrix(as.matrix(factor.list[[2]]))), x3=as.vector(as.matrix(as.matrix(factor.list[[3]]))), x4=as.vector(as.matrix(as.matrix(factor.list[[4]]))), x5=as.vector(as.matrix(as.matrix(factor.list[[5]]))))
equation <- lm(y~.,df)
fill <-summary(equation)$adj.r.squared       
fill2<-summary(equation)$coefficients[1,1]        
fill3<-summary(equation)$coefficients[2,1]  
fill4<-summary(equation)$coefficients[3,1]  
fill5<-summary(equation)$coefficients[4,1]   
fill6<-summary(equation)$coefficients[5,1] 
fill7<-summary(equation)$coefficients[6,1] 
fill8<- paste(colnames(factor.list[[1]]))
fill9<- paste(colnames(factor.list[[2]]))
fill10<- paste(colnames(factor.list[[3]]))   
fill11<- paste(colnames(factor.list[[4]]))
fill12<- paste(colnames(factor.list[[5]]))
results[1,i] <- fill
results[2,i] <- fill2
results[3,i] <- fill3
results[4,i] <- fill4
results[5,i] <- fill5
results[6,i] <- fill6
results[7,i] <- fill7
results[8,i] <- fill8
results[9,i] <- fill9
results[10,i] <- fill10
results[11,i] <- fill11
results[12,i] <- fill12

colnames(results)[i]<- paste(colnames(factor.list[[1]]),colnames(factor.list[[2]]),colnames(factor.list[[3]]),colnames(factor.list[[4]]), colnames(factor.list[[5]]), sep="&")

 }
 
results <- t(results)

#PAra exportar las combinaciones
setwd("H:/Proyectos Especiales/Proyectos/GAM/EnBan/Automatizacion Factores Junio 2019/Modulo 1/Demanda");
direccion <- "H:/Proyectos Especiales/Proyectos/GAM/EnBan/Automatizacion Factores Junio 2019/Modulo 1/Demanda"
write.csv(results, file="resultsOLS3.csv")

