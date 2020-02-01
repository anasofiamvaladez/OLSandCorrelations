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
direccion <- "H:/Proyectos Especiales/Proyectos/GAM/EnBan/Automatizacion Junio 2019/Modulo 1/Demanda"

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
results <- as.data.frame (matrix(0, nrow= observations, ncol=n))


 for (i in 1:n) {
   for (j in 1:nc){
     
    
# j es el renglon de la matriz de combinaciones
    
     numberfactor <- matrix_combinations[j,i]
     names(numberfactor) <- paste("factor", j, sep = "")
     x <- data[numberfactor]
     factor.list[[j]] <- x
     
   }
 
   
 
if(operation=="sum"){
  
  fill <- as.vector(as.matrix(as.matrix(factor.list[[1]])+as.matrix(factor.list[[2]])+as.matrix(factor.list[[3]])+ as.matrix(factor.list[[4]])+ as.matrix(factor.list[[5]])))} else {
  
  fill <- as.vector(as.matrix(as.matrix(factor.list[[1]])+as.matrix(factor.list[[2]])+as.matrix(factor.list[[3]])+ as.matrix(factor.list[[4]])+ as.matrix(factor.list[[5]]))/j)
} 
  
results[,i] <- fill
colnames(results)[i]<- paste(colnames(factor.list[[1]]),colnames(factor.list[[2]]),colnames(factor.list[[3]]),colnames(factor.list[[4]]), colnames(factor.list[[5]]), sep="&")

 }
 

results
#PAra exportar las combinaciones
OutputAdress<- paste(direccion)
OutputName <- "results.xlsx"
ruta = paste(OutputAdress,OutputName,sep="\\")
#write.xlsx(results,file=ruta,sheet =  paste("Comb Dem con 31 en", nc, sep="_"),append=FALSE)

#}


#combination_function(2, operation)


###remove everything except##############


rm(list=setdiff(ls(), c("results", "data1")))


################# Correlation one-by-one variable#############


columns <- ncol(results)
output <- as.data.frame(matrix(0, nrow= columns, ncol=1))
for (i in 1:columns) {
  y <- data1[,2]
  x <- results[,i]
  output[i,] <- cor(y, x, method="pearson")
  rownames(output)[i]<- paste(colnames(results)[i])
  print(i)
  
}


colnames(output)[1]<-"Correlations"

########ORDEN DESCENDENTE##################

rm(list=setdiff(ls(), "output"))
output <- cbind(output, sampleids=rownames(output))
newdata <-output[order(output$Correlations, output$sampleids, decreasing = TRUE), ]



##########OUTPUT##########################
#Exportamos a excel las correlaciones de la demanda con las combinaciones de 31 en 2
#output
setwd("H:/Proyectos Especiales/Proyectos/GAM/EnBan/Automatizacion Factores Junio 2019/Modulo 1/Demanda");
direccion <- "H:/Proyectos Especiales/Proyectos/GAM/EnBan/Automatizacion Factores Junio 2019/Modulo 1/Demanda"
write.csv(newdata, file="results_trial5_junio19.csv")
#write.xlsx(output,file=ruta,sheet = paste("Corr Dem con 31 en", nc, sep="_"),append=TRUE)















#######Choose Plots (5)##############################
newdata <- read.xlsx("results_trial3.xlsx", sheetName = "results_trial3", header = T)
newdata <- newdata[1:10,]
factors <- newdata$sampleids
factors <- str_split_fixed(newdata$sampleids, "&", 5)

Grupo1 <- factors[1,]
Grupo2 <- factors[2,]
Grupo3 <- factors[3,]
Grupo4 <- factors[4,]
Grupo5 <- factors[5,]
Grupo6 <- factors[6,]
Grupo7 <- factors[7,]
Grupo8 <- factors[8,]
Grupo9 <- factors[9,]
Grupo10 <- factors[10,]


ExpandedArray <- array(0, c(16,5,10)) 

for (j in 1:10){

newdata <- as.data.frame(get(paste0("Grupo", j)))
matrixplot <- matrix(0, 16, 5)

   for (i in 1:5){
         
         factor <-newdata[i,]
         names(factor) <- paste0("factor", i)
         fill <- as.vector(data[,factor])
         ExpandedArray[,i, j] <- fill  
     
   }
  
  plot <- as.matrix(t(ExpandedArray[,, j]))
  colnames(plot)  <- c("201501","201502", "201503", "201504", "201601","201602", "201603", "201604","201701","201702", "201703", "201704", "201801","201802", "201803", "201804") 
  grafica <- barplot(plot, col = c("darkgoldenrod1","darkorange", "azure3", "darkolivegreen3", "darkred"),  
  ylim = c(-.8, 1), legend =get(paste0("Grupo", j)), main= "Factores Relacionados con la Demanda")
  lines(x = grafica, y = data1$y_t)
  points(x = grafica, y = data1$y_t)
  
pdf(file=paste0("grafica", j, ".pdf"),width=10,height=7)

}

dev.off()
          


#######Plot##############################

corrplot <- ggplot() +     
  geom_bar(data = newdata, aes(x=sampleids, y=Correlations),stat = "identity") +        
  colScale+
  labs(x="Factores", y="Correlacion") +        
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  theme(legend.position="bottom", legend.direction="horizontal",
        legend.title = element_blank(),legend.justification=c(-1,0))+
  theme(axis.line.x = element_line(color="black", size = .5),
        axis.line.y = element_line(color="black", size = .5))+
  theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank(),
        panel.background = element_blank()) +
  theme(plot.title = element_text(size = 11, face = "bold"),            
        axis.text.x=element_text(colour="black", size = 11),
        axis.text.y=element_text(colour="black", size = 11))+
  scale_y_continuous()+guides(fill=guide_legend(ncol=6,nrow=3))

corrplot



##########OUTPUT##########################
#Exportamos a excel las correlaciones de la demanda con las combinaciones de 31 en 2
#output
setwd("H:/Proyectos Especiales/Proyectos/GAM/EnBan/Automatizacion Factores Junio 2019/Modulo 1/Demanda");
direccion <- "H:/Proyectos Especiales/Proyectos/GAM/EnBan/Automatizacion Factores Junio 2019/Modulo 1/Demanda"
write.csv(newdata, file="results_trial5_junio.csv")
#write.xlsx(output,file=ruta,sheet = paste("Corr Dem con 31 en", nc, sep="_"),append=TRUE)