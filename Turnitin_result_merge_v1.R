
rm(list=ls())

#!!!!!!!!!!!!!!!#
#!!!! INPUT !!!!#
#!!!!!!!!!!!!!!!#

rubrics  = 30   #Vul hier het aantal RUBRICS Excel bestanden in.
semester =  2   #Vul hier het semester in waar de data over gaat.

# Geef elk bestand dat gemerged moet worden de naam van een getal, beginnend met 1 en daarna oplopend.
#!!! Plaats de excel bestanden op de Q:/schijf in onderstaande map:
# Psychologie/Onderwijsinstituut/Onderwijsbalie Psychologie/ICT en Onderwijs/Projecten Paul/Analyses/Turnitin Jos/Input/
# De output zal verschijnen in de map output

#!!!!!!!!!!!!!!!#
#!! END INPUT !!#
#!!!!!!!!!!!!!!!#



##################
# PREPAIR R-FILE #
##################

if("gdata" %in% rownames(installed.packages()) == FALSE) {install.packages("gdata")}
if("reshape" %in% rownames(installed.packages()) == FALSE) {install.packages("reshape")}
if("plyr" %in% rownames(installed.packages()) == FALSE) {install.packages("plyr")}
library(gdata)
library(reshape)
library(plyr)

if(Sys.info()["sysname"]=="Windows"){
setwd("C:/Users/paullodder/SURFdrive/ICT en Onderwijs/Turnitin/Turnitin R project/Input")}
if(Sys.info()["sysname"]=="Darwin"){
  setwd("/Users/paullodder/SURFdrive/ICT en Onderwijs/Turnitin/Turnitin R project/Input")}

####################
# RUBRIC DATAFILES #
####################

rdata=list()
students=c()

# Create dataframe for each group
for(i in (1+100*semester):(100*semester+rubrics)){
    nam <- paste("r", i, sep = "")
    assign(nam, read.xls(gsub(" ","",paste("R",toString(i),".xls")),pattern="author"))
    students[i-100*semester]<-dim(read.xls(gsub(" ","",paste("R",toString(i),".xls")),pattern="author"))[1]
}

# Combine dataframes in a list
for(j in 1:rubrics){
    rdata[[j]]<-get(ls(pattern = "^r...$")[j])
}

# Merge dataframes in list to a single dataframe
rdata<-do.call("rbind", rdata)

# Add assignment and group column
groep=c()
low=1
high=students[1]
for(i in 1:rubrics){
   groep[low:high]<-i
   low<-low+students[i]
   high<-high+students[i+1]    # I+1 veranderd ipv i
}

opdracht=rep(gsub("^.*?VRT","VRT",names(read.xls(gsub(" ","",paste("R",toString(semester),"01",".xls")),pattern="VRT"))[1]),sum(students))

rdata<-cbind(opdracht,groep,rdata)


#######################
# QUICKMARK DATAFILES #
#######################

qdata=list()
studentsq=c()

# Create dataframe for each group
for(i in (1+100*semester):(100*semester+rubrics)){
  if(file.exists(gsub("Jos ","Jos",paste(getwd(),gsub(" ","",paste("/",toString(201),".xls")))))==TRUE){
    nam <- paste("q", i, sep = "")
    assign(nam, read.xls(gsub(" ","",paste(toString(i),".xls")),pattern="author"))
    studentsq[i-100*semester]<-dim(read.xls(gsub(" ","",paste(toString(i),".xls")),pattern="author"))[1]
  }
}

# Combine dataframes in a list
for(j in 1:rubrics){
  if(is.na(studentsq[j])==TRUE){print(paste(j,"=false"))}
  else{
    qdata[[j]]<-get(ls(pattern = "^q...$")[j])
  }
}

# Merge dataframes in list to a single dataframe
qdata<-do.call("rbind.fill", qdata)


###############
# Final steps #
###############

# Remove irrelevant rows
rdata<-rdata[-which(rdata[,"author"]=="(* rubric was not scored)"),]

# Sort columns in dataframe
sortq <- qdata[,c(names(qdata)[1:2],sort(names(qdata)[3:dim(qdata)[2]]))]  

# Merge quickmark with rubric data
finaldata=merge(rdata,sortq,by="author",all.x=TRUE)

# Change title names
colnames(finaldata)[which(names(finaldata)=="title.x")]<-c("Rubrics onderdeel")
colnames(finaldata)[which(names(finaldata)=="title.y")]<-c("Quickmarks onderdeel")

# Write data to .csv file
write.csv(finaldata,"VRT5EERSTE.csv")


###############
### The End ###
###############
















