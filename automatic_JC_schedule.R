#script to make an automated table filling for JC
source("~/Desktop/scripts_isilon/R/commonFunctions.R")
library(openxlsx)
peopleList<-readLines("~/Documents/Stark_oficial/people4JC.list")
peopleList<-tolower(peopleList)
currentSchedule<-read_excel_allsheets("~/Documents/Stark_oficial/JC_automatic_fill.xlsx")
currentSchedule<-currentSchedule$`Sheet 1`
currentSchedule$date1=as.Date(currentSchedule$date1)
currentSchedule$date_JC=as.Date(currentSchedule$date_JC)
currentSchedule$nameGM<-tolower(currentSchedule$nameGM)

#fulfill JC dates 
currentSchedule[is.na(currentSchedule$date_JC),'date_JC']=currentSchedule[is.na(currentSchedule$date_JC),'date1']+2

#for each date JC make a random sample of three person, eliminated those one 
index2start=min(which(is.na(currentSchedule$name1)), which(is.na(currentSchedule$name2)),
    which(is.na(currentSchedule$name3)))
for(i in seq(index2start, nrow(currentSchedule), by=1)){
  namesEliminate=c(as.character(unlist(strsplit(currentSchedule[i, 'nameGM'],"/"))),
                   as.character(unlist(as.list(currentSchedule[(i-3):(i),c('name1','name2','name3')]))),
                   as.character(gsub("\\S+\\s+(.+)\\s+Monday seminar","\\1", currentSchedule[i, 'other']))
                   ) %>% tolower() %>% unique() %>% na.omit() %>% as.character()
  
  l=c('name1','name2','name3')
  for(j in l){
    if(is.na(currentSchedule[i, j])){
      res.name=sample(setdiff(peopleList, namesEliminate),1)
      namesEliminate=c(namesEliminate, res.name)
      currentSchedule[i,j]=res.name
    }
  }
}

write.xlsx(currentSchedule, file="~/Documents/Stark_oficial/JC_automatic_fill.xlsx")

#check that every person has appeared more or less equal number of times 
#allnames= as.character(unlist(as.list(currentSchedule[,c('name1','name2','name3')]))) %>% tolower() %>% table()
