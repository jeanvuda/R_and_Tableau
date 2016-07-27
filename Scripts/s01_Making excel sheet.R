# 1) Load Functions
source("C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\R\\Scripts\\s01_Functions.R")

# 2) Bring in data
# All tabs from excel file is saved to a list
wb<-loadWorkbook("C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Tables.xlsx")
lst<- readWorksheet(wb, sheet = getSheets(wb))
rm(wb)

# 3) 
names(lst)
head(lst[[1]])
lst<-lapply(lst, function(x) {x[!duplicated(x),]})


for(j in c(1,2,5,7,1)){
  print(j)
  lst[[j]]$var<-paste0("[", lst[[j]]$Name, "]")
  }




tmp<-lapply(lst[[4]]$Formula, function(x) { #print(x)})
  temp<-x
  for(j in c(1,2,5,7,1)){
    for(i in 1:length(lst[[j]]$Variable)){
      temp<-gsub(lst[[j]]$Variable[i], lst[[j]]$var[i], temp, fixed=T)
    }
  }
  temp
})

lst[[4]]$nf<-do.call(rbind, tmp)

tes<-rbind(lst[[1]][,c(1,2)],lst[[2]][,c(1,2)], lst[[4]][,c(1,2)], lst[[6]])
lst[[3]]<-merge(lst[[3]], tes, by="Variable", all.x=T, all.y = F)


tofile<-"C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Tables_raw.xlsx"
writeWorksheetToFile(tofile, data = lst, sheet=names(lst))


