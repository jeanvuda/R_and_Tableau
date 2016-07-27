# Solution

# Get functions
source('~/01_Okaki/02_CPSA/Tableau Tabular/R/Scripts/s01_Functions.R')

# Get the file and process it
rootnode<-make_rootnodes(wb='C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Test\\wb1\\wb20160713.twb')

# Generate list of parameters
ca<-data.table::data.table(get_vars(proc=rootnode, parameter=TRUE)[1])

# Generate variables from all datasources
cb<-get_vars(proc=rootnode, parameter=FALSE)

# Cleaning formula
cb<-lapply(cb, function(x){
  ccc<-x
  ccd<-ccc[, {te=as.character(ccc[,formula]);
  for(i in 1:nrow(ccc)){te<-gsub(ccc[i,name], ccc[i,caption], te, fixed = TRUE)};
  list(nformula=te)}]
  ccd<-cbind(ccd,ccc)
})

# Generating folder group
ce<-get_folder()
data.table::setnames(ce, "name", "folder.name")

#Extract from excel sheet
library(XLConnect)
wb<-loadWorkbook("C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Tables_original_714.xlsx")
lst<-readWorksheet(wb, sheet = getSheets(wb))

# For reporting measures/dimensions
# By datasource / mesaure&dimension
# DataSource, role, Folder, Caption, datatype,  and formula if there is any

# get the list of Calculated fields
cc<-merge(ce, cb[[2]], by.x="var", by.y="name", all=TRUE)
cc[]<-lapply(cc, as.character)
cc<-merge(cc, lst[[4]], all.x=T, by.x=c("var"), by.y=c("Variable"))

cc<-lapply(c("dimensions", "measures"), function(x){
cc[role.x==x & !is.na(class),.(DS, role.y, folder.name.x, caption, datatype, var, nformula, Checked, Definition, Duplicate, Suggestion)]})


# get the list of data variables from the dbase 
cd<-merge(ce, cb[[2]], by.x="var", by.y="name", all=TRUE)
cd<-lapply(c("dimension", "measure"), function(x){
  cd[role.y==x & is.na(class),.(DS, role.y, folder.name.x, caption, datatype, var)]})


# For reporting parameters
# Role, Folder, Caption, datatype, etc
ca<-get_vars(proc=rootnode, parameter=TRUE)[[1]]
ca<-merge(ce, ca, by.x="var", by.y="name", all.y=TRUE)
ca[,default.format:=NULL]
ca[,role.y:=NULL]

Calculated.fields<-data.table::rbindlist(cc, fill=TRUE);rm(cc)
Measures<-cd[[2]]
Dimensions<-cd[[1]];rm(cd)
Parameters<-ca;rm(ca)

obje<-list(Measures=Measures, Dimensions=Dimensions, 
           Parameters=Parameters, Calculated.fields=Calculated.fields)

write_to_excel()



