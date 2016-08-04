# Solution
library(data.table)

# Get functions
source('~/01_Okaki/02_CPSA/Tableau Tabular/R/Scripts/s01_Functions.R')


# call the workbook
wb='C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Test\\wb1\\wb20160803.twb'
library("XML")
rootnode <- xmlRoot(xmlParse(file = wb))

# Get database properties
dsource<-get_properties(xmlfile = rootnode)


# Generate list of parameters/formulas/calculated-fields
ab<-get_vars()

# Get parameters
# Generate list of parameters
ca<-ab[[1]]

# Generate variables from a datasource
names(ab)
cb<-ab[[5]]

# Cleaning formulas in cb
cb[!is.na(formula), nformula:=lapply(cb[!is.na(formula),formula], function(y) {
  for(i in c(1:nrow(cb))){y<-gsub(cb[i,name], cb[i,caption], y, fixed = TRUE)}
  return(y)})]


# Generating folder group
ce<-get_folder(rootnode)
names(ce)
data.table::setnames(ce, "name", "folder")


#Extract from excel sheet
library(XLConnect)
wb<-loadWorkbook("C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Tables_original_714.xlsx")
lst<-readWorksheet(wb, sheet = getSheets(wb))


# For reporting measures/dimensions
# By datasource / mesaure&dimension
# DataSource, role, Folder, Caption, datatype,  and formula if there is any

# get the list of Calculated fields
cc<-merge(ce, cb, by.x="var", by.y="name", all=TRUE)
cc[]<-lapply(cc, as.character)
cc<-merge(cc, lst[[4]], all.x=T, by.x=c("var"), by.y=c("Variable"))

cc<-lapply(c("dimensions", "measures"), function(x){
  cc[role.x==x & !is.na(class),.(role.y, folder, caption, datatype, var, nformula, Checked, Definition, Duplicate, Suggestion)]})

lapply(cc, head)


# get the list of data variables from the dbase 
cd<-merge(ce, cb, by.x="var", by.y="name", all=TRUE)
cd<-lapply(c("dimension", "measure"), function(x){
  cd[role.y==x & is.na(class),.(role.y, folder, caption, datatype, var)]})

lapply(cd, head)


# Get actions and worksheet info
ab<-get_actions(rootnode)
wbs<-get_worksheets(rootnode)


# For reporting parameters
# Role, Folder, Caption, datatype, etc
#ca<-get_vars(proc=rootnode, parameter=TRUE)[[1]]
ca<-get_parameter(rootnode)
ca<-merge(ce, ca, by.x="var", by.y="name", all.y=TRUE)
ca[,default.format:=NULL]
ca[,role.y:=NULL]


Calculated.fields<-data.table::rbindlist(cc, fill=TRUE);rm(cc)
Measures<-cd[[2]]
Dimensions<-cd[[1]];rm(cd)
Parameters<-ca;rm(ca)
Actions<-ab;rm(ab)
WorkBooks<-wbs;rm(wbs)

obje<-list(Data=dsource, Measures=Measures, Dimensions=Dimensions, 
           Parameters=Parameters, Calculated.fields=Calculated.fields,
           Actions=Actions, WorkBooks=WorkBooks)

write_to_excel()

