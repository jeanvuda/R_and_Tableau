write_to_excel<-function(template='C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Template.xlsx',
                         ffinal="C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Tables_New_1.xlsx",
                         objects=obje){
  tmp<-loadWorkbook(template)
  for(objs in objects){
    sheet_name<-deparse(substitute(objs))
    cloneSheet(tmp, sheet="Set", name=sheet_name)
    writeWorksheet(tmp, objs,sheet_name, startCol = 2)
  }
  #deparse(substitute(Measures))
  #wb<-"C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Tables_New_1.xlsx"
  removeSheet(tmp, sheet="Set")
  saveWorkbook(tmp, wb)
}

sheet_name<-"Measures"
tmp<-loadWorkbook('C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Template.xlsx')
cloneSheet(tmp, sheet="Set", name=sheet_name)
#createName(tmp, name=sheet_name, formula=paste0(sheet_name,"!$B$1"))
#writeNamedRegion(tmp, Measures, name=sheet_name)
writeWorksheet(tmp, Measures,sheet_name, startCol = 2)
cloneSheet(tmp, sheet="Set", name="Dimensions")
writeWorksheet(tmp, Dimensions,'Dimensions', startCol = 2)
wb<-"C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Tables_New_1.xlsx"

removeSheet(tmp, sheet="Set")
saveWorkbook(tmp, wb)

write_to_excel()

write_to_excel<-function(template='C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Template.xlsx',
                         ffinal="C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Tables_New_1.xlsx",
                         objects=obje){
tmp<-loadWorkbook(template)
lapply(names(obje), function(xy) {
  sheet_name<-xy
  cloneSheet(tmp, sheet="Set", name=sheet_name)
  writeWorksheet(tmp, obje[[sheet_name]], sheet_name, startCol=2)})
removeSheet(tmp, sheet="Set")
saveWorkbook(tmp, ffinal)}

write_to_excel()

obje$Measures

lapply(names(obje), function(xy) {
  sheet_name<-xy
  print(sheet_name)})

names(deparse(substitute(obje[2])))

tmp<-loadWorkbook('C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Template.xlsx')
sheet_name<-"Measures"
cloneSheet(tmp, sheet="Set", name=sheet_name)
#createName(tmp, name=sheet_name, formula=paste0(sheet_name,"!$B$1"))
#writeNamedRegion(tmp, Measures, name=sheet_name)
writeWorksheet(tmp, Measures,sheet_name, startCol = 2)
wb<-"C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Tables_New_1.xlsx"
removeSheet(tmp, sheet="Set")
saveWorkbook(tmp, wb)

