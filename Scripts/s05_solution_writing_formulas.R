
# Call the new formulas
library(XLConnect)
wb1<-loadWorkbook("C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Documentation\\Tables_New_with NewFormulas.xlsx")
lst<-readWorksheet(wb1, sheet = 5)
lst<-lst[!is.na(lst$nformula2),c(1:3,5,7)]

# call the workbook
wb1='C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Test\\wb1\\wb20160803.twb'
library("XML")
doc <- xmlParse(file = wb1, useInternalNodes = T)

## Select the nodes we want to update
nodes <- getNodeSet(doc, "//workbook//datasource//column")


# Replace formulas
for(i in c(1:nrow(lst))){ #print(lst$var[i])
  lapply(nodes, function(n){
    if(xpathSApply(n, ".", xmlGetAttr, "name")==lst$var[i]){
      form1<-xpathSApply(n, ".//calculation", xmlGetAttr, "formula")
      #xmlAttrs(n) <- gsub(as.character(form1), lst$nformula2[i], xmlAttrs(n), fixed = TRUE)
      sapply(getNodeSet(n, ".//calculation"), function(x) { xmlAttrs(x) <-
        gsub(as.character(form1), lst$nformula2[i], xmlAttrs(x), fixed = TRUE)})
    } }) }

# Write the file back  
saveXML(doc, file='C:\\Users\\User_2\\Documents\\01_Okaki\\10_Test\\xml testing\\sample_w.twb')

