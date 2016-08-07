
# call the workbook
wb='C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Test\\wb1\\wb20160803.twb'
library("XML")
doc <- xmlParse(file = wb, useInternalNodes = T)

## Select the nodes we want to update
nodes <- getNodeSet(doc, "//workbook//datasource//column//calculation")
## For each node, apply gsub on the content of the node

form1<-t(xpathSApply(n, '.', xmlAttrs))[2]

lapply(nodes, function(n) {
  n<-nodes[[36]]
  #xpathSApply(n, '.', xmlAttrs)
  #xmlAttrs(n) <- gsub("# High Quantity Known Prescribers","# High Quantity Known Prescribers_JohnNiruba",xmlAttrs(n))
  #t(xpathSApply(n, '.', xmlAttrs))
  #xmlAttrs(n) <- gsub(as.character(form1), as.character(form2), xmlAttrs(n), fixed = TRUE)
  
  xpathSApply(n, ".", xmlGetAttr, "formula")
  xpathSApply(n, ".//formula", xmlV)
})

saveXML(doc, file='C:\\Users\\User_2\\Documents\\01_Okaki\\10_Test\\xml testing\\sample_w.twb')

which(cb[,formula]==form1)
stringi::stri_trim(cb[which(cb[,formula]==form1),nformula][[1]])
form2<-"IF {FIXED [Pharmacy Licence Number], [Quarter + Year]:\r\nMAX(IF [Analytic Class]='Opioid' or [Analytic Class]='Benzodiazepine' THEN\r\n(IF {FIXED [Patient Alias Group], [Quarter + Year]:\r\nCOUNTD(IF [Analytic Class]='Opioid' THEN [Main Ingredient] END)>=[Parameters].[Select # Opiod Ingredients] AND\r\nCOUNTD(IF [Analytic Class]='Benzodiazepine' THEN [Main Ingredient] END)>=[Parameters].[Select # Benzo Ingredients for Benzos PLUS Opioids]\r\n}=TRUE\r\nTHEN 1 ELSE 0 END)\r\nEND)}>=1 THEN [Pharmacy Licence Number] END"

cb[3,]
