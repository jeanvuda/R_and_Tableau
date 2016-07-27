# Parse the XML file
result <- xmlParse(file = filed)

# Get the top nodes
rootnode <- xmlRoot(result)

# Checking attributes of topnodes
xmlName(rootnode)
xmlSize(rootnode)
xmlName(rootnode[[1]])
xmlChildren(rootnode[[3]])
xmlAttrs(rootnode)
xmlAttrs(rootnode[[1]])

sapply(getNodeSet(result, "//workbook"), function(el) xmlGetAttr(el, "version"))

data.frame(xmlAttrs(rootnode[[1]]))

els = getNodeSet(rootnode, "//datasources/datasource/connection")
els<-sapply(els, function(el) xmlGetAttr(el, "dbname"))
els

els = getNodeSet(rootnode, "//datasources/datasource/connection/relation")
els<-sapply(els, function(el) xmlGetAttr(el, "table"))
els

xmlSApply(rootnode[[3]], xmlName)
xmlSApply(rootnode[[3]], xmlSize)
xmlSApply(rootnode[[3]], xmlAttrs)
xmlSApply(rootnode[[3]], xmlAttrs)[1,2]

alert.df<-plyr::ldply(t(xmlSApply(rootnode[[3]][[1]], xmlAttrs)), data.frame)

calculation.df <- as.data.frame(t(xpathSApply(rootnode,".//datasources/datasource/repository-location",xmlAttrs)))
calculation.df <- as.data.frame(t(xpathSApply(rootnode,".//datasources/datasource/column",xmlAttrs)))
calculation.df <- as.data.frame(t(xpathSApply(rootnode,".//datasources/datasource/connection",xmlAttrs)))
calculation.df <- as.data.frame(t(xpathSApply(rootnode,".//datasources/datasource/connection/calculations/calculation",xmlAttrs)))


calculation.df <- as.data.frame(t(xpathSApply(rootnode,".//datasources/datasource/column/calculation",xmlAttrs)))

getNodeSet(rootnode, './/datasources/datasource/column/calculation')


plyr::ldply(xpathSApply(rootnode,".//datasources/datasource/column",xmlAttrs))


baz <- xml_find_all(x, ".//datasources/datasource/column")
fp<-xml_children(baz)
fp<-baz
idx<-which(grepl('<aliases enabled="yes"/>', fp))+1
fp<-fp[idx:length(fp)]

namess<-c("caption", "datatype", "name", "role", "formula")
namess
fp<-plyr::llply(namess, 
                function(x) xml_attr( fp, x))
names(fp)
fp<-do.call(cbind, fp)
colnames(fp)<-namess
fp<-fp[complete.cases(fp),]
return(fp)


ids<-getNodeSet(rootnode, ".//datasources/datasource/column")
capt<-lapply(ids, function(x) xmlAttrs(x)['caption'])
fff<-sapply(xmlChildren(rootnode[[3]][[1]]), xmlAttrs)
ffg<-sapply(xmlChildren(rootnode[[3]][[2]]), xmlAttrs)
ffh<-sapply(xmlChildren(rootnode[[3]][[3]]), xmlAttrs)
ffh<-sapply(xmlChildren(rootnode[[3]][[1]][[2]], xmlAttrs)
fff<- lapply(ids, xpathApply, path='./Calculation[@formula]', xmlAttrs)

xmlAttrs(rootnode[[3]][[1]][[2]])
xmlAttrs(rootnode[[3]][[1]][[1]])
xmlAttrs(rootnode[[3]][[1]][[2]][[1]])

test<-do.call(rbind.data.frame, mapply(cbind, capt, fff))
            