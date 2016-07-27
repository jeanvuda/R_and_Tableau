xmlParent(rootnode[[3]][[1]][[2]])
xmlChildren(rootnode[[3]][[1]][[2]])
xmlName(rootnode[[3]][[1]][[2]])

idNodes<-getNodeSet(rootnode[[3]][[1]][[2]], "//column[@caption]")
c1<-lapply(idNodes, function(x) xmlValue(x))
ca<-lapply(idNodes, function(x) xmlAttrs(x)['caption'])
cb<-lapply(idNodes, function(x) xmlAttrs(x)['formula'])
cb<-lapply(idNodes, function(x) xmlAttrs(x)['name'])

head(c1)

idNodes[1]

idNodes<-getNodeSet(rootnode[[3]][[1]][[2]], "//column[@caption]")
ca<-data.table::rbindlist(lapply(idNodes, function(x) {
  caption<-xmlAttrs(x)['caption']
  name<-xmlAttrs(x)['name']
  cb<-t(xpathSApply(x, ".//calculation", xmlAttrs))
  cc<-t(xpathSApply(x, ".//range", xmlAttrs))
  data.frame(caption,name,cb,cc)
}), fill=TRUE)

cd<-t(xpathSApply(x, ".//members/member", xmlAttrs))

conv.tabl<-function(x=filed, lisnum){
  library(xml2)
  x<-read_xml(x)
  baz <- xml_find_all(x, ".//datasources/datasource[name()='Parameters']", xml_ns(x))
  fp<-xml_children(baz[lisnum])
  idx<-which(grepl('<aliases enabled="yes"/>', fp))+1
  fp<-fp[idx:length(fp)]
  
  namess<-c("caption", "datatype", "name", "role")
  
  fp<-plyr::llply(namess, 
                  function(x) xml_attr( fp, x))
  
  fp[[length(fp)+1]]<-xml_attr(baz[1], attr="name")
  
  fp<-do.call(cbind, fp)
  colnames(fp)<-c(namess, "name2")`
  fp<-fp[complete.cases(fp),]
  return(fp)
}


x<-read_xml(filed)
length(xml_attrs(x))
length(xml_children(x))

xml_name(
xml_children(x)[3])

xml_ns(xml_children(x)[3])
xml_path(xml_children(x)[3])

xml_attr

for(i in 1:length(xml_children(x))){
  print(xml_name(xml_children(x))[i])
}


lapply(1:length(xml_children(x)), function(i){
  print(xml_name(xml_children(x))[i])
})

xml_find_all(x, ".//datasources//datasource")
xml_find_all(x, "//workbook//datasources//datasource")
t1<-xml_find_all(x, ".//datasource")
xml_attrs(t1[125])

xml_attrs(t1)

xpathSApply(rootnode, "")

xml_attrs(x)
xml_attrs(xml_children(x)[3])


lapply(1:length(xml_children(x)[3]), function(i){
  print(xml_name(xml_children(x)[3])[i])
})

xmlToList(xmlParse(t1[1]))
