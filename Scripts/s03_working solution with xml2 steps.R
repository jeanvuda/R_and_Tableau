
# Access Tableau Workbook
wb_doc<-'C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Test\\wb1\\wb20160623.twb'

tempd<-tempdir()
tmpf<-tempfile(tmpdir=tempd, fileext=".twb")

file.copy(wb_doc, tmpf)
filed<-sub(pattern = ".twb", ".xml", tmpf)
file.rename(tmpf, filed)

# Read file
library(xml2)
library("XML")

x<-read_xml(filed)


# tfile<-"C:/Users/User_2/Documents/01_Okaki/02_CPSA/Tableau Tabular/Main Tableau 20160525.xml"
# x<-read_xml(tfile)


baz <- xml_find_all(x, ".//datasources//datasource")
length(baz)
xml_children(baz[1])

getNodeSet(xmlParse(as.character(baz)))



# Work on Parameters
fp<-xml_children(baz[1])
idx<-which(grepl('<aliases enabled="yes"/>', fp))+1
fp<-fp[idx:length(fp)]
namess<-c("caption", "datatype", "name", "role")
namess
fp<-plyr::llply(namess, 
       function(x) xml_attr( fp, x))
names(fp)
fp<-do.call(cbind, fp)
colnames(fp)<-namess

#
fp<-xml_children(baz[1])
idx<-which(grepl('<aliases enabled="yes"/>', fp))+1
fp<-fp[idx:length(fp)]

namess<-c("caption", "datatype", "name", "role")
namess
fp<-plyr::llply(namess, 
                function(x) xml_attr( fp, x))
names(fp)
fp<-do.call(cbind, fp)
colnames(fp)<-namess

conv.tabl<-function(x=filed, lisnum){
  library(xml2)
  x<-read_xml(x)
  baz <- xml_find_all(x, ".//datasources//datasource")
  fp<-xml_children(baz[lisnum])
  idx<-which(grepl('<aliases enabled="yes"/>', fp))+1
  fp<-fp[idx:length(fp)]
  
  namess<-c("caption", "datatype", "name", "role")
  namess
  fp<-plyr::llply(namess, 
                  function(x) xml_attr( fp, x))
  names(fp)
  fp<-do.call(cbind, fp)
  colnames(fp)<-namess
  fp<-fp[complete.cases(fp),]
  return(fp)
}

fp<-conv.tabl(lisnum=1)

