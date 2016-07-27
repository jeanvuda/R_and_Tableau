
# Load Libraries
#install.packages("XLConnect", dep=T)
library(XLConnect)

# Read Tableau workbook
conv.tabl<-function(x=filed, lisnum){
  library(xml2)
  x<-read_xml(x)
  baz <- xml_find_all(x, ".//datasources//datasource")
  fp<-xml_children(baz[lisnum])
  idx<-which(grepl('<aliases enabled="yes"/>', fp))+1
  fp<-fp[idx:length(fp)]
  
  namess<-c("caption", "datatype", "name", "role")

  fp<-plyr::llply(namess, 
                  function(x) xml_attr( fp, x))
  
  fp[[length(fp)+1]]<-xml_attr(baz[1], attr="name")

  fp<-do.call(cbind, fp)
  colnames(fp)<-c(namess, "name2")
  fp<-fp[complete.cases(fp),]
  return(fp)
}



make_rootnodes<-function(wb='C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Test\\wb1\\wb20160623.twb'){
  tempd<-tempdir()
  fext<-paste0(".", sub('.*\\.','',wb))
  tmpf<-tempfile(tmpdir=tempd, fileext=fext)
  file.copy(wb, tmpf)
  filed<-sub(pattern = fext, ".xml", tmpf)
  file.rename(tmpf, filed)
  library("XML")
  rootnode <- xmlRoot(xmlParse(file = filed))
  unlink("tempd", recursive=TRUE)
  return(rootnode)
}


get_vars<-function(proc=rootnode, parameter=TRUE){
  # find out how many datasources
  #for(yy in 2:length(xpathSApply(rootnode[[3]], ".//datasource", xmlGetAttr, "caption")){
  lapply(2:length(xpathSApply(rootnode[[3]], ".//datasource", xmlGetAttr, "caption")), function(yy) {  
    yy<-ifelse(parameter, 1,yy)
    nn<-ifelse(parameter, 1,yy)
    cc<-getNodeSet(proc[[3]][[nn]], ".//column")
    ca<-data.table::rbindlist(lapply(cc, function(x) {
      
      DS<-xmlAttrs(proc[[3]][[yy]])['caption']
      ac<-t(xpathSApply(x, paste0(".//calculation"), xmlAttrs))
      ad<-t(xpathSApply(x, paste0(".//range"), xmlAttrs))
      ab<-t(xpathSApply(x, paste0("."), xmlAttrs))
      ae<-t(xpathSApply(x, paste0(".//members//member"), xmlAttrs))
      if(nn==1) data.frame(ab, ad, ae) else data.frame( DS, ab, ac)
    }), fill=TRUE)
    ca[is.na(caption), caption:=gsub("\\[|\\]","",name)]
    ca[,caption:=paste0("[",caption,"]")]
    return(ca)
  })
}

get_folder<-function(){
  cc<-getNodeSet(rootnode[[3]], ".//folder")
  data.table::rbindlist(lapply(cc, function(x) {
    r<-t(xpathSApply(x, ".", xmlAttrs)) # works
    #s<-xpathSApply(x, ".", xmlGetAttr, "name") #not works
    var<-xpathSApply(x, ".//folder-item", xmlGetAttr, 'name') #works
    data.frame(r,var)
  }), fill=TRUE)
}


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

