
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

  fp<-plyr::llply(namess, function(x) xml_attr( fp, x))
  
  fp[[length(fp)+1]]<-xml_attr(baz[1], attr="name")

  fp<-do.call(cbind, fp)
  colnames(fp)<-c(namess, "name2")
  fp<-fp[complete.cases(fp),]
  return(fp)
}


# Get the rootnode of a tableau workbook's XML blob
# Some steps in this function are unnecessary
# You can parse tableau workbook (*.twb) directly by the following code
# wb='C:\\Users\\User_2\\Documents\\01_Okaki\\02_CPSA\\Tableau Tabular\\Test\\wb1\\wb20160623.twb'
# library("XML")
# rootnode <- xmlRoot(xmlParse(file = wb))
# xmlAttrs(rootnode)
# This function creates rootnode object in memory
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

# Get the list of parameters/variables/calculated dields from the XML blob
# Pay attention to number of datasources in the file. You can find out using: 
# xpathSApply(rootnode[[4]], ".//datasource", xmlGetAttr, "caption")
get_vars<-function(proc=rootnode){
  # Get the node names & find which one is datasource
  nodenames<-sapply(1:length(xmlChildren(proc)), function(x) {#print(x)
    xmlName(proc[[x]])})
  xy<-which(nodenames=='datasources')
  # Find out how many datasources are connected
  nds<-length(xpathSApply(proc[[xy]], ".//datasource", xmlGetAttr, "caption"))
  # Get parameters first 
  # get fomulas/measures/dimensions from all other datasources
  ca<-lapply(1:nds, function(nn) { #print(nn)})
    cc<-getNodeSet(proc[[xy]][[nn]], ".//column")
    data.table::rbindlist(lapply(cc, function(xz) {
      ab<-t(xpathSApply(xz, paste0("."), xmlAttrs))
      ac<-t(xpathSApply(xz, paste0(".//calculation"), xmlAttrs))
      ad<-t(xpathSApply(xz, paste0(".//range"), xmlAttrs))
      ae<-t(xpathSApply(xz, paste0(".//members//member"), xmlAttrs))
      if(nn==1) {data.frame( ac, ab, ae)} else{data.frame(ab,ad,ae)}
    }), fill=TRUE)})
  names(ca)<-c('Parameters', unlist(xpathSApply(proc[[xy]], ".//datasource", xmlGetAttr, "caption")[2:nds]))
  # Clean the caption
  ca<-lapply(ca, function(xz) {
    library(data.table)
    xz[is.na(caption), caption:=gsub("\\[|\\]","",name)]
    xz[,caption:=paste0("[",caption,"]")]
  })
  return(ca)
}

# Get list of parameters from the Tableau workbook
# Not separately needed
get_parameter<-function(proc=rootnode){
  nodenames<-sapply(1:length(xmlChildren(proc)), function(x) {#print(x)
    xmlName(proc[[x]])})
  xy<-which(nodenames=='datasources')
  cc<-getNodeSet(proc[[xy]][[1]], ".//column")
  data.table::rbindlist(lapply(cc, function(xz) {
    ab<-t(xpathSApply(xz, paste0("."), xmlAttrs))
    ad<-t(xpathSApply(xz, paste0(".//range"), xmlAttrs))
    ae<-t(xpathSApply(xz, paste0(".//members//member"), xmlAttrs))
    data.frame(ab,ad,ae)}),fill=TRUE)
}


# This function creates the folder structure used in Tableau
get_folder<-function(proc=rootnode){
  # Get the node names & find which one is datasource
  nodenames<-sapply(1:length(xmlChildren(proc)), function(x) {#print(x)
    xmlName(proc[[x]])})
  xy<-which(nodenames=='datasources')
  cc<-getNodeSet(proc[[xy]], ".//folder")
  data.table::rbindlist(lapply(cc, function(x) {
    r<-t(xpathSApply(x, ".", xmlAttrs)) # works
    #s<-xpathSApply(x, ".", xmlGetAttr, "name") #not works
    var<-xpathSApply(x, ".//folder-item", xmlGetAttr, 'name') #works
    data.frame(r,var)
  }), fill=TRUE)
}

# Write the data to excel workbook (template)
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


# This function gets the list of actions and associated attributes in the workbook
get_actions<-function(proc=rootnode){
  nodenames<-sapply(1:length(xmlChildren(proc)), function(x) {#print(x)
    xmlName(proc[[x]])})
  xy<-which(nodenames=='actions')
  cc<-getNodeSet(proc[[xy]], ".//action")
  #x<-cc[[1]]
  z<-data.table::rbindlist(lapply(cc, function(x) {
    y<-t(xpathSApply(x, ".", xmlAttrs))
    r<-t(xpathSApply(x, ".//activation", xmlAttrs)) # works
    s<-t(xpathSApply(x, ".//source",xmlAttrs)) #works
    t<-t(xpathSApply(x, ".//link",xmlAttrs)) #works
    data.frame(y,r,s,t)
  }),fill = TRUE)
  return(z)
}


# This function gets the list of worksheets and their X,Y axes
get_worksheets<-function(proc=rootnode){
  # get worksheet function
  get_worksheet<-function(proc){
    nodenames<-sapply(1:length(xmlChildren(proc)), function(x) {#print(x)
      xmlName(proc[[x]])})
    xy<-which(nodenames=='worksheets')
    cc<-getNodeSet(proc[[xy]], ".//worksheet")
    #x<-cc[[1]]
    z<-data.table::rbindlist(lapply(cc, function(x) {
      y<-t(xpathSApply(x, ".", xmlAttrs))
      title<-paste(t(xpathSApply(x, ".//layout-options//title//formatted-text//run", xmlValue)), collapse='') # Need to be cleaned
      tooltip<-t(xpathSApply(x, ".//table//tooltip-style", xmlAttrs)) #works
      cols<-xpathSApply(x, ".//table//cols", xmlValue) # Needs cleaning
      rows<-xpathSApply(x, ".//table//rows", xmlValue) # Needs cleaning
      data.frame(y,title, tooltip, cols, rows)
    }),fill = TRUE)}
  
  dsource<-get_properties(xmlfile = proc)
  cln<-dsource[,paste0("[",name,"].")]
  wbs<-get_worksheet(proc)
  wbs[]<-lapply(wbs, as.character)
  wbs[,rows:=lapply(wbs[,rows], function(x) {
    for(ii in cln){x<-gsub(ii,"",x, fixed = TRUE)}
    return(x)})]
  wbs[,cols:=lapply(wbs[,cols], function(x) {
    for(ii in cln){x<-gsub(ii,"",x, fixed = TRUE)}
    return(x)})]
  wbs[,title:=lapply(wbs[,title], function(x) {
    for(ii in cln){x<-gsub(ii,"",x, fixed = TRUE)}
    return(x)})]
  ab<-get_vars()
  cb<-ab[[5]]
  
  wbs[,rows:=lapply(wbs[,rows], function(x) {
    for(ii in 1:nrow(cb)){x<-gsub(gsub("\\[|\\]", "",cb[ii,name]),gsub("\\[|\\]", "",cb[ii,caption]),x, fixed = TRUE)}
    return(x)})]
  wbs[,cols:=lapply(wbs[,cols], function(x) {
    for(ii in 1:nrow(cb)){x<-gsub(gsub("\\[|\\]", "",cb[ii,name]),gsub("\\[|\\]", "",cb[ii,caption]),x, fixed = TRUE)}
    return(x)})]
  wbs[,title:=lapply(wbs[,title], function(x) {
    for(ii in 1:nrow(cb)){x<-gsub(gsub("\\[|\\]", "",cb[ii,name]),gsub("\\[|\\]", "",cb[ii,caption]),x, fixed = TRUE)}
    return(x)})]
  
  return(wbs)
}

# Get properties of datasources
get_properties<-function(xmlfile=rootnode){
  nodenames<-sapply(1:length(xmlChildren(xmlfile)), function(x) {#print(x)
    xmlName(xmlfile[[x]])})
  xy<-which(nodenames=='datasources')
  dsources<-t(xpathSApply(xmlfile[[xy]], ".//datasource", xmlAttrs))
  caption<-cbind(xpathSApply(xmlfile[[xy]], ".//datasource", xmlGetAttr, "caption"))
  data.table::data.table(data.frame(caption,dsources))
}
