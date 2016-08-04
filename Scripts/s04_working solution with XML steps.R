
get_vars<-function(proc=rootnode){
  # find out how many datasources
  nodenames<-sapply(1:length(xmlChildren(proc)), function(x) {#print(x)
    xmlName(proc[[x]])})
  xy<-which(nodenames=='datasources')
  nds<-length(xpathSApply(proc[[xy]], ".//datasource", xmlGetAttr, "caption"))
  
  ca<-lapply(1:nds, function(nn) { #print(x)
    cc<-getNodeSet(proc[[xy]][[nn]], ".//column")
    #x<-cc[[1]]
    ca<-lapply(cc, function(x) {
      DS<-xmlAttrs(x)['caption']
      ac<-t(xpathSApply(x, paste0(".//calculation"), xmlAttrs))
      ad<-t(xpathSApply(x, paste0(".//range"), xmlAttrs))
      ab<-t(xpathSApply(x, paste0("."), xmlAttrs))
      ae<-t(xpathSApply(x, paste0(".//members//member"), xmlAttrs))
      if(nn==1) data.frame(ab, ad, ae) else data.frame( DS, ab, ac)
    })
    library(data.table)
    ca[is.na(caption), caption:=gsub("\\[|\\]","",name)]
    ca[,caption:=paste0("[",caption,"]")])
  
  
  return(ca)})

# #for(yy in 2:length(xpathSApply(rootnode[[3]], ".//datasource", xmlGetAttr, "caption")){
# lapply(2:length(xpathSApply(proc[[xy]], ".//datasource", xmlGetAttr, "caption")), function(yy) {  
#   yy<-ifelse(parameter, 1,yy)
#   nn<-ifelse(parameter, 1,yy)
#   cc<-getNodeSet(proc[[xy]][[5]], ".//column")
#   ca<-data.table::rbindlist(lapply(cc, function(x) {
#     
#     DS<-xmlAttrs(proc[[xy]][[yy]])['caption']
#     ac<-t(xpathSApply(x, paste0(".//calculation"), xmlAttrs))
#     ad<-t(xpathSApply(x, paste0(".//range"), xmlAttrs))
#     ab<-t(xpathSApply(x, paste0("."), xmlAttrs))
#     ae<-t(xpathSApply(x, paste0(".//members//member"), xmlAttrs))
#     if(nn==1) data.frame(ab, ad, ae) else data.frame( DS, ab, ac)
#   }), fill=TRUE)
#   ca[is.na(caption), caption:=gsub("\\[|\\]","",name)]
#   ca[,caption:=paste0("[",caption,"]")]
#   return(ca)
# })
}

get_actions<-function(){
  cc<-getNodeSet(rootnode[[5]], ".//action")
  #x<-cc[1]
  z<-data.table::rbindlist(lapply(cc, function(x) {
    y<-t(xpathSApply(x, ".", xmlAttrs))
    r<-t(xpathSApply(x, ".//activation", xmlAttrs)) # works
    s<-t(xpathSApply(x, ".//source",xmlAttrs)) #works
    #t<-t(xpathSApply(x, ".//link",xmlAttrs))[c('caption')] #works
    data.frame(y,r,s)
  }),fill = TRUE)
  return(z)
}

ab<-get_actions()

cc<-getNodeSet(rootnode[[7]], ".//dashboard")
z<-data.table::rbindlist(lapply(cc, function(x) {
  y<-t(xpathSApply(x, ".", xmlAttrs))
  r<-t(xpathSApply(x, ".//activation", xmlAttrs)) # works
  s<-t(xpathSApply(x, ".//source",xmlAttrs))[c('dashboard')] #works
  data.frame(y,r,s)
}),fill = TRUE)



cc<-getNodeSet(rootnode[[6]], ".//worksheet")
t(xpathSApply(cc[[1]], ".", xmlAttrs))
t(xpathSApply(cc[[1]], ".//layout-options/title/formatted-text/run", xmlAttrs))
t(xpathSApply(cc[[1]], ".//layout-options//title//formatted-text//run", xmlGetAttr, '<value>'))

t(xpathSApply(cc[[5]], ".//table//panes", xmlValue))

xpathSApply(cc[[5]], ".//table//", xmlValue)

xmlAttrs(cc[[5]], ".//table//panes")

xmlName(xmlChildren(cc[[6]]))



cc<-getNodeSet(rootnode[[6]], ".//worksheet")
z<-data.table::rbindlist(lapply(cc, function(x) {
  y<-t(xpathSApply(x, ".", xmlAttrs))
  title<-paste(t(xpathSApply(x, ".//layout-options//title//formatted-text//run", xmlValue)), collapse='') # Need to be cleaned
  tooltip<-t(xpathSApply(x, ".//table//tooltip-style", xmlAttrs)) #works
  data.frame(y,title, tooltip)
}),fill = TRUE)


c(1:length(xmlChildren(rootnode)))

get_properties<-function(xmlfile=rootnode){
  nodenames<-sapply(1:length(xmlChildren(xmlfile)), function(x) {#print(x)
    xmlName(xmlfile[[x]])})
  xy<-which(nodenames=='datasources')
  dsources<-t(xpathSApply(xmlfile[[xy]], ".//datasource", xmlAttrs))
  caption<-cbind(xpathSApply(xmlfile[[xy]], ".//datasource", xmlGetAttr, "caption"))
  data.table::data.table(data.frame(caption,dsources))
}

dsource<-get_properties(xmlfile = rootnode)



get_vars<-function(proc=rootnode){
  # Get parameters from 
  get_parameter(proc=rootnode)
  
  nodenames<-sapply(1:length(xmlChildren(proc)), function(x) {#print(x)
    xmlName(proc[[x]])})
  xy<-which(nodenames=='datasources')
  nds<-length(xpathSApply(proc[[xy]], ".//datasource", xmlGetAttr, "caption"))
  ca<-lapply(2:nds, function(nn) { #print(nn)})
    cc<-getNodeSet(proc[[xy]][[nn]], ".//column")
    data.table::rbindlist(lapply(cc, function(xz) {
      ac<-t(xpathSApply(xz, paste0(".//calculation"), xmlAttrs))
      ab<-t(xpathSApply(xz, paste0("."), xmlAttrs))
      ae<-t(xpathSApply(xz, paste0(".//members//member"), xmlAttrs))
      data.frame( ac, ab, ae)
    }), fill=TRUE)})
  names(ca)<-unlist(xpathSApply(proc[[xy]], ".//datasource", xmlGetAttr, "caption")[2:nds])
  
  ca<-lapply(ca, function(xz) {
    library(data.table)
    xz[is.na(caption), caption:=gsub("\\[|\\]","",name)]
    xz[,caption:=paste0("[",caption,"]")]
  })
  
  return(ca)
}

ca<-get_vars()

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



ca<-get_vars()


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
      if(nn==1) {data.frame( ac, ab, ae)} else{data.frame(ab,ac)}
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


get_worksheets<-function(proc=rootnode){
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

wbs<-get_worksheets(rootnode)
