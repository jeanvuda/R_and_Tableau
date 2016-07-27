#install.packages("XML", dep=T)
library(xml2)
tfile<-"C:/Users/User_2/Documents/01_Okaki/02_CPSA/Tableau Tabular/Main Tableau 20160525.xml"
x<-read_xml(tfile)

x<-read_xml(filed)

# Find all baz nodes anywhere in the document
baz <- xml_find_all(x, ".//value")

baz
xml_path(baz)

xml_contents(baz)

xml_children(xml_children(baz)[1])


baz <- xml_find_all(x, ".//calculation")
xml_attr(baz, "formula")
table(duplicated(xml_attr(baz, "formula")))
xml_attr(baz, "formula")[!duplicated(xml_attr(baz, "formula"))]

cat(xml_attr(baz, "formula")[!duplicated(xml_attr(baz, "formula"))][20])

baz <- xml_find_all(x, ".//datasources//datasource")
baz
xml_attr(baz, "caption")
xml_attr(baz, "formula")

lapply(baz, xml_attr, "caption")
occupation <- lapply(baz, xml_find_all, "//formula")

xml_find_all(x, "formula")


dff<-xmlToList(xmlParse(tfile))
xml_children(x)
names(dff[[1]])
