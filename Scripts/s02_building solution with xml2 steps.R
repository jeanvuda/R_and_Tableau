x <- read_xml('
<root>
              <doc1 xmlns = "http://foo.com"><baz /></doc1>
              <doc2 xmlns = "http://bar.com"><baz /></doc2>
              </root>
              ')
xml_ns(x)
# When there are default namespaces, it's a good idea to rename
# them to give informative names:
ns <- xml_ns_rename(xml_ns(x), d1 = "foo", d2 = "bar")
ns
# Now we can pass ns to other xml function to use fully qualified names
baz <- xml_children(xml_children(x))
xml_name(baz)
xml_name(baz, ns)
xml_find_all(x, "//baz")
xml_find_all(x, "//foo:baz", ns)
str(as_list(x))
str(as_list(x, ns))

xml_children(xml_children(xml_chibvvldren(xml_children(xml_children(xml_children(xml_children(xml_children(xml_children(xml_children(x)
                                                                                                                     )))))))))
baz <- xml_find_all(x, ".//datasource")
xml_children(baz)[45]
xml_attr(baz, "formula")
table(duplicated(xml_attr(baz, "formula")))
xml_attr(baz, "formula")[!duplicated(xml_attr(baz, "formula"))]

xml_attr(xml_children(baz)[45], "caption")
xml_children(xml_children(baz)[45])
xml_attr(xml_children(xml_children(baz)[45]), "formula")
