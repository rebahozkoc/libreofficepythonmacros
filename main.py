import helper
desktop = helper.start_libreoffice()
doc = helper.open_doc(desktop, "/home/rebahlinux/libreofficepythonmacros/impress_example.odp")

#helper.impress_remove_page_by_page_number(doc, 2)

print(helper.impress_text_search(doc, "Introducing"))

#print(helper.impress_text_search(doc, "Mudur"))


#helper.impress_new_page(doc)
