import helper

#(oCursor.DBG_methods)

start_libreoffice_result = helper.start_libreoffice()
desktop = start_libreoffice_result[0]
doc = helper.open_doc(desktop, "/home/rebahlinux/libreofficepythonmacros/impress_example.odp")
dispatcher = helper.dispatcher(start_libreoffice_result[1])
search = helper.impress_text_search_dispatcher(doc, dispatcher, "Introducing")
