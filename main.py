import helper

#(oCursor.DBG_methods)

start_libreoffice_result = helper.start_libreoffice()
desktop = start_libreoffice_result[0]



#Calc replace all
#doc = helper.open_doc(desktop, "/home/rebahlinux/libreofficepythonmacros/calc_example.ods")
#active_sheet = helper.calc_get_active_sheet(doc)
#print(helper.calc_text_replace_all(active_sheet, "E4 cell", "E5"))


#Impress replace all
#doc = helper.open_doc(desktop, "/home/rebahlinux/libreofficepythonmacros/impress_example.odp")
#print(helper.impress_text_replace_all(doc, "open", "acik"))

helper.macro_wiper("calc_embeded.ods", "myscript.py", "calc_set_cell_text_with_addr")
helper.delete_all_macros("calc_embeded.ods")

#Impress searhing and selecting
#doc = helper.open_doc(desktop, "/home/rebahlinux/libreofficepythonmacros/impress_example.odp")
#dispatcher = helper.dispatcher(start_libreoffice_result[1])
#search = helper.impress_text_search_dispatcher(doc, dispatcher, "fork")
