import helper
import helper_oop
import time

#These two line is common. Uncomment a code block to try an example.
#libreoffice = helper_oop.WordProcessor()
#desktop = libreoffice.getDesktop()

#print(type(libreoffice))
#a = input("sfsz")
#Calc replace all
#helper.new_calc_doc(desktop)

#active_sheet = helper.calc_get_active_sheet(model)
#print(helper.calc_text_replace_all(active_sheet, "E4 cell", "E5"))


#Impress replace all
# These lines doesn't working due a bug from LibreOffice API. The bug is reported.
#doc = helper.open_doc(desktop, "/home/rebahlinux/libreofficepythonmacros/impress_example.odp")
#print(helper.impress_text_replace_all(doc, "open", "acik"))

#Delete embeded macros
#helper.macro_wiper("calc_embeded.ods", "myscript.py", "calc_set_cell_text_with_addr")
#helper.delete_all_macros("calc_embeded.ods")


#Impress searhing and selecting
#doc = helper.open_doc(desktop, "/home/rebahlinux/libreofficepythonmacros/impress_example.odp")
#dispatcher = helper.dispatcher(start_libreoffice_result[1])
#search = helper.impress_text_search_dispatcher(doc, dispatcher, "fork")
imp = helper_oop.Presentation("/home/rebahlinux/libreofficepythonmacros/impress_example.odp")
imp.checkTextExists("fork")
#Calc open a new file.
#model = helper.new_calc_doc(desktop)
