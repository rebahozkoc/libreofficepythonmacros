import helper
import time
desktop = helper.start_libreoffice()
doc = helper.open_doc(desktop, "/home/rebahlinux/libreProject/text_example.odt")

#text = doc.Text
#helper.text_replace_all(doc, "free software", "open source", case_sensitive=False, whole_words=True)
print("bulundu mu?:", helper.text_find(doc,"free software"))

# model = helper.new_calc_doc()
# active_sheet = helper.get_active_sheet(model)
# cell = helper.get_cell_from_sheet("B3", active_sheet)
# helper.set_cell_text(cell, 'deneme2')
#
#
# helper.set_cell_text_with_addr("AA2","Testing",active_sheet)
# print(helper.str_exists_in_cells("A1","AB3","Testing", active_sheet))
# helper.set_cell_text_with_addr("B4","I am at B4", active_sheet)

#desktop = helper.start_libreoffice()
#helper.open_doc(desktop,"/home/rebahlinux/libreProject/deneme.ods")


# import os
# directory = os.path.dirname(model.URL)[7:]
# print(directory)
#(oCursor.DBG_methods)


#search = doc.createSearchDescriptor()
#search.SearchString = "free software"
# found = doc.findFirst(search)
# viewCursor = doc.getCurrentController()
# oVC = viewCursor.getViewCursor()
# ara_isim = oVC.getText()
# oCursor= ara_isim.createTextCursorByRange(oVC)
# oVC.goRight(4,False)
#helper.writer_text_insert(doc)


#helper.writer_image_instert_from_file(doc)
