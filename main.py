import helper
desktop = helper.start_libreoffice()
doc = helper.open_doc(desktop, "/home/rebahlinux/libreProject/text_example.odt")

#helper.writer_text_replace_all(doc, "office", "ofis", case_sensitive=False, whole_words=False)


# model = helper.new_calc_doc()
# active_sheet = helper.get_active_sheet(model)
# cell = helper.get_cell_from_sheet("B3", active_sheet)
# helper.set_cell_text(cell, 'deneme2')


#(oCursor.DBG_methods)


helper.writer_image_insert_from_file(doc,"/home/rebahlinux/libreProject/image.jpg")
import time
time.sleep(5)
doc.close(True)
