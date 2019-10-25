import helper


#doc = helper.new_impress_doc()

#doc = helper.new_writer_doc()
#text = doc.getText()
#text.setString("Hello world")


doc_calc = helper.new_calc_doc()
active_sheet = helper.get_active_sheet(doc_calc)
cell = helper.get_cell_from_sheet("B3", active_sheet)
helper.set_cell_text(cell, 'deneme2')


helper.set_cell_text_with_addr("AA2","Testing",active_sheet)
print(helper.str_exists_in_cells("A1","AB3","Testing", active_sheet))
helper.set_cell_text_with_addr("B4","I am at B4", active_sheet)
