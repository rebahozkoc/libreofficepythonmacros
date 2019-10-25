import helper

helper.start_libreoffice()

print("here")

ctx = helper.connect_to_libreoffice()

print(type(ctx))

model = helper.get_current_doc(ctx)

active_sheet = helper.get_active_sheet(model)

cell = helper.get_cell_from_sheet("B2", active_sheet)

helper.set_cell_text(cell, "deneme")


# print(helper.str_exists_in_cells("A1","AB3","Testing"))
#
# helper.set_cell_text_with_addr("B4","I am at B4")
