import helper
import time


def standard_testing(function_name, expected_result, *args):
    """Assumes function_name is a function
    expected_result ==  3 ==> next argument becomes standard_testing's argument and given function's return is compared with this argument.
    expected_result ==  2 ==> expected result is True,
    expected_result ==  1 ==> expected result is value,
    expected_result ==  0 ==> expected result is False,
    expected_result == -1 ==> expected result is nothing
    rest is the given function's arguments.
    :returns True if function works properly, false otherwise."""
    try:
        if expected_result == 3:
            check_value = args[0]
            function_arguments = args[1:]
            assert check_value == function_name(*function_arguments)
            result = True
        elif expected_result == 2 :
            result = function_name(*args)
            assert result == True
        elif expected_result == 1:
            result = function_name(*args)
            assert result is not None
        elif expected_result == 0:
            result = not function_name(*args)
            assert not function_name(*args)
        elif expected_result == -1:
            function_name(*args)
            result = True
        else:
            raise Exception("ERROR: Arguments of the standard_testing function are not provided properly.\n")
    except Exception as e:
        print("ERROR: SOMETHING WENT WRONG WITH:",function_name, e, "\n")
        return False
    else:
        print(function_name, "is executed successfully.\n")
        time.sleep(0.5)
        return result


def test_writer_functions(desktop):
    doc = standard_testing(helper.open_doc, 1, desktop, "/home/rebahlinux/libreofficepythonmacros/text_example.odt")
    standard_testing(helper.writer_text_replace_all, -1, doc, "office", "placeholder",False, False)
    standard_testing(helper.writer_text_find, 2, doc,"placeholder")
    standard_testing(helper.writer_text_replace_all, -1, doc, "office", "placeholder", False, False)
    standard_testing(helper.writer_text_find, 2, doc,"placeholder")
    standard_testing(helper.writer_text_replace_all,-1, doc, "fork", "spoon", False, True)
    standard_testing(helper.writer_text_find, 0, doc, "spoon")
    standard_testing(helper.writer_text_replace_all, -1, doc, "starplaceholder", "planetplaceholder", True, False)
    standard_testing(helper.writer_text_find, 0, doc, "planetplaceholder")
    standard_testing(helper.writer_text_find, 2, doc, "Starplaceholder")
    standard_testing(helper.writer_text_insert, -1, doc, "TEST")
    standard_testing(helper.writer_text_find, 2, doc,"TEST")
    viewCursor = doc.getCurrentController()
    oVC = viewCursor.getViewCursor()
    oVC.goRight(1, False)
    standard_testing(helper.writer_text_insert, -1, doc, "TEST")
    standard_testing(helper.writer_text_find, 2, doc,"TEST.TEST")
    standard_testing(helper.writer_image_insert_from_file, -1, doc, "/home/rebahlinux/libreofficepythonmacros/image.jpg")
    standard_testing(helper.writer_image_insert_from_file, -1, doc, "/home/rebahlinux/libreofficepythonmacros/image.jpg")
    doc.close(False)


def test_open_new_files(desktop):
    writer = standard_testing(helper.new_writer_doc, 1, desktop)
    writer.close(False)
    impress = standard_testing(helper.new_impress_doc, 1, desktop)
    impress.close(False)
    calc = standard_testing(helper.new_calc_doc, 1, desktop)
    calc.close(False)


def test_open_existing_file(desktop):
    doc = standard_testing(helper.open_doc, 1, desktop, "/home/rebahlinux/libreofficepythonmacros/text_example.odt")
    doc.close(False)
    calc = standard_testing(helper.open_doc, 1, desktop, "/home/rebahlinux/libreofficepythonmacros/calc_example.ods")
    calc.close(False)


def test_calc_functions(desktop):
    doc = helper.open_doc(desktop, "/home/rebahlinux/libreofficepythonmacros/calc_example.ods")
    active_sheet = standard_testing(helper.calc_get_active_sheet, 1, doc)
    cell = standard_testing(helper.calc_get_cell_from_sheet, 1, "A3", active_sheet)
    standard_testing(helper.calc_set_cell_text, -1, cell, "TESTING SET CELL")
    standard_testing(helper.calc_get_cell_text, 3, "TESTING SET CELL", cell)
    standard_testing(helper.calc_set_cell_text, -1, cell, "TESTING SET CELL AGAIN")
    standard_testing(helper.calc_get_cell_text, 3, "TESTING SET CELL AGAIN", cell)
    cell = standard_testing(helper.calc_get_cell_from_sheet, 1, "A4", active_sheet)
    standard_testing(helper.calc_get_cell_text, 3, "I am at A4", cell )
    standard_testing(helper.calc_str_exists_in_cells, 2, "A1", "AA21", "E4 cell", active_sheet)
    standard_testing(helper.calc_str_exists_in_cells, 0, "A1","D5", "E4 cell", active_sheet)
    standard_testing(helper.calc_value_exists_in_cells, 2, "A1","AA21", 42, active_sheet) # I am here
    standard_testing(helper.calc_value_exists_in_cells, 0, "AA1","BD5", 42, active_sheet)
    standard_testing(helper.calc_search_str_in_cells, 3, "A4","A1","CC21","I am at A4",active_sheet)
    standard_testing(helper.calc_search_str_in_cells, 3, "A3", "A1", "CC21", "TESTING SET CELL AGAIN", active_sheet)
    standard_testing(helper.calc_search_str_in_cells, 0, "BB1","CC21","TESTING SET CELL AGAIN",active_sheet)
    standard_testing(helper.calc_search_value_in_cells, 0, "BB1","CC200", 3.14159, active_sheet)
    standard_testing(helper.calc_search_value_in_cells, 3, "B5", "A4", "CC21", 3.14159, active_sheet)
    #standard_testing(helper.calc_search_value_in_cells, 0, "BB1","CC2000", 3.14159, active_sheet)
    standard_testing(helper.calc_get_cell_text_with_addr, 3, "E4 cell", "E4", active_sheet)
    standard_testing(helper.calc_get_cell_text_with_addr, 3, "", "C20", active_sheet)
    standard_testing(helper.calc_set_cell_text_with_addr, -1, "K9", "K9 is a smart dog.", active_sheet)
    standard_testing(helper.calc_get_cell_text_with_addr, 3, "K9 is a smart dog.", "K9", active_sheet)
    doc.close(False)


desktop = standard_testing(helper.start_libreoffice, 1)
test_open_new_files(desktop)
test_open_existing_file(desktop)
test_calc_functions(desktop)
test_writer_functions(desktop)

print("out of program.")
