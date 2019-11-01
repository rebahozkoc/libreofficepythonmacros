import helper
import time


def test_start_libreoffice():
    try:
        desktop = helper.start_libreoffice()
        print("LibreOffice is started.")
        assert desktop is not None
        return desktop
    except:
        print("ERROR: LibreOffice could not be started.")
        return False


def test_open_new_files(desktop):
    try:
        calc = helper.new_calc_doc(desktop)
        print("calc is started.")
        assert calc is not None
        time.sleep(0.5)
        calc.close(False)
        print("calc is closed.")
        time.sleep(1)
    except:
        print("ERROR: Calc could not be started.")
        return False
    try:
        impress = helper.new_impress_doc(desktop)
        print("impress is started.")
        assert impress is not None
        time.sleep(0.5)
        impress.close(False)
        print("impress is closed.")
        time.sleep(1)
    except:
        print("ERROR: Impress could not be started.")
        return False
    try:
        writer = helper.new_writer_doc(desktop)
        print("writer is started.")
        assert writer is not None
        time.sleep(0.5)
        writer.close(False)
        print("writer is closed.")
        time.sleep(1)
    except:
        print("ERROR: writer could not be started.")
        return False


def test_open_existing_file(desktop):
    try:
        doc = helper.open_doc(desktop, "/home/rebahlinux/libreofficepythonmacros/text_example.odt")
        assert doc is not None
        print("Existing writer document opened.\n\n\n")
        time.sleep(1)
        doc.close(False)
        print("Existing file is closed.")
        return True
    except:
        print("ERROR: Existing writer document could not be opened.")
        return False


def test_writer_functions(desktop):
    doc = helper.open_doc(desktop, "/home/rebahlinux/libreofficepythonmacros/text_example.odt")
    error_counter = 0
    try:
        helper.writer_text_replace_all(doc, "office", "ofis", case_sensitive=False, whole_words=False)
        print("All 'office' words are changed with 'ofis")
        assert helper.writer_text_find(doc,"ofis")
    except:
        print("ERROR: Function: write_text_replace_all or writer_text_find SECTION: 1 ")
        error_counter += 1
    try:
        helper.writer_text_replace_all(doc, "office", "ofis", case_sensitive=False, whole_words=False)
        print("All 'office' words are changed with 'ofis', indeed.")
        assert helper.writer_text_find(doc,"ofis")
    except:
        print("ERROR: Function: write_text_replace_all or writer_text_find SECTION: 2 ")
        error_counter += 1
    try:
        helper.writer_text_replace_all(doc, "fork", "catal", case_sensitive=False, whole_words=True)
        print("All 'fork' whole words are changed with 'catal'.")
        assert not helper.writer_text_find(doc,"catal")
    except:
        print("ERROR: Function: write_text_replace_all or writer_text_find SECTION: 3 ")
        error_counter += 1
    try:
        helper.writer_text_replace_all(doc, "starofis", "yıldızofis", case_sensitive=True, whole_words=False)
        print("All 'starofis' words with lower case letter are changed with 'yıldızofis'.")
        assert not helper.writer_text_find(doc, "yıldızofis")
        assert helper.writer_text_find(doc, "Starofis")
    except:
        print("ERROR: Function: write_text_replace_all or writer_text_find SECTION: 4 ")
        error_counter += 1
    try:
        print("Test of writer_text_insert(doc, content) is starting")
        time.sleep(1)
        helper.writer_text_insert(doc, "TEST")
        print("Selected word changed with TEST")
        assert helper.writer_text_find(doc,"TEST")
    except:
        print("ERROR: Function: writer_text_insert SECTION: 1")
        error_counter += 1
    try:
        time.sleep(1)
        #text = doc.Text
        viewCursor = doc.getCurrentController()
        oVC = viewCursor.getViewCursor()
        oVC.goRight(1, False)
        helper.writer_text_insert(doc, "TEST")
        print("TEST is inserted.")
        assert helper.writer_text_find(doc,"TEST.TEST")
    except:
        print("ERROR: Function: writer_text_insert SECTION: 2")
        error_counter += 1
    try:
        time.sleep(1)
        helper.writer_image_insert_from_file(doc, "/home/rebahlinux/libreofficepythonmacros/image.jpg")
        print("An image imported.")
    except:
        print("ERROR: Function: writer_image_insert_from_file SECTION: 1")
        error_counter += 1
    try:
        time.sleep(1)
        helper.writer_image_insert_from_file(doc, "/home/rebahlinux/libreofficepythonmacros/image.jpg")
        print("Another image imported.")
    except:
        print("ERROR: Function: writer_image_insert_from_file SECTION: 2")
        error_counter += 1
    try:
        time.sleep(1)
        doc.close(False)
        print("End of the Writer testing.")
    except:
        print("ERROR: Writer could not be closed.")
        error_counter += 1
    else:
        print("test_writer_functions is completed with", error_counter,"error(s).\n\n\n")
        return True


def test_calc_functions(desktop):
    doc = helper.open_doc(desktop, "/home/rebahlinux/libreofficepythonmacros/calc_example.ods")
    print("Existing calc document opened.")
    error_counter = 0
    try:
        time.sleep(1)
        active_sheet = helper.calc_get_active_sheet(doc)
    except:
        print("ERROR: Could not get the active sheet.")
        error_counter += 1
    try:
        time.sleep(1)
        cell = helper.calc_get_cell_from_sheet("A3",active_sheet)
        assert cell is not None
    except:
        print("ERROR: Could not get cell.")
        error_counter += 1
    try:
        time.sleep(1)
        helper.calc_set_cell_text(cell, "TESTING SET CELL")
        print("Added a string to specified cell.")
        assert helper.calc_get_cell_text(cell) == "TESTING SET CELL"
    except:
        print("ERROR: calc_set_cell_text or calc_get_cell_text failed SECTION: 1")
        error_counter += 1
    try:
        time.sleep(1)
        helper.calc_set_cell_text(cell, "TESTING SET CELL AGAIN")
        print("Added a string to the same cell.")
        assert helper.calc_get_cell_text(cell) == "TESTING SET CELL AGAIN"
    except:
        print("ERROR: calc_set_cell_text or calc_get_cell_text failed SECTION: 2")
        error_counter += 1
    try:
        cell = helper.calc_get_cell_from_sheet("A4",active_sheet)
        assert helper.calc_get_cell_text(cell) == "I am at A4"
        print("An existing string checked. (Success)")
    except:
        print("ERROR: calc_get_cell_text failed SECTION: 1")
        error_counter += 1
    try:
        time.sleep(1)
        assert helper.calc_str_exists_in_cells("A1","AA21", "E4 cell",active_sheet)
        print("calc_str_exists_in_cells is called (Success)")
    except:
        print("ERROR: calc_str_exists_in_cells failed SECTION: 1")
        error_counter += 1
    try:
        time.sleep(1)
        assert not helper.calc_str_exists_in_cells("A1","D5", "E4 cell",active_sheet)
        print("calc_str_exists_in_cells is called again (Success)")
    except:
        print("ERROR: calc_str_exists_in_cells failed SECTION: 2")
        error_counter += 1
    try:
        time.sleep(1)
        assert helper.calc_value_exists_in_cells("A1","AA21", 42, active_sheet)
        print("calc_value_exists_in_cells is called (Success)")
    except:
        print("ERROR: calc_value_exists_in_cells failed SECTION: 1")
        error_counter += 1
    helper.calc_value_exists_in_cells("AA1","BD5", 42, active_sheet)
    try:
        time.sleep(1)
        assert not helper.calc_value_exists_in_cells("AA1","BD5", 42, active_sheet)
        print("calc_value_exists_in_cells is called again (Success)")
    except:
        print("ERROR: calc_value_exists_in_cells failed SECTION: 2")
        error_counter += 1







desktop = test_start_libreoffice()
#test_open_new_files(desktop)
#test_open_existing_file(desktop)
#test_writer_functions(desktop)
test_calc_functions(desktop)





#test_function()


