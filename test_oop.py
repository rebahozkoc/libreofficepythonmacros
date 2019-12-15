import helper_oop
import time


def standard_testing(function_name, expected_result, *args):
    """Assumes function_name is a function
    expected_result ==  3 ==> next argument becomes standard_testing's argument and given function's return is compared with this argument.
    expected_result ==  2 ==> expected result is True,
    expected_result ==  1 ==> expected result is value,
    expected_result ==  0 ==> expected result is False,
    expected_result == -1 ==> expected result is nothing
    expected_result == -2 ==> self
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
            print(result)
        elif expected_result == -2:
            print(function_name, "is executed successfully.\n")
            return function_name(*args)
        else:
            raise Exception("ERROR: Arguments of the standard_testing function are not provided properly.\n")
    except Exception as e:
        print("ERROR: SOMETHING WENT WRONG WITH:",function_name, e, "\n")
        return False
    else:
        print(function_name, "is executed successfully.\n")
        time.sleep(0.5)
        if expected_result != -2:
            return result


def test_open_new_files():
    writer = standard_testing(helper_oop.WordProcessor, -2)
    writer.closeFile()
    impress = standard_testing(helper_oop.Presentation, -2)
    impress.closeFile()
    calc = standard_testing(helper_oop.SpreadSheet, -2)
    calc.closeFile()


def test_open_existing_file():
    doc = standard_testing(helper_oop.WordProcessor, -2, "/home/rebahlinux/libreofficepythonmacros/text_example.odt")
    doc.closeFile()
    calc = standard_testing(helper_oop.SpreadSheet, -2, "/home/rebahlinux/libreofficepythonmacros/calc_example.ods")
    calc.closeFile()
    impress = standard_testing(helper_oop.Presentation, -2, "/home/rebahlinux/libreofficepythonmacros/impress_example.odp")
    impress.closeFile()


def test_calc_functions():
    calc = standard_testing(helper_oop.SpreadSheet, -2, "/home/rebahlinux/libreofficepythonmacros/calc_example.ods")
    active_sheet = standard_testing(calc.getActiveSheet, -2)
    cell = standard_testing(calc.getCellByName, -2, "A3")
    standard_testing(calc.setCellText, -2, "A3", "TESTING SET CELL")
    print(calc.getCellText("A3") == "TESTING SET CELL")
    standard_testing(calc.setCellText, -2, "A3", "TESTING SET CELL AGAIN")
    print(calc.getCellText("A3") == "TESTING SET CELL AGAIN")
    print(calc.getCellText("A4") == "I am at A4")
    print("checkstr:", standard_testing(calc.checkTextExists, -2, "A1", "AA21", "E4 cell"))
    print("checkstr:", not standard_testing(calc.checkTextExists, -2, "A1", "D5", "E4 cell"))
    print("checkvalueexists:", standard_testing(calc.checkValueExists, -2, "A1","AA21", 42))
    print("checkvaluexeists:", not standard_testing(calc.checkValueExists, -2, "AA1","BD5", 42))
    print("getStraddress:", calc.getStrAdressInCells("A1", "CC21", "I am at A4") == "A4")
    print("getStraddress:", calc.getStrAdressInCells("A1", "CC21", "TESTING SET CELL AGAIN") == "A3")
    print("getStraddress:", calc.getStrAdressInCells("BB1", "CC21", "TESTING SET CELL AGAIN") is False)
    print("getValueaddress:", calc.getValueAdressInCells("BB1", "CC200", 3.14159) is False)
    print("getValueaddress:", calc.getValueAdressInCells("A4", "CC200", 3.14159) == "B5")
    print("getCellText:", calc.getCellText("E4") == "E4 cell")
    print("getCellText:", calc.getCellText("C20") == "")
    standard_testing(calc.setCellText, -2, "K9", "K9 is a smart dog.")
    print("getCellText:", calc.getCellText("K9") == "K9 is a smart dog.")
    calc.closeFile()


def test_writer_functions():
    doc = standard_testing(helper_oop.WordProcessor, -2, "/home/rebahlinux/libreofficepythonmacros/text_example.odt")
    doc.textReplaceAll("office", "placeholder")

def test_impress_functions():
    pass


#test_open_new_files()
#test_calc_functions()
test_writer_functions()
