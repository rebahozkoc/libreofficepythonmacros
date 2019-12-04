import socket #Only for windows
import os
import shutil
import subprocess
import tempfile
import time
import uno
import uuid
from com.sun.star.connection import NoConnectException
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER

# This is another version of start_libreoffice()
# def start_libreoffice():
#    subprocess.Popen("soffice --calc --accept=\"socket,host=localhost,port=2002;urp;StarOffice.ServiceManager\"", shell=True)
#    time.sleep(1)


def start_libreoffice():
    # This depends on your LibreOffice installation location.
    sofficePath = '/opt/libreofficedev6.4/program/soffice'
    tempDir = tempfile.mkdtemp()

    # Restore cached profile if available
    userProfile = tempDir + '/profile'
    cacheDir = os.getenv('XDG_CACHE_DIR', os.environ['HOME'] + '/.cache') + '/lo_profile_cache'
    if os.path.isdir(cacheDir):
        shutil.copytree(cacheDir, userProfile)
        profileCached = True
    else:
        os.mkdir(userProfile)
        profileCached = False
    # Launch the LibreOffice server
    pipeName = uuid.uuid4().hex
    args = [
        sofficePath,
        '-env:UserInstallation=file://' + userProfile,
        '--pidfile=' + tempDir + '/soffice.pid',
        '--accept=pipe,name=' + pipeName + ';urp;',
        '--norestore',
        '--invisible']
    sofficeEnvironment = os.environ
    sofficeEnvironment['TMPDIR'] = tempDir
    subprocess.Popen(args, env=sofficeEnvironment, preexec_fn=os.setsid)
    # Open connection to server
    for i in range(100):
        try:
            localContext = uno.getComponentContext()
            resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",
                                                                             localContext)
            context = resolver.resolve("uno:pipe,name=%s;urp;StarOffice.ComponentContext" %
                                       pipeName)
            break
        except NoConnectException:
            time.sleep(0.1)
            if i == 99:
                raise
    # Cache profile if required
    if not profileCached:
        shutil.copytree(userProfile, cacheDir)
    return_list = [context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context), context]
    return return_list


def new_writer_doc(desktop):
    # Open a new writer file
    return desktop.loadComponentFromURL('private:factory/swriter', 'blank', 0, ())


def new_calc_doc(desktop):
    # Open a new spreadsheet file
    return desktop.loadComponentFromURL('private:factory/scalc', 'blank', 0, ())


def new_impress_doc(desktop):
    # Open a new Impress file
    return desktop.loadComponentFromURL('private:factory/simpress', 'blank', 0, ())


# path = uno.fileUrlToSystemPath(url) may help later

def open_doc(desktop, path1):
    fileURL = uno.systemPathToFileUrl(path1)
    return desktop.loadComponentFromURL(fileURL, "_blank", 0, ())


# Old version of connection function
# def connect_to_libreoffice():
#     localContext = uno.getComponentContext()
#     resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
#     return resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
# def get_current_doc(ctx1):
#     smgr = ctx1.ServiceManager
#     desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx1)
#     return desktop.getCurrentComponent()


def calc_get_active_sheet(model):
    return model.CurrentController.ActiveSheet


def calc_get_cell_from_sheet(cell_name, sheet):
    return sheet.getCellRangeByName(cell_name)


def calc_set_cell_text(cell, text):
    cell.String = text


def calc_get_cell_text(cell):
    return cell.String


# A helper function for Calc searching.
def inside_address_spliter(string):
    """Assumes string consist with letters and numbers, respectively.
     :returns a tuple whose first element is letters second is numbers"""
    digits = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
    number = ""
    for i in range(len(string) - 1, -1, -1):
        if string[i] in digits:
            number = string[i] + number
        else:
            return string[0:i + 1], int(number)


# A helper function for Calc searching.
def inside_base26_to_decimal(start, end):
    """A supplementary function for field_determine"""
    base_number = 1
    end_value = 0
    start_value = 0
    while len(end) != 0:
        end_value += (ord(end[-1]) - 64) * base_number
        base_number *= 26
        end = end[:-1]
    base_number = 1
    while len(start) != 0:
        start_value += (ord(start[-1]) - 64) * base_number
        base_number *= 26
        start = start[:-1]
    return start_value, end_value


# A helper function for Calc searching.
def inside_base26(x):
    """Converts the value of the alphabetic address to letters back."""
    number = ""
    while x > 0:
        modulo = (x - 1) % 26
        number = chr(modulo + 65) + number
        x = (x-modulo)//26
    return number


# A helper function for Calc searching.
def inside_field_determine(starting, ending):
    """Assumes starting and ending are addresses of cells
    :returns cells' addresses between starting and ending."""
    start, end = inside_address_spliter(starting), inside_address_spliter(ending)
    decimal_value = inside_base26_to_decimal(start[0], end[0])
    address_list = []
    for i in range(end[1] - start[1] + 1):
        for j in range(decimal_value[0], decimal_value[1] + 1):
            address_list.append((inside_base26(j)) + str(start[1] + i))
    return address_list


def calc_str_exists_in_cells(start_addr, end_addr, search_term, sheet):
    """Assumes starting and ending are addresses of cells, and search_term is a string
    :returns True if search_term exits between start_addr and end_addr, false otherwise """
    scope = inside_field_determine(start_addr, end_addr)
    for i in scope:
        inner_cell = sheet.getCellRangeByName(i)
        if inner_cell.String == search_term:
            return True
    return False


def calc_value_exists_in_cells(start_addr, end_addr, value, sheet):
    """Assumes starting and ending are adresses of cells, and search_term is a value
    :returns True if search_term exits between start_addr and end_addr, false otherwise """
    scope = inside_field_determine(start_addr, end_addr)
    for i in scope:
        inner_cell = sheet.getCellRangeByName(i)
        if inner_cell.Value == value:
            return True
    return False


def calc_search_str_in_cells(start_addr, end_addr, search_term, sheet):
    """Assumes starting and ending are adresses of cells, and search_term is a string
    :returns first occurrence of search_term's address if search_term exits between start_addr and end_addr, false otherwise """
    scope = inside_field_determine(start_addr, end_addr)
    for i in scope:
        inner_cell = sheet.getCellRangeByName(i)
        if inner_cell.String == search_term:
            return i
    return False


def calc_search_value_in_cells(start_addr, end_addr, value, sheet):
    """Assumes starting and ending are adresses of cells, and search_term is a value
    :returns first occurrence of search_term's address if search_term exits between start_addr and end_addr, false otherwise """
    scope = inside_field_determine(start_addr, end_addr)
    for i in scope:
        inner_cell = sheet.getCellRangeByName(i)
        if inner_cell.Value == value:
            return i
    return False


def calc_get_cell_text_with_addr(cell_addr, sheet):
    """Assumes cell_addr is an address of a cell
    returns cell_addr's content as a string """
    inner_cell = sheet.getCellRangeByName(cell_addr)
    return inner_cell.String


def calc_set_cell_text_with_addr(cell_addr, new_str, sheet):
    """Assumes cell_addr is an address of a cell and new_str is a string
    sets cell_addr to new_str """
    inner_cell = sheet.getCellRangeByName(cell_addr)
    inner_cell.String = new_str


def calc_text_replace_all(active_sheet, find_text, replace_with, case_sensitive=False, whole_words=False):
    # Create an object from find_text and specify its properties.
    search = active_sheet.createSearchDescriptor()
    search.SearchString = find_text
    search.SearchWords = whole_words
    search.SearchCaseSensitive = case_sensitive
    # Find the first occurrence of find_text
    found = active_sheet.findFirst(search)
    if found is None:
        return False
    # Replace all occurrences.
    while found is not None:
        found.String = replace_with
        search = active_sheet.createSearchDescriptor()
        search.SearchString = find_text
        search.SearchWords = whole_words
        search.SearchCaseSensitive = case_sensitive
        found = active_sheet.findFirst(search)
    return True


def writer_text_find(model, find_text, case_sensitive=False):
    # find text in the document considering case sensitivity as given by case-sensitive field, and select the found text
    # default case sensitive: false
    # Select the found text with cursor.
    search = model.createSearchDescriptor()
    search.SearchString = find_text
    if case_sensitive:
        search.SearchCaseSensitive = True
    found = model.findFirst(search)
    if found is not None:
        # oVC means to View Cursor Object.
        viewCursor = model.getCurrentController()
        oVC = viewCursor.getViewCursor()
        # False means do not select text between the current position of the cursor and found.
        oVC.gotoRange(found, False)
        return True
    else:
        return False


def writer_text_replace_all(model, find_text, replace_with, case_sensitive=False, whole_words=False):
    """Replaces all find_text with replace_with in document"""
    # Create an object from find_text and specify its properties.
    search = model.createSearchDescriptor()
    search.SearchString = find_text
    search.SearchWords = whole_words
    search.SearchCaseSensitive = case_sensitive
    # Find the first occurrence of find_text
    found = model.findFirst(search)
    if found is None:
        return False
    # Replace all occurrences.
    while found is not None:
        found.String = replace_with
        search = model.createSearchDescriptor()
        search.SearchString = find_text
        search.SearchWords = whole_words
        search.SearchCaseSensitive = case_sensitive
        found = model.findFirst(search)
    return True


def writer_text_insert(model, content):
    text = model.Text
    viewCursor = model.getCurrentController()
    # oVC means to View Cursor Object.
    oVC = viewCursor.getViewCursor()
    if oVC.isCollapsed():
        text.insertString(oVC, content, 0)
    else:
        oVC.String = ""
        text = model.Text
        text.insertString(oVC, content, 0)


def writer_image_insert_from_file(model, file_path):
    fileURL = uno.systemPathToFileUrl(file_path)
    text = model.getText()
    oCursor = text.createTextCursor()
    oGraph = model.createInstance("com.sun.star.text.GraphicObject")
    oGraph.GraphicURL = fileURL
    oGraph.AnchorType = AS_CHARACTER
    # Set the image size.
    oGraph.Width = 10000
    oGraph.Height = 8000
    text.insertTextContent(oCursor, oGraph, False)


def dispatcher(context):
    """You can find context at second value of value list of start_libreoffice() """
    serviceManager = context.ServiceManager
    return serviceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", context)


def impress_text_search_dispatcher(model, dispatcher, find_text, case_sensitive=False, whole_words=False):
    """Model, dispatcher, and find_text are required"""
    pages = model.getDrawPages()
    # Search every page individually.
    for selected_page in pages:
        search = selected_page.createSearchDescriptor()
        search.SearchString = find_text
        search.SearchWords = whole_words
        search.SearchCaseSensitive = case_sensitive
        found = selected_page.findFirst(search)
        # Find the first occurrence and select it with cursor.
        if found is not None:
            currentController = model.getCurrentController()
            struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
            struct.Name = "SearchItem.SearchString"
            struct.Value = find_text
            dispatcher.executeDispatch(currentController, ".uno:ExecuteSearch", "", 0, tuple([struct]))
            return True
    return False


def impress_text_replace_all(model, find_text, replace_with, case_sensitive=False, whole_words=False):
    """Replaces all find_text with replace_with in document"""
    pages = model.getDrawPages()
    # Search every page individually.
    for selected_page in pages:
        replace = selected_page.createReplaceDescriptor()
        replace.setSearchString(find_text)
        replace.setReplaceString(replace_with)
        replace.SearchCaseSensitive = case_sensitive
        replace.SearchWords = whole_words
        selected_page.replaceAll(replace)


def calc_dispatcher_example(model, dispatcher):
    # An example function for how to use dispatcher in calc.
    currentController = model.getCurrentController()
    struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    struct.Name = 'ToPoint'
    struct.Value = 'Sheet1.A1'
    dispatcher.executeDispatch(currentController, ".uno:GoToCell", "", 0, tuple([struct]))


def macro_wiper(document_name, macro_name, function_name):
    """document_name is the name of file which contains embedded macros like "calc_example.ods",macro_name is the name
    of embedded macro like "myscript.py" function_name is the name of the funtion to be deleted like "function"
     """
    # This function should be written again.
    import zipfile
    import shutil
    import os
    shutil.rmtree("without_macro", True)
    os.mkdir("without_macro")
    filename = "without_macro/" + document_name
    shutil.copyfile(document_name, filename)
    exporting = zipfile.ZipFile(filename, 'a')
    myfile = exporting.open("Scripts/python/" + macro_name, 'r')
    new = open("temp.py", "wb")
    for aline in myfile.readlines():
        new.write(aline)
    new.close()
    exporting.close()
    editing = open("temp.py", "r")
    importing = open(macro_name, "w")
    out_of_function = 1
    deleted = 0
    funcLen = len(function_name)
    for line in editing.readlines():
        line_list = line.split()
        try:
            if line_list[0] == "def" and line_list[1][:funcLen] == function_name:
                deleted = 1
                out_of_function = 0
                continue
            if deleted and line_list[0] == "def":
                out_of_function = 1
            if out_of_function:
                importing.write(line)
        except IndexError:
            continue
    editing.close()
    importing.close()
    os.remove("temp.py")
    doc = zipfile.ZipFile(filename, 'a')
    doc.write(macro_name, "Scripts/python/" + macro_name)
    manifest = []
    for line in doc.open('META-INF/manifest.xml'):
        if '</manifest:manifest>' in line.decode('utf-8'):
            for path in ['Scripts/', 'Scripts/python/', 'Scripts/python/' + macro_name]:
                manifest.append(' <manifest:file-entry manifest:media-type="application/binary" manifest:full-path="%s"/>' % path)
        manifest.append(line.decode('utf-8'))
    doc.writestr('META-INF/manifest.xml', ''.join(manifest))
    doc.close()
    os.remove(macro_name)


def delete_all_macros(document_name):
    """Deletes all embedded macros from the given libreoffice document."""
    import zipfile
    import shutil
    import os
    shutil.rmtree("without_macros", True)
    os.mkdir("without_macros")
    filename = "without_macros/" + document_name
    zin = zipfile.ZipFile(document_name, 'r')
    zout = zipfile.ZipFile(filename, 'w')
    for item in zin.infolist():
        buffer = zin.read(item.filename)
        if item.filename[-3:] != '.py':
            zout.writestr(item, buffer)
    zout.close()
    zin.close()
