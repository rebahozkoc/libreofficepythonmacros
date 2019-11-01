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


# get the uno component context from the PyUNO runtime


# create the UnoUrlResolver

# connect to the running office
# ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
# smgr = ctx.ServiceManager

# get the central desktop object
# desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)

# access the current document/component module
# model = desktop.getCurrentComponent()

# access the active sheet
# active_sheet = model.CurrentController.ActiveSheet


# def start_libreoffice():
#    subprocess.Popen("soffice --calc --accept=\"socket,host=localhost,port=2002;urp;StarOffice.ServiceManager\"", shell=True)
#    time.sleep(1)

def start_libreoffice():
    sofficePath = '/usr/bin/soffice'
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
    return context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context)


def new_writer_doc(desktop):
    return desktop.loadComponentFromURL('private:factory/swriter', 'blank', 0, ())


def new_calc_doc(desktop):
    return desktop.loadComponentFromURL('private:factory/scalc', 'blank', 0, ())


def new_impress_doc(desktop):
    return desktop.loadComponentFromURL('private:factory/simpress', 'blank', 0, ())


# path = uno.fileUrlToSystemPath(url) may help later

def open_doc(desktop, path1):
    FileURL = uno.systemPathToFileUrl(path1)
    return desktop.loadComponentFromURL(FileURL,"_blank",0,())


#Old version of connection function
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


def inside_base26(x):
    """A supplementary function for field_determine"""
    number = ""
    while 0 > x > 26:
        number = chr(x % 26 + 64) + number
        x = x // 26
    return chr(x + 64) + number


def inside_field_determine(starting, ending):
    """Assumes starting and ending are adresses of cells
    :returns cells' addresses between starting and ending"""
    start, end = inside_address_spliter(starting), inside_address_spliter(ending)
    decimal_value = inside_base26_to_decimal(start[0], end[0])
    address_list = []
    for i in range(end[1] - start[1] + 1):
        for j in range(decimal_value[0], decimal_value[1] + 1):
            address_list.append((inside_base26(j)) + str(start[1] + i))
    return address_list


def calc_str_exists_in_cells(start_addr, end_addr, search_term, sheet):
    """Assumes starting and ending are adresses of cells, and search_term is a string
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
    returns cell_addr's content as string """
    inner_cell = sheet.getCellRangeByName(cell_addr)
    return inner_cell.String


def calc_set_cell_text_with_addr(cell_addr, new_str, sheet):
    """Assumes cell_addr is an address of a cell and new_str is a string
    sets cell_addr to new_str """
    inner_cell = sheet.getCellRangeByName(cell_addr)
    inner_cell.String = new_str


def writer_text_find(doc, find_text, case_sensitive = False):
    # find text in the document considering case sentivity as given by case-sentive field, and select the found text
    #default case sensitive: false
    #Select the found text with cursor.
    search = doc.createSearchDescriptor()
    search.SearchString = find_text
    if case_sensitive:
        search.SearchCaseSensitive = True
    found = doc.findFirst(search)
    if found is not None:
        #print("found",found)
        viewCursor = doc.getCurrentController()
        oVC = viewCursor.getViewCursor()
        oVC.gotoRange(found, False)
        return True
    else:
        return False


def writer_text_replace_all(doc, find_text, replace_with, case_sensitive = False, whole_words = False):
    """Replaces all find_text with replace_with in document"""
    search = doc.createSearchDescriptor()
    search.SearchString = find_text
    search.SearchWords = whole_words
    search.SearchCaseSensitive = case_sensitive
    found = doc.findFirst(search)
    if found is None:
        return False
    while found is not None:
        found.String = replace_with
        search = doc.createSearchDescriptor()
        search.SearchString = find_text
        search.SearchWords = whole_words
        search.SearchCaseSensitive = case_sensitive
        found = doc.findFirst(search)
    return True



def writer_text_insert(doc, content):
    text = doc.Text
    viewCursor = doc.getCurrentController()
    oVC = viewCursor.getViewCursor()
    if oVC.isCollapsed():
        text.insertString(oVC, content, 0)
    else:
        oVC.String = ""
        text = doc.Text
        text.insertString(oVC, content, 0)


def writer_image_insert_from_file(doc,file_path):
    FileURL = uno.systemPathToFileUrl(file_path)
    text = doc.getText()
    oCursor = text.createTextCursor()
    oGraph = doc.createInstance("com.sun.star.text.GraphicObject")
    oGraph.GraphicURL = FileURL
    oGraph.AnchorType = AS_CHARACTER
    oGraph.Width = 10000
    oGraph.Height = 8000
    text.insertTextContent(oCursor, oGraph, False)

