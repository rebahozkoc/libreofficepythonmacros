import socket
import uno
import os
from sys import platform
import random
import subprocess
import time

# get the uno component context from the PyUNO runtime


# create the UnoUrlResolver

# connect to the running office
#ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
#smgr = ctx.ServiceManager

# get the central desktop object
#desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)

# access the current document/component module
#model = desktop.getCurrentComponent()

# access the active sheet
#active_sheet = model.CurrentController.ActiveSheet


#def start_libreoffice():
#    subprocess.Popen("soffice --calc --accept=\"socket,host=localhost,port=2002;urp;StarOffice.ServiceManager\"", shell=True)
#    time.sleep(1)

def start_libreoffice():
    print("rehrefg")


def connect_to_libreoffice():
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext )
    return resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")


def get_current_doc(ctx1):
    smgr = ctx1.ServiceManager
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx1)
    return desktop.getCurrentComponent()


def get_active_sheet(model1):
    return model1.CurrentController.ActiveSheet


#active_sheet = get_active_sheet(get_current_doc(connect_to_libreoffice()))


def get_cell_from_sheet(cell_name, sheet):
    return sheet.getCellRangeByName(cell_name)


def set_cell_text(cell1,text):
    cell1.String = text


def get_cell_text(cell):
    return cell.String


def address_spliter(string):
    """Assumes string consist with letters and numbers, respectively.
     :returns a tuple whose first element is letters second is numbers"""
    digits = ['0','1','2','3','4','5','6','7','8','9']
    number = ""
    for i in range(len(string)-1,-1,-1):
        if string[i] in digits:
            number = string[i] + number
        else:
            return string[0:i+1], int(number)

def base26_to_decimal(start,end):
    """A supplementary function for field_determine"""
    base_number = 1
    end_value = 0
    start_value = 0
    while len(end) != 0:
        end_value += (ord(end[-1])-64)*base_number
        base_number *= 26
        end = end[:-1]
    base_number = 1
    while len(start) != 0:
        start_value += (ord(start[-1])-64)*base_number
        base_number *= 26
        start = start[:-1]
    return start_value, end_value

def base26(x):
    """A supplementary function for field_determine"""
    number = ""
    while x>26:
        number = chr(x%26+64) + number
        x = x//26
    return chr(x+64) + number

def field_determine(starting,ending):
    """Assumes starting and ending are adresses of cells
    :returns cells' addresses between starting and ending"""
    start,end = address_spliter(starting),address_spliter(ending)
    decimal_value = base26_to_decimal(start[0],end[0])
    address_list = []
    for i in range(end[1]-start[1]+1):
        for j in range(decimal_value[0], decimal_value[1]+1):
            address_list.append((base26(j))+str(start[1]+i))
    return address_list


def str_exists_in_cells(start_addr,end_addr, search_term):
    """Assumes starting and ending are adresses of cells, and search_term is a string
    :returns True if search_term exits between start_addr and end_addr, false otherwise """
    scope = field_determine(start_addr,end_addr)
    for i in scope:
        inner_cell = active_sheet.getCellRangeByName(i)
        if inner_cell.String == search_term:
            return True
    return False


def value_exists_in_cells(start_addr,end_addr, value):
    """Assumes starting and ending are adresses of cells, and search_term is a value
    :returns True if search_term exits between start_addr and end_addr, false otherwise """
    scope = field_determine(start_addr,end_addr)
    for i in scope:
        inner_cell = active_sheet.getCellRangeByName(i)
        if inner_cell.Value == value:
            return True
    return False


def search_str_in_cells(start_addr,end_addr, search_term):
    """Assumes starting and ending are adresses of cells, and search_term is a string
    :returns first occurrence of search_term's address if search_term exits between start_addr and end_addr, false otherwise """
    scope = field_determine(start_addr,end_addr)
    for i in scope:
        inner_cell = active_sheet.getCellRangeByName(i)
        if inner_cell.String == search_term:
            return i
    return False


def search_value_in_cells(start_addr,end_addr, value):
    """Assumes starting and ending are adresses of cells, and search_term is a value
    :returns first occurrence of search_term's address if search_term exits between start_addr and end_addr, false otherwise """
    scope = field_determine(start_addr,end_addr)
    for i in scope:
        inner_cell = active_sheet.getCellRangeByName(i)
        if inner_cell.Value == value:
            return i
    return False


def get_cell_text_with_addr(cell_addr):
    """Assumes cell_addr is an address of a cell
    returns cell_addr's content as string """
    inner_cell = active_sheet.getCellRangeByName(cell_addr)
    return inner_cell.String


def set_cell_text_with_addr(cell_addr,new_str):
    """Assumes cell_addr is an address of a cell and new_str is a string
    sets cell_addr to new_str """
    inner_cell = active_sheet.getCellRangeByName(cell_addr)
    inner_cell.String = new_str
