import uno

class OfficeMacro:
    def __init__(self):
        import os
        import tempfile
        import shutil
        import uuid
        import subprocess
        import uno
        from com.sun.star.connection import NoConnectException
        import time
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
        self.desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
        self.context = context

    def getDesktop(self):
        return self.desktop

    def getContext(self):
        return self.context

    def getDispatcher(self):
        serviceManager = self.getContext().ServiceManager
        return serviceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", self.getContext())

class DocumentTypeError(Exception):
    """Wrong document type is given."""
    # Raise when a object is tried to construct with a wrong document type.
    def __init__(self):
        Exception.__init__(self, "Wrong document type is given.")


class SpreadSheet(OfficeMacro):
    def __init__(self, document_name=None):
        OfficeMacro.__init__(self)
        # Open a blank document.
        if document_name is None:
            self.model = self.getDesktop().loadComponentFromURL('private:factory/scalc', 'blank', 0, ())
        # Check document type and open.
        elif "ods" in document_name:
            fileURL = uno.systemPathToFileUrl(document_name)
            self.model = self.getDesktop().loadComponentFromURL(fileURL, "_blank", 0, ())
        else:
            raise DocumentTypeError()

    def getModel(self):
        return self.model

    # Close the existing file.
    def closeFile(self, arg=False):
        self.getModel().close(arg)

    # Return active sheet.
    def getActiveSheet(self):
        return self.getModel().CurrentController.ActiveSheet

    # Return cell from active sheet by cell name.
    def getCellByName(self, cell_name):
        return self.getActiveSheet().getCellRangeByName(cell_name)

    # Write the text to given cell as cell adress.
    def setCellText(self, cell, text):
        """Assumes cell is an address of a cell and text is a string
         writes text to cell  """
        self.getCellByName(cell).String = text


    def getCellText(self, cell_name):
        """Assumes cell_name is an address of a cell
        returns cell_name's content as a string
        """
        return self.getCellByName(cell_name).String

    @staticmethod
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

    @staticmethod
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

    @staticmethod
    # A helper function for Calc searching.
    def inside_base26(x):
        """Converts the value of the alphabetic address to letters back."""
        number = ""
        while x > 0:
            modulo = (x - 1) % 26
            number = chr(modulo + 65) + number
            x = (x-modulo)//26
        return number

    @staticmethod
    # A helper function for Calc searching.
    def inside_field_determine(starting, ending):
        """Assumes starting and ending are addresses of cells
        :returns cells' addresses between starting and ending."""
        start, end = SpreadSheet.inside_address_spliter(starting), SpreadSheet.inside_address_spliter(ending)
        decimal_value = SpreadSheet.inside_base26_to_decimal(start[0], end[0])
        address_list = []
        for i in range(end[1] - start[1] + 1):
            for j in range(decimal_value[0], decimal_value[1] + 1):
                address_list.append((SpreadSheet.inside_base26(j)) + str(start[1] + i))
        return address_list

    def checkTextExists(self, start_addr, end_addr, search_term):
        """Assumes starting and ending are addresses of cells, and search_term is a string
        :returns True if search_term exits between start_addr and end_addr, false otherwise """
        scope = SpreadSheet.inside_field_determine(start_addr, end_addr)
        for i in scope:
            inner_cell = self.getCellByName(i)
            if inner_cell.String == search_term:
                return True
        return False

    def checkValueExists(self, start_addr, end_addr, value):
        """Assumes starting and ending are adresses of cells, and search_term is a value
        :returns True if search_term exits between start_addr and end_addr, false otherwise """
        scope = SpreadSheet.inside_field_determine(start_addr, end_addr)
        for i in scope:
            inner_cell = self.getCellByName(i)
            if inner_cell.Value == value:
                return True
        return False

    def getStrAdressInCells(self, start_addr, end_addr, search_term):
        """Assumes starting and ending are adresses of cells, and search_term is a string
        :returns first occurrence of search_term's address if search_term exits between start_addr and end_addr, false otherwise """
        scope = SpreadSheet.inside_field_determine(start_addr, end_addr)
        for i in scope:
            inner_cell = self.getCellByName(i)
            if inner_cell.String == search_term:
                return i
        return False

    def getValueAdressInCells(self, start_addr, end_addr, value):
        """Assumes starting and ending are adresses of cells, and search_term is a string
        :returns first occurrence of search_term's address if search_term exits between start_addr and end_addr, false otherwise """
        scope = SpreadSheet.inside_field_determine(start_addr, end_addr)
        for i in scope:
            inner_cell = self.getCellByName(i)
            if inner_cell.Value == value:
                return i
        return False


    def textReplaceAll(self, find_text, replace_with, case_sensitive=False, whole_words=False):
        # Create an object from find_text and specify its properties.
        active_sheet = self.getActiveSheet()
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

    def insertText(self, cell_name, text):
        """Insert text in a cell after existing text."""
        existing_text = self.getCellText(cell_name)
        existing_text = existing_text + text
        self.setCellText(cell_name, existing_text)
        return existing_text


class WordProcessor(OfficeMacro):
    def __init__(self, document_name=None):
        OfficeMacro.__init__(self)
        # Open a blank document.
        if document_name is None:
            self.model = self.getDesktop().loadComponentFromURL('private:factory/swriter', 'blank', 0, ())
        # Check document type and open.
        elif "odt" in document_name:
            fileURL = uno.systemPathToFileUrl(document_name)
            self.model = self.getDesktop().loadComponentFromURL(fileURL, "_blank", 0, ())
        else:
            raise DocumentTypeError()

    def getModel(self):
        return self.model

    # Close the opened file.
    def closeFile(self, arg=False):
        self.getModel().close(arg)

    def checkTextExists(self, search_term, case_sensitive=False):
        """find text in the document considering case sensitivity as given by case-sensitive field, and select the found text
        default case sensitive: false
        Select the found text with cursor.
        """
        model = self.getModel()
        search = model.createSearchDescriptor()
        search.SearchString = search_term
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


    def textReplaceAll(self, search_term, replace_with, case_sensitive=False, whole_words=False):
        """Replaces all find_text with replace_with in document"""
        # Create an object from find_text and specify its properties.
        model = self.getModel()
        replace = model.createReplaceDescriptor()
        replace.setSearchString(search_term)
        replace.setReplaceString(replace_with)
        replace.SearchCaseSensitive = case_sensitive
        replace.SearchWords = whole_words
        model.replaceAll(replace)


    def insertText(self, text):
        """ Insert text at the beginning of file."""
        model = self.getModel()
        text = model.Text
        viewCursor = model.getCurrentController()
        # oVC means to View Cursor Object.
        oVC = viewCursor.getViewCursor()
        if oVC.isCollapsed():
            text.insertString(oVC, text, 0)
        else:
            oVC.String = ""
            text = model.Text
            text.insertString(oVC, text, 0)

    def insertImage(self, file_path):
        from com.sun.star.text.TextContentAnchorType import AS_CHARACTER
        fileURL = uno.systemPathToFileUrl(file_path)
        model = self.getModel()
        text = model.getText()
        oCursor = text.createTextCursor()
        oGraph = model.createInstance("com.sun.star.text.GraphicObject")
        oGraph.GraphicURL = fileURL
        oGraph.AnchorType = AS_CHARACTER
        # Set the image size.
        oGraph.Width = 10000
        oGraph.Height = 8000
        text.insertTextContent(oCursor, oGraph, False)

class Presentation(OfficeMacro):
    def __init__(self, document_name=None):
        OfficeMacro.__init__(self)
        # Open a blank document.
        if document_name is None:
            self.model = self.getDesktop().loadComponentFromURL('private:factory/simpress', 'blank', 0, ())
        # Check document type and open.
        elif "odp" in document_name:
            fileURL = uno.systemPathToFileUrl(document_name)
            self.model = self.getDesktop().loadComponentFromURL(fileURL, "_blank", 0, ())
        else:
            raise DocumentTypeError()

    def getModel(self):
        return self.model

    # Close the existing file.
    def closeFile(self, arg=False):
        self.getModel().close(arg)

    def checkTextExists(self, search_term, case_sensitive=False, whole_words=False):
        """find text in the document
        Select the found text with cursor.
        default case sensitivity: false
        default whole word sensitivity: false
        """
        model = self.getModel()
        pages = model.getDrawPages()
        # Search every page individually.
        for selected_page in pages:
            search = selected_page.createSearchDescriptor()
            search.SearchString = search_term
            search.SearchWords = whole_words
            search.SearchCaseSensitive = case_sensitive
            found = selected_page.findFirst(search)
            # Find the first occurrence and select it with cursor.
            if found is not None:
                currentController = model.getCurrentController()
                struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
                struct.Name = "SearchItem.SearchString"
                struct.Value = search_term
                self.getDispatcher().executeDispatch(currentController, ".uno:ExecuteSearch", "", 0, tuple([struct]))
                return True
        return False

    def textReplaceAll(self, search_term, replace_with, case_sensitive=False, whole_words=False):
        """Replaces all find_text with replace_with in document"""
        # Create an object from find_text and specify its properties.
        model = self.getModel()
        pages = model.getDrawPages()
        # Search every page severally.
        for selected_page in pages:
            replace = selected_page.createReplaceDescriptor()
            replace.setSearchString(search_term)
            replace.setReplaceString(replace_with)
            replace.SearchCaseSensitive = case_sensitive
            replace.SearchWords = whole_words
            selected_page.replaceAll(replace)
