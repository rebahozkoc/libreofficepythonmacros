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


class DocumentTypeError(Exception):
    """Wrong document type is given."""
    # Raise when a object is tried to construct with a wrong document type.
    def __init__(self):
        Exception.__init__(self, "Wrong document type is given.")



class WordProcessor(OfficeMacro):
    def __init__(self, document_name=None):
        OfficeMacro.__init__(self)
        # Open a blank document.
        if document_name is None:
            self.model = self.getDesktop().loadComponentFromURL('private:factory/swriter', 'blank', 0, ())
        # Check document type and open.
        elif "odt" in document_name:
            import uno
            fileURL = uno.systemPathToFileUrl(document_name)
            self.model = self.getDesktop().loadComponentFromURL(fileURL, "_blank", 0, ())
        else:
            raise DocumentTypeError()

    def getModel(self):
        return self.model


class SpreadSheet(OfficeMacro):
    def __init__(self, document_name=None):
        OfficeMacro.__init__(self)
        # Open a blank document.
        if document_name is None:
            self.model = self.getDesktop().loadComponentFromURL('private:factory/scalc', 'blank', 0, ())
        # Check document type and open.
        elif "ods" in document_name:
            import uno
            fileURL = uno.systemPathToFileUrl(document_name)
            self.model = self.getDesktop().loadComponentFromURL(fileURL, "_blank", 0, ())
        else:
            raise DocumentTypeError()

    def getModel(self):
        return self.model


class Presentation(OfficeMacro):
    def __init__(self, document_name=None):
        OfficeMacro.__init__(self)
        # Open a blank document.
        if document_name is None:
            self.model = self.getDesktop().loadComponentFromURL('private:factory/simpress', 'blank', 0, ())
        # Check document type and open.
        elif "odp" in document_name:
            import uno
            fileURL = uno.systemPathToFileUrl(document_name)
            self.model = self.getDesktop().loadComponentFromURL(fileURL, "_blank", 0, ())
        else:
            raise DocumentTypeError()

    def getModel(self):
        return self.model





