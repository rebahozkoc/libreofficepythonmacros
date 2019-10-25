import os
import shutil
import subprocess
import tempfile
import time
import uno
import uuid

from com.sun.star.connection import NoConnectException

sofficePath = '/usr/bin/soffice'
tempDir = tempfile.mkdtemp()

# Restore cached profile if available
userProfile = tempDir + '/profile'
cacheDir = os.getenv('XDG_CACHE_DIR', os.environ['HOME'] + '/.cache')+'/lo_profile_cache'
if os.path.isdir(cacheDir):
    shutil.copytree(cacheDir, userProfile)
    profileCached = True
else:
    os.mkdir(userProfile)
    profileCached = False

# Launch LibreOffice server
pipeName = uuid.uuid4().hex
args = [
    sofficePath,
    '-env:UserInstallation=file://' + userProfile,
    '--pidfile=' + tempDir + '/soffice.pid',
    '--accept=pipe,name=' + pipeName + ';urp;',
    '--norestore',
    '--invisible'
]
sofficeEnvironment = os.environ
sofficeEnvironment['TMPDIR'] = tempDir
proc = subprocess.Popen(args, env=sofficeEnvironment, preexec_fn=os.setsid)

# Open connection to server
for i in range(100):
    try:
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",
localContext)
        context =resolver.resolve("uno:pipe,name=%s;urp;StarOffice.ComponentContext" %
pipeName)
        break
    except NoConnectException:
        time.sleep(0.1)
        if i == 99:
            raise

# Cache profile if required
if not profileCached:
    shutil.copytree(userProfile, cacheDir)


desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context)

#Buradan yukarısı bağlantı
doc = desktop.loadComponentFromURL('private:factory/swriter', 'blank', 0, ())

#doc = desktop.loadComponentFromURL('private:factory/calc', 'blank', 0, ())

#doc = component.getDocument()

text = doc.getText()

text.setString("Hello world")
