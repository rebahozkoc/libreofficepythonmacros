import subprocess
import time


#Old version of start_libreoffice, works but does not support multiple files.
def start_libreoffice():
    a = subprocess.Popen("soffice --calc --accept=\"socket,host=localhost,port=2002;urp;StarOffice.ServiceManager\"", shell=True)

    time.sleep(1)
    print("Python done")
    print("out of starting")


