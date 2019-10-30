import subprocess
import time


#Old version of start_libreoffice, works but does not support multiple files.
#def start_libreoffice():
#    a = subprocess.Popen("soffice --calc --accept=\"socket,host=localhost,port=2002;urp;StarOffice.ServiceManager\"", shell=True)
#    time.sleep(1)

def deneme(a,b=False):
    if b:
        print("if is true",a)
    else:
        print("if is false",a)

deneme(21)
