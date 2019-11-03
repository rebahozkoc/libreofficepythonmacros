import subprocess
import time
import helper


#Old version of start_libreoffice, works but does not support multiple files.
#def start_libreoffice():
#    a = subprocess.Popen("soffice --calc --accept=\"socket,host=localhost,port=2002;urp;StarOffice.ServiceManager\"", shell=True)

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


def deneme(*args):
    a = args[0:2]
    print(*args)
    print("args:",args)
    print("*a:",*a)




deneme(1,3243,5,2445,65,4,43,7)
