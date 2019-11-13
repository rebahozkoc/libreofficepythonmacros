

#Old version of start_libreoffice, works but does not support multiple files.
#def start_libreoffice():
#    a = subprocess.Popen("soffice --calc --accept=\"socket,host=localhost,port=2002;urp;StarOffice.ServiceManager\"", shell=True)


def documentation(x):
    return x.replace(",",",\n")


del(documentation)

print(documentation("ere,df,ggedf, fdg"))


