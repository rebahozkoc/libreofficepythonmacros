import subprocess
import time
#import helper

#Old version of start_libreoffice, works but does not support multiple files.
#def start_libreoffice():
#    a = subprocess.Popen("soffice --calc --accept=\"socket,host=localhost,port=2002;urp;StarOffice.ServiceManager\"", shell=True)

# #def test_open_new_files(desktop):
#     try:
#         calc = helper.new_calc_doc(desktop)
#         print("calc is started.")
#         assert calc is not None
#         time.sleep(0.5)
#         calc.close(False)
#         print("calc is closed.")
#         time.sleep(1)
#     except:
#         print("ERROR: Calc could not be started.")
#         return False
#     try:
#         impress = helper.new_impress_doc(desktop)
#         print("impress is started.")
#         assert impress is not None
#         time.sleep(0.5)
#         impress.close(False)
#         print("impress is closed.")
#         time.sleep(1)
#     except:
#         print("ERROR: Impress could not be started.")
#         return False
#     try:
#         writer = helper.new_writer_doc(desktop)
#         print("writer is started.")
#         assert writer is not None
#         time.sleep(0.5)
#         writer.close(False)
#         print("writer is closed.")
#         time.sleep(1)
#     except:
#         print("ERROR: writer could not be started.")
#         return False

def documentation(x):
    return x.replace(",",",\n")



print(documentation("oview cursor: pyuno object (com.sun.star.frame.XController)0x1ba7b88{implementationName=DrawController, supportedServices={com.sun.star.drawing.DrawingDocumentDrawView}, supportedInterfaces={com.sun.star.frame.XController2,com.sun.star.frame.XControllerBorder,com.sun.star.frame.XDispatchProvider,com.sun.star.task.XStatusIndicatorSupplier,com.sun.star.ui.XContextMenuInterception,com.sun.star.awt.XUserInputInterception,com.sun.star.frame.XDispatchInformationProvider,com.sun.star.frame.XTitle,com.sun.star.frame.XTitleChangeBroadcaster,com.sun.star.lang.XInitialization,com.sun.star.lang.XTypeProvider,com.sun.star.uno.XWeak,com.sun.star.beans.XMultiPropertySet,com.sun.star.beans.XFastPropertySet,com.sun.star.beans.XPropertySet,com.sun.star.view.XSelectionSupplier,com.sun.star.lang.XServiceInfo,com.sun.star.drawing.XDrawView,com.sun.star.view.XSelectionChangeListener,com.sun.star.view.XFormLayerAccess,com.sun.star.drawing.framework.XControllerManager,com.sun.star.lang.XUnoTunnel,com.sun.star.lang.XTypeProvider,com.sun.star.frame.XController2,com.sun.star.frame.XControllerBorder,com.sun.star.frame.XDispatchProvider,com.sun.star.task.XStatusIndicatorSupplier,com.sun.star.ui.XContextMenuInterception,com.sun.star.awt.XUserInputInterception,com.sun.star.frame.XDispatchInformationProvider,com.sun.star.frame.XTitle,com.sun.star.frame.XTitleChangeBroadcaster,com.sun.star.lang.XInitialization,com.sun.star.lang.XTypeProvider,com.sun.star.uno.XWeak}}"))
