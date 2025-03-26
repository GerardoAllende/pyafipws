import pythoncom
import win32com.client

def print_dummy(obj):
    obj.Dummy()
    print("Dummy: ", obj.AppServerStatus, obj.DbServerStatus, obj.AuthServerStatus)

def com_test():
    """
    Test COM object creation and mathod calling
    """
    wsaa = win32com.client.Dispatch("WSAA")
    print("WSAA", wsaa.Version, wsaa.InstallDir)
    
    wsfev1 = win32com.client.Dispatch("WSFEv1")
    print("\nWSFEv1", wsfev1.Version)
    print("Conectar: ", wsfev1.Conectar())
    print_dummy(wsfev1)
    
    # ~ wsbfev1 = win32com.client.Dispatch("WSBFEv1")
    # ~ print("\nWSBFEv1", wsbfev1.Version)
    # ~ print("Conectar: ", wsbfev1.Conectar())
    # ~ print_dummy(wsbfev1)
    
    # ~ wsfexv1 = win32com.client.Dispatch("WSFEXv1")
    # ~ print("\nWSFEXv1", wsfexv1.Version)
    # ~ print("Conectar: ", wsfexv1.Conectar())
    # ~ print_dummy(wsfexv1)
    
    # ~ wsmtx = win32com.client.Dispatch("WSMTXCA")
    # ~ print("\nWSMTXCA", wsmtx.Version)
    # ~ print("Conectar: ", wsmtx.Conectar())
    
    # ~ wsct = win32com.client.Dispatch("WSCT")
    # ~ print("\nWSCT", wsct.Version)
    # ~ print("Conectar: ", wsct.Conectar())
    
    # ~ cot = win32com.client.Dispatch("COT")
    # ~ print("\nCOT", cot.Version)
    
    padron = win32com.client.Dispatch("WSSrPadronA5")
    print("\nWSSrPadronA5", padron.Version)
    
    
    pyqr = win32com.client.Dispatch("PyQR")
    print("\nPyQR", pyqr.Version)
    
    pyfepdf = win32com.client.Dispatch("PyFEPDF")
    print("\nPyFEPDF", pyfepdf.Version)
    
    # ~ cred = win32com.client.Dispatch("WSFECred")
    # ~ print("\nWSFECred", cred.Version)
    # ~ print("Conectar: ", cred.Conectar())
    # ~ print_dummy(cred)
    input("Presione intro para salir")

    
    

if __name__ == "__main__":
    # Initialize COM
    pythoncom.CoInitialize()

    # Run the examples
    com_test()

    # Uninitialize COM
    pythoncom.CoUninitialize()
