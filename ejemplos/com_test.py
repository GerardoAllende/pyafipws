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
    cot = win32com.client.Dispatch("COT")
    print("\nCOT", cot.Version)
    pyqr = win32com.client.Dispatch("PyQR")
    print("\nPyQR", pyqr.Version)
    cred = win32com.client.Dispatch("WSFECred")
    print("\nWSFECred", cred.Version)
    print("Conectar: ", cred.Conectar())
    print_dummy(cred)

if __name__ == "__main__":
    # Initialize COM
    pythoncom.CoInitialize()

    # Run the examples
    com_test()

    # Uninitialize COM
    pythoncom.CoUninitialize()
