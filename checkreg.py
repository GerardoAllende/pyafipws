import winreg
import pythoncom
import os
import ast
import sys
import ctypes


def get_com_registry_info(clsid_str, bitness=64):
    """
    Retrieves registry information for a COM component.

    Args:
        clsid_str: The CLSID of the COM component as a string (e.g., "{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}").
        bitness: 32 or 64 to specify which registry to access.

    Returns:
        A dictionary containing the registry data, or None if not found.
    """

    clsid_str = clsid_str.strip('{}')
    clsid_path = r"CLSID\{" + clsid_str + "}"
    reg_data = None

    try:
        if bitness == 32:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, clsid_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
        elif bitness == 64:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, clsid_path, 0, winreg.KEY_READ)
        else:
            raise ValueError("bitness must be 32 or 64")

        reg_data = {}
        i = 0
        while True:
            try:
                name, value, type_ = winreg.EnumValue(key, i)
                reg_data[name] = value
                i += 1
            except OSError:
                break

        subkey_count, _, _ = winreg.QueryInfoKey(key)
        for subkey_index in range(subkey_count):
            subkey_name = winreg.EnumKey(key, subkey_index)
            sub_key_path = clsid_path + "\\" + subkey_name
            if bitness == 32:
                sub_key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, sub_key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
            else:
                sub_key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, sub_key_path, 0, winreg.KEY_READ)
            sub_reg = {}
            j = 0
            while True:
                try:
                    name, value, type_ = winreg.EnumValue(sub_key, j)
                    sub_reg[name] = value
                    j += 1
                except OSError:
                    break
            reg_data[subkey_name] = sub_reg
            winreg.CloseKey(sub_key)
        winreg.CloseKey(key)

    except FileNotFoundError:
        pass  # CLSID not found
    except ValueError as e:
        print(f"Error: {e}")
        return None

    return reg_data


def get_clsid_from_progid(progid, bitness=64):
    """
    Retrieves the CLSID from a ProgID.

    Args:
        progid: The ProgID string (e.g., "Word.Application").
        bitness: 32 or 64 to specify which registry to access.

    Returns:
        The CLSID string, or None if the ProgID is not found.
    """
    try:
        key_path = f"{progid}\\CLSID"

        if bitness == 32:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
        elif bitness == 64:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_READ)
        else:
            raise ValueError("bitness must be 32 or 64")

        with key: #using with to close key automatically.
            clsid, _ = winreg.QueryValueEx(key, None)
            return clsid
    except FileNotFoundError:
        return None
    except ValueError as e:
        print(f"Error: {e}")
        return None
        

def find_reg_clsids(folder_path):
    """
    Finds all class properties named _reg_clsid_ in Python files within a folder.

    Args:
        folder_path: The path to the folder containing the Python files.

    Returns:
        A dictionary where keys are file paths and values are lists of _reg_clsid_ values.
    """
    clsids = {}
    for filename in os.listdir(folder_path):
        if filename.endswith(".py"):
            file_path = os.path.join(folder_path, filename)
            try:
                with open(file_path, "r", encoding="utf-8") as file:
                    file_content = file.read()
                    tree = ast.parse(file_content)
                    for node in ast.walk(tree):
                        if isinstance(node, ast.ClassDef):
                            progid = ''
                            for body_node in node.body:
                                if isinstance(body_node, ast.Assign) and len(body_node.targets) == 1:
                                    target = body_node.targets[0]
                                    if isinstance(target, ast.Name) and isinstance(body_node.value, ast.Str):
                                        if target.id == "_reg_progid_":
                                            progid = body_node.value.s
                                            print(progid)
                                        elif target.id == "_reg_clsid_":
                                            print(body_node.value.s)
                                            if progid:
                                                clsids[progid] = [body_node.value.s, file_path]
                                                if not "/Automate" in file_content:
                                                    print ("No /Automate in", file_path)
                                            else:
                                                print("No progid:", body_node.value.s)
            except (SyntaxError, UnicodeDecodeError) as e:
                print(f"Error processing {file_path}: {e}")

    return clsids        
        
def is_admin():
    """Checks if the current user has administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def elevate_privileges():
    """Elevates the script to administrator privileges."""
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit()
        
def delete_key(hive, key):
    try:
        with winreg.OpenKey(hive, key) as sub_key:
            while True:
                try:
                    sub_sub_key_name = winreg.EnumKey(sub_key, 0)
                    print (sub_sub_key_name)
                    delete_key(sub_key, sub_sub_key_name)
                except OSError:
                    break
    except OSError:
        print("Key does not exist:", key)
        return
    winreg.DeleteKeyEx(hive, key)

if __name__ == '__main__':
    if "--print" in sys.argv:
        from pprint import pprint
        #regids = find_reg_clsids("pyafipws")
        regids = {'COT': ['{7518B2CF-23E9-4821-BC55-D15966E15620}', 'pyafipws\\cot.py'],
                 'IIBB': ['{2C7E29D2-0C99-49D8-B04B-A16B807BB123}', 'pyafipws\\iibb.py'],
                 'PadronAFIP': ['{6206DF5E-3EEF-47E9-A532-CD81EBBAF3AA}',
                                'pyafipws\\padron.py'],
                 'PyEmail': ['{2BEF3037-BF38-41AA-84A3-6F109D543FC9}', 'pyafipws\\pyemail.py'],
                 'PyFEPDF': ['{C9B5D7BB-0388-4A5E-87D5-0B4376C7A336}', 'pyafipws\\pyfepdf.py'],
                 'PyI25': ['{5E6989E8-F658-49FB-8C39-97C74BC67650}', 'pyafipws\\pyi25.py'],
                 'PyQR': ['{B176B1CE-E7B5-4BB2-ADEC-9EB9F249DF07}', 'pyafipws\\pyqr.py'],
                 'SIRED': ['{3DC74AD5-939F-42AB-8381-FCA7AF783C77}', 'pyafipws\\sired.py'],
                 'TrazaFito': ['{39793931-450A-4F66-9324-D4D981FC5319}',
                               'pyafipws\\trazafito.py'],
                 'TrazaMed': ['{8472867A-AE6F-487F-8554-C2C896CFFC3E}',
                              'pyafipws\\trazamed.py'],
                 'TrazaProdMed': ['{D4112556-EF2E-45D3-A2A2-7A2849A364D9}',
                                  'pyafipws\\trazaprodmed.py'],
                 'TrazaRenpre': ['{461298DB-0531-47CA-B3D9-B36FE6967209}',
                                 'pyafipws\\trazarenpre.py'],
                 'TrazaVet': ['{2E28D79D-23B0-4C1E-85B9-64BDC41B8FCF}',
                              'pyafipws\\trazavet.py'],
                 'WSAA': ['{51342E57-9681-4610-AF2B-686267470930}', 'pyafipws\\wsaa.py'],
                 'WSBFE': ['{02CBC6DA-455D-4EE6-8302-411D13253CBF}', 'pyafipws\\wsbfev1.py'],
                 'WSBFEv1': ['{EE4ABEE2-76DD-450F-880B-66710AE464D6}', 'pyafipws\\wsbfev1.py'],
                 'WSCDC': ['{D1B97BDD-A78C-4D51-8999-1D9A5034EC10}', 'pyafipws\\wscdc.py'],
                 'WSCOC': ['{B30406CE-326A-46D9-B807-B7916E3F1B96}', 'pyafipws\\wscoc.py'],
                 'WSCPE': ['{37F6A7B5-344E-45C5-9198-0CF7B206F409}', 'pyafipws\\wscpe.py'],
                 'WSCT': ['{5DE7917D-CE97-4C88-B6C7-DAF8CEB54E93}', 'pyafipws\\wsct.py'],
                 'WSCTG': ['{4383E947-57C4-47C5-8419-85221580CB48}', 'pyafipws\\wsctg.py'],
                 'WSCTGv2': ['{ACDEFB8A-34E1-48CF-94E8-6AF6ADA0717A}', 'pyafipws\\wsctg.py'],
                 'WSFECred': ['{F4B2B652-C992-4E46-9134-121F62011C46}',
                              'pyafipws\\wsfecred.py'],
                 'WSFEX': ['{B3C8D3D3-D5DA-44C9-B003-11845803B2BD}', 'pyafipws\\wsfexv1.py'],
                 'WSFEXv1': ['{8106F039-D132-4F87-8AFE-ADE47B5503D4}', 'pyafipws\\wsfexv1.py'],
                 'WSFEv1': ['{FA1BB90B-53D1-4FDA-8D1F-DEED2700E739}', 'pyafipws\\wsfev1.py'],
                 'WSLPG': ['{9D21C513-21A6-413C-8592-047357692608}', 'pyafipws\\wslpg.py'],
                 'WSLSP': ['{9750BBD4-FBC3-4FE7-8DE5-E193667D6813}', 'pyafipws\\wslsp.py'],
                 'WSLTV': ['{C6EEAE8A-7560-4538-B29C-76434A8C2DC3}', 'pyafipws\\wsltv.py'],
                 'WSLUM': ['{4CBB2DF8-7AAE-434E-916D-9D663BB1CAFC}', 'pyafipws\\wslum.py'],
                 'WSMTXCA': ['{8128E6AB-FB22-4952-8EA6-BD41C29B17CA}', 'pyafipws\\wsmtx.py'],
                 'WSRemAzucar': ['{448F912A-C013-4E19-8D52-7FC88305590A}',
                                 'pyafipws\\wsremazucar.py'],
                 'WSRemCarne': ['{71DB0CB9-2ED7-4226-A1E6-C3FA7FB18F41}',
                                'pyafipws\\wsremcarne.py'],
                 'WSRemHarina': ['{72BFB9B9-0FD9-497C-8C62-5D41F7029377}',
                                 'pyafipws\\wsremharina.py'],
                 'WSSIREc2005': ['{E941985A-5C1C-4B17-80C1-7FBB7BE7D713}',
                                 'pyafipws\\ws_sire.py'],
                 'WSSrPadronA4': ['{C2270008-4324-46F6-A2D3-60836EE63BD7}',
                                  'pyafipws\\ws_sr_padron.py'],
                 'WSSrPadronA5': ['{DF7447DD-EEF3-4E6B-A93B-F969B5075EC8}',
                                  'pyafipws\\ws_sr_padron.py']}
        print("\nRegistro de 32 bits")
        for key, value in regids.items():
            a = get_com_registry_info(value[0], 32)
            if a:
                pprint(a)
        print("\nRegistro de 64 bits")
        for key, value in regids.items():
            a = get_com_registry_info(value[0], 64)
            if a:
                pprint(a)
        input("Presione intro para salir")
    elif "--clean" in sys.argv:
        if not is_admin():
            elevate_privileges() #relaunch with admin privs, then exit this process.
        regids = find_reg_clsids("pyafipws")
        for key, value in regids.items():
            clsid = value[0]
            progid = key
            delete_key(winreg.HKEY_CURRENT_USER, rf"SOFTWARE\Classes\CLSID\{clsid}")
            delete_key(winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Classes\CLSID\{clsid}")
            
            delete_key(winreg.HKEY_CURRENT_USER, rf"SOFTWARE\Classes\AppID\{clsid}")
            delete_key(winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Classes\AppID\{clsid}")
            
            delete_key(winreg.HKEY_CURRENT_USER, rf"SOFTWARE\Classes\{progid}")
            delete_key(winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Classes\{progid}")
            
            delete_key(winreg.HKEY_CURRENT_USER, rf"SOFTWARE\Classes\Wow6432Node\CLSID\{clsid}")
            delete_key(winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Classes\Wow6432Node\CLSID\{clsid}")
            
            delete_key(winreg.HKEY_CURRENT_USER, rf"SOFTWARE\Classes\Wow6432Node\AppID\{clsid}")
            delete_key(winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Classes\Wow6432Node\AppID\{clsid}")
            
            delete_key(winreg.HKEY_CURRENT_USER, rf"SOFTWARE\Classes\Wow6432Node\{progid}")
            delete_key(winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Classes\Wow6432Node\{progid}")
            
            delete_key(winreg.HKEY_CLASSES_ROOT, rf"Wow6432Node\AppID\{clsid}")
            
            delete_key(winreg.HKEY_CLASSES_ROOT, rf"{progid}")
            
            
            
        input("Presione intro para salir")
