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
        regids = find_reg_clsids("pyafipws")
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
