import os
import sys
import winreg

APP_NAME = "ChemicalManager"

def get_executable_path():
    """
    Returns the path to the executable to run on startup.
    If frozen (PyInstaller), returns the path to the .exe.
    Otherwise, returns pythonw.exe main.py.
    """
    if getattr(sys, 'frozen', False):
        return sys.executable
    else:
        # Use pythonw so it doesn't open a console window if run from source
        python_exe = sys.executable.replace('python.exe', 'pythonw.exe')
        main_script = os.path.abspath(sys.argv[0])
        return f'"{python_exe}" "{main_script}"'

def set_run_on_startup(enable: bool):
    """
    Adds or removes the application from the Windows registry Run key.
    When enabling, adds the '-a' flag to run minimized and auto-sync.
    """
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
        
        if enable:
            exe_path = get_executable_path()
            # If the path has spaces but isn't wrapped in quotes (e.g., PyInstaller exe), wrap it.
            if getattr(sys, 'frozen', False) and not exe_path.startswith('"'):
                exe_path = f'"{exe_path}"'
                
            command = f'{exe_path} -a'
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, command)
        else:
            try:
                winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError:
                pass # Already removed
                
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Failed to set startup registry key: {e}")
        return False

def check_run_on_startup() -> bool:
    """
    Checks if the application is currently in the registry Run key.
    """
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
        try:
            val, _ = winreg.QueryValueEx(key, APP_NAME)
            return True
        except FileNotFoundError:
            return False
        finally:
            winreg.CloseKey(key)
    except Exception:
        return False
