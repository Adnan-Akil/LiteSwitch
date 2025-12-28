import sys
import os
import winreg
from converter.document_converter import CONVERSION_MAP

APP_NAME = "LiteSwitch"
CLI_PATH = os.path.abspath("cli.py")
# Use pythonw.exe to avoid terminal popup
PYTHON_EXEC = sys.executable.replace("python.exe", "pythonw.exe")

def cleanup_old_keys():
    """Aggressively remove old/broken keys from previous versions."""
    print("Cleaning up old Registry keys...")
    for source_ext, targets in CONVERSION_MAP.items():
        ext_key = f".{source_ext}"
        
        # 1. Remove the old flat keys (e.g. LiteSwitch_PDF)
        for target_ext in targets:
            old_key_path = f"Software\\Classes\\SystemFileAssociations\\{ext_key}\\shell\\{APP_NAME}_{target_ext}"
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, old_key_path + "\\command")
            except: pass
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, old_key_path)
            except: pass

def register_menu():
    print(f"Registering {APP_NAME}...")
    
    # Run cleanup first
    cleanup_old_keys()
    
    # We register under SystemFileAssociations for each supported extension
    success_count = 0
    
    for source_ext, targets in CONVERSION_MAP.items():
        ext_key = f".{source_ext}"
        
        # Path to icon (Prefer ICO for valid resource scaling)
        # Look in assets folder
        icon_path = os.path.abspath(os.path.join("assets", "LiteSwitch_Logo_NEW.ico"))
        
        if not os.path.exists(icon_path):
             icon_path = PYTHON_EXEC # Fallback
        
        # 1. Create the Main Parent Menu Item ("Convert with LiteSwitch")
        # Key: HKCU\Software\Classes\SystemFileAssociations\.ext\shell\LiteSwitch
        parent_key_path = f"Software\\Classes\\SystemFileAssociations\\{ext_key}\\shell\\{APP_NAME}"
        
        try:
            # Create Parent Key
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, parent_key_path)
            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Convert with LiteSwitch")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, icon_path)
            winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ, "") # Enables cascading
            winreg.CloseKey(key)
            
            # 2. Add Sub-commands under the 'shell' subkey of the Parent
            # Key: ...\LiteSwitch\shell
            shell_key_path = f"{parent_key_path}\\shell"
            winreg.CreateKey(winreg.HKEY_CURRENT_USER, shell_key_path)
            
            for target_ext, _ in targets.items():
                # Sub-item Key: ...\LiteSwitch\shell\to_format
                sub_key_path = f"{shell_key_path}\\{APP_NAME}_to_{target_ext}"
                menu_text = f"to {target_ext.upper()}"
                command = f'"{PYTHON_EXEC}" "{CLI_PATH}" "%1" --to {target_ext}'
                
                # Create Sub-item
                sub_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, sub_key_path)
                winreg.SetValueEx(sub_key, "", 0, winreg.REG_SZ, menu_text)
                winreg.SetValueEx(sub_key, "Icon", 0, winreg.REG_SZ, icon_path) # Add Icon to submenu too
                
                # Command
                cmd_key = winreg.CreateKey(sub_key, "command")
                winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, command)
                
                winreg.CloseKey(cmd_key)
                winreg.CloseKey(sub_key)
                
            success_count += 1
        except Exception as e:
            print(f"Failed to register for {source_ext}: {e}")

    print(f"Successfully registered context menus for {success_count} file types.")

def unregister_menu():
    print(f"Unregistering {APP_NAME}...")
    
    count = 0 
    for source_ext in CONVERSION_MAP.keys():
        ext_key = f".{source_ext}"
        parent_key_path = f"Software\\Classes\\SystemFileAssociations\\{ext_key}\\shell\\{APP_NAME}"
        
        try:
             # Deleting registry keys recursively is tricky in pure Python standard lib 
             # without a helper. We'll try to delete the known structure.
             
             # 1. Delete sub-commands first
             shell_key_path = f"{parent_key_path}\\shell"
             try:
                 shell_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, shell_key_path)
                 while True:
                     try:
                         sub = winreg.EnumKey(shell_key, 0)
                         winreg.DeleteKey(winreg.HKEY_CURRENT_USER, f"{shell_key_path}\\{sub}\\command")
                         winreg.DeleteKey(winreg.HKEY_CURRENT_USER, f"{shell_key_path}\\{sub}")
                     except OSError:
                         break # No more keys
                 winreg.CloseKey(shell_key)
                 winreg.DeleteKey(winreg.HKEY_CURRENT_USER, shell_key_path)
             except FileNotFoundError:
                 pass
             
             # 2. Delete Parent
             winreg.DeleteKey(winreg.HKEY_CURRENT_USER, parent_key_path)
             count += 1
        except Exception as e:
             # It might fail if keys don't exist or permissions, mostly fine
             # print(f"Error removing {parent_key_path}: {e}")
             pass
                 
    print(f"Removed keys for {count} extensions.")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--unregister":
        unregister_menu()
    else:
        # Default to register
        if len(sys.argv) > 1 and sys.argv[1] == "--register":
            register_menu()
        else:
            print("Usage: menu_manager.py [--register | --unregister]")
