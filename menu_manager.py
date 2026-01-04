import sys
import os
if os.name == 'nt':
    import winreg
else:
    winreg = None
    
from converter.document_converter import CONVERSION_MAP

APP_NAME = "LiteSwitch"
CLI_PATH = os.path.abspath("cli.py")
# Use pythonw.exe to avoid terminal popup
PYTHON_EXEC = sys.executable.replace("python.exe", "pythonw.exe")

# Linux Paths
LINUX_APP_DIR = os.path.expanduser("~/.local/share/applications")
LINUX_ICON_DIR = os.path.expanduser("~/.local/share/icons")
LINUX_ICON_NAME = "liteswitch.png" # Prefer PNG for Linux
LINUX_DESKTOP_FILE = "liteswitch.desktop"

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

def register_linux():
    """Registers LiteSwitch on Linux using a .desktop file."""
    print(f"Registering {APP_NAME} for Linux...")
    
    # 1. Install Icon
    if not os.path.exists(LINUX_ICON_DIR):
        os.makedirs(LINUX_ICON_DIR)
        
    icon_src = os.path.abspath(os.path.join("assets", "LiteSwitch_Logo_NEW.ico"))
    icon_dst = os.path.join(LINUX_ICON_DIR, LINUX_ICON_NAME)
    
    if os.path.exists(icon_src):
        # Convert ICO to PNG using Pillow (installed in requirements)
        try:
             from PIL import Image
             img = Image.open(icon_src)
             img.save(icon_dst, format='PNG')
             print(f"Icon installed (converted to PNG) to: {icon_dst}")
        except Exception as e:
             print(f"Failed to convert icon: {e}. Copying original...")
             import shutil
             shutil.copy(icon_src, icon_dst)
    else:
        print(f"Warning: Icon not found at {icon_src}")

    # 2. Create .desktop file
    if not os.path.exists(LINUX_APP_DIR):
        os.makedirs(LINUX_APP_DIR)
        
    desktop_path = os.path.join(LINUX_APP_DIR, LINUX_DESKTOP_FILE)
    
    # We use sys.executable to ensure we use the SAME python environment (e.g. the venv)
    # The CLI path must be absolute
    
    # MimeTypes for office docs + pdf
    mimes = "application/pdf;application/vnd.openxmlformats-officedocument.wordprocessingml.document;application/vnd.openxmlformats-officedocument.presentationml.presentation;"
    
    # Exec command: python /path/to/cli.py %f --to pdf (Default behavior?)
    # Context menu usually implies options. 
    # Since standard "Open With" just passes the file, we can't easily have sub-menus (to pdf, to png) without Actions.
    
    desktop_content = f"""[Desktop Entry]
Type=Application
Name={APP_NAME}
Comment=Convert files with LiteSwitch
Icon={icon_dst}
Exec={sys.executable} "{CLI_PATH}" %F
Terminal=false
Categories=Utility;
MimeType={mimes}
Actions=ConvertToPDF;ConvertToPNG;ConvertToTXT;

[Desktop Action ConvertToPDF]
Name=Convert to PDF
Exec={sys.executable} "{CLI_PATH}" "%f" --to pdf

[Desktop Action ConvertToPNG]
Name=Convert to PNG
Exec={sys.executable} "{CLI_PATH}" "%f" --to png

[Desktop Action ConvertToTXT]
Name=Convert to TXT
Exec={sys.executable} "{CLI_PATH}" "%f" --to txt
"""
    
    with open(desktop_path, "w") as f:
        f.write(desktop_content)
        
    print(f"Desktop entry created: {desktop_path}")
    print("LiteSwitch should now appear in 'Open With' menus!")

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

def unregister_linux():
    """Removes Linux desktop entry and icon."""
    print(f"Unregistering {APP_NAME} for Linux...")
    
    desktop_path = os.path.join(LINUX_APP_DIR, LINUX_DESKTOP_FILE)
    if os.path.exists(desktop_path):
        os.remove(desktop_path)
        print(f"Removed: {desktop_path}")
        
    icon_dst = os.path.join(LINUX_ICON_DIR, LINUX_ICON_NAME)
    if os.path.exists(icon_dst):
        os.remove(icon_dst)
        print(f"Removed: {icon_dst}")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--unregister":
        if os.name == "nt":
            unregister_menu()
        else:
            unregister_linux()
    else:
        # Default to register
        if len(sys.argv) > 1 and sys.argv[1] == "--register":
            if os.name == "nt":
                register_menu()
            else:
                register_linux()
        else:
            print("Usage: menu_manager.py [--register | --unregister]")
