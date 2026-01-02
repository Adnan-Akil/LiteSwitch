import sys
import os
import argparse
import logging
import tempfile
import platform
import subprocess
import shutil
import ctypes

from converter.document_converter import CONVERSION_MAP

# Setup Logger
LOG_FILE = os.path.join(tempfile.gettempdir(), "liteswitch.log")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def show_linux_message(title, msg, is_error=False):
    """Show a GUI message on Linux using Zenity or Kdialog."""
    # Try Zenity (GNOME/GTK)
    if shutil.which("zenity"):
        kind = "--error" if is_error else "--info"
        subprocess.run(["zenity", kind, "--title", title, "--text", msg, "--no-wrap"], stderr=subprocess.DEVNULL)
        return

    # Try Kdialog (KDE)
    if shutil.which("kdialog"):
        kind = "--error" if is_error else "--msgbox"
        subprocess.run(["kdialog", "--title", title, kind, msg], stderr=subprocess.DEVNULL)
        return
        
    # Fallback to console
    print(f"\n[{title}] {msg}\n")

def show_message(title, msg, is_error=False):
    """Show a GUI message box using ctypes (Windows) or Zenity/Kdialog (Linux)."""
    if platform.system() == "Windows":
        # 0x40 = MB_ICONINFORMATION, 0x10 = MB_ICONERROR, 0x0 = MB_OK, 0x40000 = MB_TOPMOST
        flags = 0x0 | 0x40000
        if is_error:
            flags |= 0x10
        else:
            flags |= 0x40
            
        try:
            ctypes.windll.user32.MessageBoxW(0, msg, title, flags)
        except Exception as e:
            logging.error(f"Failed to show popup: {e}")
            print(msg) # fallback
    else:
        # Linux/macOS
        show_linux_message(title, msg, is_error)

def main():
    parser = argparse.ArgumentParser(description="LiteSwitch File Converter")
    parser.add_argument("input_file", help="Path to the input file")
    parser.add_argument("--to", required=False, help="Target format extension (e.g. pdf, docx)")
    
    args = parser.parse_args()
    
    input_path = os.path.abspath(args.input_file.strip('"').strip("'"))
    input_path = os.path.abspath(args.input_file.strip('"').strip("'"))
    
    if args.to:
        target_ext = args.to.lower().lstrip('.')
    else:
        target_ext = None
    
    if not os.path.exists(input_path):
        err = f"File not found: {input_path}"
        logging.error(err)
        show_message("LiteSwitch Error", f"{err}\n\nSee log: {LOG_FILE}", is_error=True)
        sys.exit(1)

    source_ext = os.path.splitext(input_path)[1].lower().lstrip('.')
    
    logging.info(f"CLI Request: {input_path} -> {target_ext}")

    if source_ext not in CONVERSION_MAP:
        err = f"No conversions supported for input type: .{source_ext}"
        logging.error(err)
        show_message("LiteSwitch Error", f"{err}\n\nSee log: {LOG_FILE}", is_error=True)
        sys.exit(1)
        
    # Smart Selection (GUI) if target not provided
    if not target_ext:
        if source_ext not in CONVERSION_MAP:
            err = f"No conversions supported for input type: .{source_ext}"
            show_message("LiteSwitch Error", err, is_error=True)
            sys.exit(1)
            
        supported = list(CONVERSION_MAP[source_ext].keys())
        
        # If only one option, default to it? No, explicit is better.
        # Launch Zenity List
        if shutil.which("zenity"):
            # Construct list for zenity
            # zenity --list --column="Format" "pdf" "png" ...
            cmd = ["zenity", "--list", "--title", "Convert to...", "--column", "Format"] + supported
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                target_ext = result.stdout.strip()
            else:
                sys.exit(0) # User cancelled
        elif shutil.which("kdialog"):
             # kdialog --menu "Convert to..." "pdf" "pdf"...
             # kdialog menu requires pairs 'tag' 'item'.
             menu_args = []
             for s in supported:
                 menu_args.extend([s, s])
             cmd = ["kdialog", "--menu", "Choose Target Format",] + menu_args
             result = subprocess.run(cmd, capture_output=True, text=True)
             if result.returncode == 0 and result.stdout.strip():
                  target_ext = result.stdout.strip()
             else:
                  sys.exit(0)
        else:
            # Console fallback
            print("\nAvailable formats:")
            for i, fmt in enumerate(supported, 1):
                print(f"{i}. {fmt}")
            try:
                choice = input("\nSelect format (number): ")
                target_ext = supported[int(choice)-1]
            except:
                sys.exit(1)

    if target_ext not in CONVERSION_MAP[source_ext]:
        err = f"Cannot convert .{source_ext} to .{target_ext}"
        logging.error(err)
        show_message("LiteSwitch Error", f"{err}\n\nSee log: {LOG_FILE}", is_error=True)
        sys.exit(1)

    try:
        converter = CONVERSION_MAP[source_ext][target_ext]
        output_path = converter(input_path)
        logging.info(f"Success: {output_path}")
        
        # Success Popup 
        filename = os.path.basename(output_path)
        
        # Try to read catchphrase from assets
        catchphrase = "From this to that--just like that."
        cp_path = os.path.join(os.path.dirname(__file__), "assets", "catchphrase.txt")
        if os.path.exists(cp_path):
            try:
                with open(cp_path, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                    if content:
                        catchphrase = content.strip('"') # Remove quotes if present
            except:
                pass

        show_message("LiteSwitch Success", f"{catchphrase}\n\nLiteSwitch has done the job!\nCreated: {filename}")
            
    except Exception as e:
        logging.exception("Conversion failed")
        show_message("LiteSwitch Error", f"An error occurred during conversion.\n\n{str(e)}\n\nSee log: {LOG_FILE}", is_error=True)
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except Exception:
        logging.exception("Critical CLI failure")
        # Try to show error even in critical fail
        try:
             if platform.system() == "Windows":
                ctypes.windll.user32.MessageBoxW(0, "Critical failure in LiteSwitch check logs.", "LiteSwitch Fatal", 0x10)
             else:
                print("Critical failure in LiteSwitch check logs.")
        except: pass