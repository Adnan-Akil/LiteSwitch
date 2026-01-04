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
    parser.add_argument("input_files", nargs='+', help="Path to the input file(s)")
    parser.add_argument("--to", required=False, help="Target format extension (e.g. pdf, docx)")
    
    args = parser.parse_args()
    
    # Robust Argument Parsing
    # Sometimes (e.g. Linux Desktop Entry with "%F"), args might be passed as a single merged string
    # e.g. ["'/path/A' '/path/B'"] instead of ["/path/A", "/path/B"]
    
    raw_files = args.input_files
    if len(raw_files) == 1 and ("' '" in raw_files[0] or '" "' in raw_files[0]):
        import shlex
        try:
            # shlex.split will handle the quotes correctly
            raw_files = shlex.split(raw_files[0])
            logging.info(f"Detected merged arguments, split into: {len(raw_files)} files")
        except Exception as e:
            logging.warning(f"Failed to split merged args: {e}")

    # args.input_files is now a proper list
    input_paths = [os.path.abspath(f.strip('"').strip("'")) for f in raw_files]
    
    if args.to:
        target_ext = args.to.lower().lstrip('.')
    else:
        target_ext = None
    
    # Validation: Check existence
    valid_files = []
    for p in input_paths:
        if os.path.exists(p):
            valid_files.append(p)
        else:
            logging.error(f"File not found: {p}")

    if not valid_files:
        sys.exit(1)

    # Detect common source extension for batch prompt
    # We take the extension of the FIRST file as the driver for simplicity
    first_ext = os.path.splitext(valid_files[0])[1].lower().lstrip('.')
    
    # Verify all have same extension? If not, we might be in trouble for a single prompt.
    # For now, let's assume batch selection usually involves same types.
    # If mixed, we only support converting the ones matching the first one, or we error?
    # Better UX: Warn if mixed? Or just try to process what we can.
    # PROMPT LOGIC: Based on first_ext.
    
    if not target_ext:
        if first_ext not in CONVERSION_MAP:
             show_message("LiteSwitch Error", f"No conversions supported for input type: .{first_ext}", is_error=True)
             sys.exit(1)

        supported = list(CONVERSION_MAP[first_ext].keys())
        
        # Launch Zenity List
        if shutil.which("zenity"):
            title = f"Convert {len(valid_files)} file(s) to..."
            cmd = ["zenity", "--list", "--title", title, "--column", "Format"] + supported
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                target_ext = result.stdout.strip()
            else:
                sys.exit(0) # User cancelled
        elif shutil.which("kdialog"):
             menu_args = []
             for s in supported:
                 menu_args.extend([s, s])
             cmd = ["kdialog", "--menu", f"Convert {len(valid_files)} files",] + menu_args
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

    # Now execute conversion loop
    success_count = 0
    errors = []
    
    for input_path in valid_files:
        current_ext = os.path.splitext(input_path)[1].lower().lstrip('.')
        
        # Check if this specific file supports the detected target
        # (Handling mixed batches gracefully)
        if current_ext not in CONVERSION_MAP or target_ext not in CONVERSION_MAP[current_ext]:
            logging.warning(f"Skipping {os.path.basename(input_path)}: cannot convert .{current_ext} to .{target_ext}")
            continue

        try:
            converter = CONVERSION_MAP[current_ext][target_ext]
            output_path = converter(input_path)
            logging.info(f"Success: {output_path}")
            success_count += 1
        except Exception as e:
            logging.exception(f"Failed to convert {input_path}")
            errors.append(os.path.basename(input_path))

    # Final Summary
    if success_count > 0:
        msg = f"Successfully converted {success_count} file(s) to {target_ext.upper()}!"
        if errors:
            msg += f"\n\nFailed ({len(errors)}): {', '.join(errors)}"
            
        # Catchphrase only on success
        catchphrase = "From this to that--just like that."
        cp_path = os.path.join(os.path.dirname(__file__), "assets", "catchphrase.txt")
        if os.path.exists(cp_path):
            try:
                with open(cp_path, "r", encoding="utf-8") as f:
                    c = f.read().strip()
                    if c: catchphrase = c.strip('"')
            except: pass
            
        show_message("LiteSwitch Success", f"{catchphrase}\n\n{msg}")
    elif errors:
        show_message("LiteSwitch Error", f"Conversion failed for all files.\nCheck logs for details.", is_error=True)

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