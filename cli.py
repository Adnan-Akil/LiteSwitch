import sys
import os
import argparse
import logging
import tempfile
import traceback
import ctypes

from converter.document_converter import CONVERSION_MAP

# Setup Logger
LOG_FILE = os.path.join(tempfile.gettempdir(), "liteswitch.log")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def show_message(title, msg, is_error=False):
    """Show a GUI message box using ctypes (no extra deps)."""
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

def main():
    parser = argparse.ArgumentParser(description="LiteSwitch File Converter")
    parser.add_argument("input_file", help="Path to the input file")
    parser.add_argument("--to", required=True, help="Target format extension (e.g. pdf, docx)")
    
    args = parser.parse_args()
    
    input_path = os.path.abspath(args.input_file.strip('"'))
    target_ext = args.to.lower().lstrip('.')
    
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
        catchphrase = "From this to that--just like that."
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
             ctypes.windll.user32.MessageBoxW(0, "Critical failure in LiteSwitch check logs.", "LiteSwitch Fatal", 0x10)
        except: pass