import os
import zipfile
import shutil

def create_zip(zip_name, files_to_include, root_dir="."):
    print(f"Creating {zip_name}...")
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file in files_to_include:
            if os.path.isfile(file):
                zipf.write(file, arcname=file)
            elif os.path.isdir(file):
                for root, dirs, files in os.walk(file):
                    for f in files:
                        abs_path = os.path.join(root, f)
                        rel_path = os.path.relpath(abs_path, root_dir)
                        zipf.write(abs_path, arcname=rel_path)
    print(f" -> Created: {zip_name}")

def main():
    # Define common files
    COMMON_FILES = [
        "cli.py",
        "menu_manager.py",
        "requirements.txt",
        "readme.md",
        "converter", # Folder
        "assets"     # Folder
    ]
    
    # Windows specific
    WINDOWS_FILES = COMMON_FILES + ["install.bat", "uninstall.bat"]
    
    # Linux specific
    LINUX_FILES = COMMON_FILES + ["install.sh", "uninstall.sh"]
    
    # Create dist directory
    if not os.path.exists("dist"):
        os.makedirs("dist")
        
    # Build Windows Zip
    create_zip(os.path.join("dist", "LiteSwitch_Windows_v1.zip"), WINDOWS_FILES)
    
    # Build Linux Zip
    create_zip(os.path.join("dist", "LiteSwitch_Linux_v1.zip"), LINUX_FILES)
    
    print("\nBuild Complete! Check the 'dist' folder.")

if __name__ == "__main__":
    main()
