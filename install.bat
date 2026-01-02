@echo off
cd /d "%~dp0"
echo ==================================================
echo       LiteSwitch - Context Menu Installer
echo ==================================================
echo.

:: 1. Check for Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in your PATH.
    echo Please install Python 3.8+ from https://python.org
    pause
    exit /b 1
)

:: 2. Install Requirements
echo [INFO] Installing required libraries...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

:: 3. Install Pandoc (Generate temp python script)
echo.
echo [INFO] Checking/Installing Pandoc...
(
echo import pypandoc
echo import sys
echo try:
echo     pypandoc.get_pandoc_path^(^)
echo except OSError:
echo     print^("Pandoc not found. Downloading..."^)
echo     try:
echo         pypandoc.download_pandoc^(^)
echo     except Exception as e:
echo         print^(e^)
echo         sys.exit^(1^)
) > temp_pandoc_install.py

python temp_pandoc_install.py
if %errorlevel% neq 0 (
    echo [WARNING] Pandoc install failed. DOCX conversions may suffer.
)
del temp_pandoc_install.py

:: 4. Registry Cleanup
echo.
echo [INFO] Cleaning old keys...
set extensions=.docx .pdf .png .odt .txt .md .tex .pptx
for %%e in (%extensions%) do (
    reg delete "HKCU\Software\Classes\SystemFileAssociations\%%e\shell\LiteSwitch" /f >nul 2>&1
    reg delete "HKCU\Software\Classes\SystemFileAssociations\%%e\shell\LiteSwitch_pdf" /f >nul 2>&1
)

:: 5. Register New Menu
echo.
echo [INFO] Registering New Context Menu...
python menu_manager.py --register
if %errorlevel% neq 0 (
    echo [ERROR] Failed to register context menu.
    pause
    exit /b 1
)

echo.
echo ==================================================
echo [SUCCESS] LiteSwitch is ready!
echo ==================================================
echo.
pause
