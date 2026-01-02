#!/bin/bash

# LiteSwitch Linux Installer
# "Transparent Venv" Strategy

INSTALL_DIR="$HOME/.local/share/liteswitch"
VENV_DIR="$INSTALL_DIR/venv"

echo "=================================================="
echo "          LiteSwitch - Linux Installer"
echo "=================================================="
echo ""

# 1. Check Install Directory
if [ -d "$INSTALL_DIR" ]; then
    echo "[INFO] Removing old installation..."
    rm -rf "$INSTALL_DIR"
fi
mkdir -p "$INSTALL_DIR"

# 2. Copy Files
echo "[INFO] Copying files to $INSTALL_DIR..."
cp -r * "$INSTALL_DIR/"

# 3. Create Virtual Environment
echo "[INFO] Creating virtual environment (to avoid breaking system python)..."
cd "$INSTALL_DIR" || exit
python3 -m venv venv

if [ $? -ne 0 ] || [ ! -d "$VENV_DIR" ]; then
    echo ""
    echo "[ERROR] Failed to create virtual environment."
    echo "It seems 'python3-venv' is missing (common on Ubuntu/Debian)."
    echo "Please run the following command and try again:"
    echo ""
    echo "    sudo apt install python3-venv"
    echo ""
    exit 1
fi

# 4. Install Dependencies
echo "[INFO] Installing dependencies..."
"$VENV_DIR/bin/pip" install --upgrade pip > /dev/null
"$VENV_DIR/bin/pip" install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "[ERROR] Failed to install requirements."
    exit 1
fi

# 5. Check External Tools (LibreOffice + Zenity)
echo "[INFO] Checking for external tools..."
if ! command -v libreoffice &> /dev/null && ! command -v soffice &> /dev/null; then
    echo "[WARNING] LibreOffice not found. DOCX/PPTX conversions will fail."
    echo "          Please install it (e.g., sudo apt install libreoffice)."
fi

if ! command -v zenity &> /dev/null && ! command -v kdialog &> /dev/null; then
    echo "[WARNING] Zenity/Kdialog not found. GUI popups will fall back to console."
    echo "          Recommended: sudo apt install zenity"
fi

# 6. Register Menu
echo "[INFO] Registering Desktop Context Menu..."
"$VENV_DIR/bin/python" menu_manager.py --register
if [ $? -ne 0 ]; then
    echo "[ERROR] Failed to register menu."
    exit 1
fi

echo ""
echo "=================================================="
echo "[SUCCESS] LiteSwitch installed!"
echo "Right-click a file -> Open With -> LiteSwitch"
echo "=================================================="
echo ""
