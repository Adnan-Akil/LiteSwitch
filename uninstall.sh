#!/bin/bash

# LiteSwitch Linux Uninstaller

INSTALL_DIR="$HOME/.local/share/liteswitch"

echo "=================================================="
echo "          LiteSwitch - Linux Uninstaller"
echo "=================================================="
echo ""

# 1. Runs the python unregistration (removes desktop file & icon)
if [ -d "$INSTALL_DIR/venv" ]; then
    echo "[INFO] Unregistering Desktop Menu..."
    "$INSTALL_DIR/venv/bin/python" menu_manager.py --unregister
else
    # Fallback if venv is gone but we want to clean up manual paths?
    # We can try to manually remove the known paths just in case
    echo "[INFO] Manual cleanup of desktop entries..."
    rm -f "$HOME/.local/share/applications/liteswitch.desktop"
    rm -f "$HOME/.local/share/icons/liteswitch.png"
fi

# 2. Remove the install directory
if [ -d "$INSTALL_DIR" ]; then
    echo "[INFO] Removing installation files ($INSTALL_DIR)..."
    rm -rf "$INSTALL_DIR"
else
    echo "[INFO] Installation directory not found."
fi

echo ""
echo "=================================================="
echo "[SUCCESS] LiteSwitch has been uninstalled."
echo "=================================================="
echo ""
