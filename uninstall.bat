@echo off
echo ==================================================
echo       LiteSwitch - Uninstaller
echo ==================================================
echo.

echo [INFO] Removing Registry Keys...
python menu_manager.py --unregister
if %errorlevel% neq 0 (
    echo [ERROR] Failed to unregister. keys might already be gone.
)

echo.
echo ==================================================
echo [SUCCESS] LiteSwitch has been removed from your context menu.
echo You can now delete this folder if you wish.
echo ==================================================
echo.
pause
