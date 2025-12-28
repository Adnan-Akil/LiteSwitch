# LiteSwitch

> **"From this to that—just like that."**

**LiteSwitch** is a lightweight, powerful context-menu tool for Windows that lets you convert files instantly with a Right-Click. No heavy GUIs, no web uploads—just click and switch.

## Quick Start
1. **Download** this folder.
2. Double-click **`install.bat`**.
3. **Right-click** any file to convert it!

---

## Features
- **Smart Context Menu**: Shows only relevant conversions (e.g., DOCX -> PDF, PDF -> PNG).
- **Formats**: Supports DOCX, PDF, PNG, ODT, TXT, Markdown, HTML.
- **Tesseract Support**: 
  - If installed, PDF->PPTX conversion is editable.
  - If missing, it safely falls back to "Image Mode" (LiteSwitch will inform you).

## Requirements
- **Python 3.8+**
- **Microsoft Word**: For best-quality DOCX -> PDF conversion.
- **(Optional) Tesseract OCR**: For editable PDF -> PPTX slides.

## Advanced
- **Manual Install**: `pip install -r requirements.txt` then `python dynamic_context_menu.py --register`
- **Uninstall**: Run `python dynamic_context_menu.py --unregister`
- **Logs**: Errors are logged to `%TEMP%\liteswitch.log`.
