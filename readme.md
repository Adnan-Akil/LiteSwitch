<div align="center">
  <img src="assets/LiteSwitch_Logo_NEW.ico" alt="LiteSwitch Logo" width="128" />
  <h1>LiteSwitch</h1>
  <p><strong>"From this to that—just like that."</strong></p>
  
  [![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)](https://github.com/Adnan-Akil/LiteSwitch/releases)
  [![Linux](https://img.shields.io/badge/Platform-Linux-FCC624?logo=linux&logoColor=black)](https://github.com/Adnan-Akil/LiteSwitch/releases)
  [![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
</div>

---

**LiteSwitch** is a premium, lightweight context-menu utility for **Windows** and **Linux**. It integrates directly into your operating system, allowing you to convert documents and images instantly with a simple Right-Click. 

No web uploads, no subscriptions, no heavy interfaces. Just seamless workflow.

## 🚀 Features

*   **Native Integration**: Appears naturally in your Context Menu / "Open With" list.
*   **Privacy First**: All conversions happen locally on your machine.
*   **Smart Formats**:
    *   **DOCX** → PDF, TXT, Markdown, ODT
    *   **PDF** → PPTX (Perfect visual layout), DOCX, Text, Images
    *   **PPTX** → PDF, PNG Slides, DOCX Handouts
    *   **Images** → PDF
*   **Cross-Platform**: Now fully supported on Linux with a native GTK/Qt feel.

## 📦 Installation

### <img src="https://img.icons8.com/color/48/000000/windows-10.png" width="20"/> Windows

1.  **Download** the latest `LiteSwitch_Windows_v1.zip`.
2.  Extract and double-click **`install.bat`**.
3.  That's it! Right-click any file to see the **LiteSwitch** menu.

### <img src="https://img.icons8.com/color/48/000000/linux--v1.png" width="20"/> Linux

1.  **Download** the latest `LiteSwitch_Linux_v1.zip`.
2.  Extract the folder.
3.  Open a terminal in the folder and run:
    ```bash
    ./install.sh
    ```
4.  **Usage**:
    *   Right-click a file → **Open With** → **LiteSwitch**.
    *   Select your target format from the popup menu!

> **Note**: Linux users may need **LibreOffice** installed for DOCX/PPTX conversions.
> ```bash
> sudo apt install libreoffice python3-venv  # Ubuntu/Debian
> ```

## 🛠️ Requirements

*   **Windows**: Windows 10/11, Microsoft Word/PowerPoint (for high-fidelity conversion).
*   **Linux**: Python 3.8+, LibreOffice (recommended), Zenity or Kdialog (for GUI prompts).

## 🗑️ Uninstallation

*   **Windows**: Run `uninstall.bat`.
*   **Linux**: Run `./uninstall.sh`.
    *   This will remove the context menu entry and clean up the `~/.local/share/liteswitch` directory.

---
<div align="center">
  <sub>Built with ❤️ for speed and simplicity.</sub>
  <br>
  <sub>Copyright © 2026 Adnan Akil</sub>
</div>
