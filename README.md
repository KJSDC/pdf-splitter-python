
PDF Splitter & Compressor
========================

This tool lets you preview, split, and compress PDF files with a simple GUI.

## Quick Start

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the app:**
   ```bash
   python pdf_splitter_gui.py
   ```

## Requirements

- PyPDF2==3.0.1
- pdf2image==1.16.3
- Pillow==10.3.0


## Windows Setup

Windows users can run the provided batch script to automatically install Python, Poppler, Ghostscript, and all required Python packages:

```bat
install-dependencies.bat
```

This will:
- Download and install Python (if missing)
- Download and install Ghostscript (if missing)
- Download and extract Poppler (if missing)
- Add Poppler to the PATH for the session
- Create a Python virtual environment and install all required packages

---

If you prefer manual setup:

### Poppler (for pdf2image)
- Download precompiled binaries: https://github.com/oschwartz10612/poppler-windows/releases
- Extract and add the `bin` folder (e.g., `C:\Program Files\poppler-24.08.0\Library\bin`) to your system PATH, or specify the path in the script.

### Ghostscript (for compression)
- Download: https://www.ghostscript.com/download/gsdnld.html
- Add the installation path (e.g., `C:\Program Files\gs\gs10.03.0\bin`) to your system PATH, or place `gswin64c.exe` in the project folder and update the script.

> **Note:** The script calls `gswin64c`. Ensure this executable is available in your PATH or bundled with your app.

## Bundling for Distribution

If you want a portable Windows executable:
- Bundle Poppler and Ghostscript folders with your app.
- Use PyInstaller:
  ```bash
  pyinstaller --onefile --windowed --add-data "ghostscript;ghostscript" --add-data "poppler;poppler" pdf_splitter_gui.py
  ```
- Update the script to use bundled paths for Poppler and Ghostscript.

## License

MIT