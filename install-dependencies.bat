@echo off
setlocal

REM Check for Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python not found. Downloading and installing Python...
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.13.5/python-3.13.5-amd64.exe' -OutFile 'python-installer.exe'"
    start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
    del python-installer.exe
) else (
    echo Python is already installed.
)

REM Check for Ghostscript
where gswin64c >nul 2>nul
if %errorlevel% neq 0 (
    echo Ghostscript not found. Downloading and installing Ghostscript...
    powershell -Command "Invoke-WebRequest -Uri 'https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs10051/gs10051w64.exe' -OutFile 'gs-installer.exe'"
    start /wait gs-installer.exe /S
    del gs-installer.exe
) else (
    echo Ghostscript is already installed.
)

REM Check for Poppler
set POPPLER_DIR="C:\Program Files\poppler-24.08.0"
if not exist %POPPLER_DIR% (
    echo Poppler not found. Downloading and extracting Poppler...
    powershell -Command "Invoke-WebRequest -Uri 'https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip' -OutFile 'poppler.zip'"
    powershell -Command "Expand-Archive -Path 'poppler.zip' -DestinationPath 'C:\Program Files'"
    del poppler.zip
) else (
    echo Poppler is already installed.
)

REM Create Python venv and install requirements
set SCRIPT_DIR=%~dp0
cd /d %SCRIPT_DIR%
if not exist .venv (
    python -m venv .venv
)
call .venv\Scripts\activate.bat
pip install --upgrade pip
pip install -r requirements.txt

echo Installation complete.
endlocal
