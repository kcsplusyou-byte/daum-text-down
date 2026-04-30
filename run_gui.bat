@echo off
cd /d "%~dp0"

echo Starting Daum Cafe Re collector GUI...
echo Folder: %CD%
echo.

if exist ".venv\pyvenv.cfg" (
    for /f "tokens=1,* delims==" %%A in ('findstr /b /c:"executable = " ".venv\pyvenv.cfg" 2^>nul') do (
        for /f "tokens=* delims= " %%C in ("%%B") do (
            if not exist "%%C" (
                echo [INFO] Existing virtual environment points to missing Python. Recreating...
                rmdir /s /q ".venv"
            )
        )
    )
)

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import sys" >nul 2>nul
    if errorlevel 1 (
        echo [INFO] Existing virtual environment is broken. Recreating...
        rmdir /s /q ".venv"
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo [INFO] Creating virtual environment...
    py -3 -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create .venv. Install Python 3 first.
        echo Download: https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" -c "import tkinter, selenium, docx, webdriver_manager" >nul 2>nul
if errorlevel 1 (
    echo [INFO] Installing required packages...
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [ERROR] Package installation failed.
        echo Check internet connection, then run this file again.
        pause
        exit /b 1
    )
)

echo [INFO] Opening GUI window...
".venv\Scripts\python.exe" gui.py
set EXITCODE=%ERRORLEVEL%

if not "%EXITCODE%"=="0" (
    echo.
    echo [ERROR] GUI exited with code %EXITCODE%.
    pause
    exit /b %EXITCODE%
)
