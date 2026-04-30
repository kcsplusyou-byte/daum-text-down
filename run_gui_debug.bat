@echo off
cd /d "%~dp0"

echo Starting GUI debug mode...
echo Folder: %CD%
echo.

if exist ".venv\pyvenv.cfg" (
    for /f "tokens=1,* delims==" %%A in ('findstr /b /c:"executable = " ".venv\pyvenv.cfg" 2^>nul') do (
        for /f "tokens=* delims= " %%C in ("%%B") do (
            if not exist "%%C" (
                echo [INFO] Existing virtual environment points to missing Python. Recreating... > gui_error.log
                rmdir /s /q ".venv" >> gui_error.log 2>&1
            )
        )
    )
)

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import sys" >nul 2>nul
    if errorlevel 1 (
        echo [INFO] Existing virtual environment is broken. Recreating... > gui_error.log
        rmdir /s /q ".venv" >> gui_error.log 2>&1
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo [INFO] Creating virtual environment... > gui_error.log
    py -3 -m venv .venv >> gui_error.log 2>&1
    if errorlevel 1 (
        echo [ERROR] Failed to create .venv. Install Python 3 first. >> gui_error.log
        echo Download: https://www.python.org/downloads/ >> gui_error.log
        type gui_error.log
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" -c "import tkinter, selenium, docx, webdriver_manager" >nul 2>nul
if errorlevel 1 (
    echo [INFO] Installing required packages... > gui_error.log
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt >> gui_error.log 2>&1
    if errorlevel 1 (
        echo.
        echo [ERROR] Package installation failed. See gui_error.log.
        type gui_error.log
        pause
        exit /b 1
    )
)

echo [INFO] Opening GUI window... > gui_error.log
".venv\Scripts\python.exe" gui.py >> gui_error.log 2>&1
set EXITCODE=%ERRORLEVEL%

if not "%EXITCODE%"=="0" (
    echo.
    echo [ERROR] GUI failed. See gui_error.log.
    type gui_error.log
    pause
    exit /b %EXITCODE%
)
