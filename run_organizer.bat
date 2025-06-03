@echo off
setlocal

:: Set the directory to where the script is located
cd /d "%~dp0"

:: Check if venv exists, if not create it
if not exist "venv" (
    :: Create venv silently
    python -m venv venv >nul 2>&1
    call venv\Scripts\activate.bat >nul 2>&1
    python -m pip install --upgrade pip >nul 2>&1
    pip install -r requirements.txt >nul 2>&1
) else (
    call venv\Scripts\activate.bat >nul 2>&1
)

:: Run the window organizer using pythonw (no console window)
start /b "" "venv\Scripts\pythonw.exe" window-organiser.pyw

:: Deactivate the virtual environment silently
call venv\Scripts\deactivate.bat >nul 2>&1

endlocal 