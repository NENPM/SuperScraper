@echo off
set venv_dir=venv

REM Check if Python is installed
where python >nul 2>nul
if errorlevel 1 (
    echo Python not found. Please install Python and try again.
    pause
    exit /b
)

REM Create virtual environment if it doesn't exist
if not exist %venv_dir% (
    echo Creating virtual environment...
    python -m venv %venv_dir%
)

REM Activate the virtual environment
call %venv_dir%\Scripts\activate.bat

REM Install required packages
echo Installing dependencies...
pip install --upgrade pip >nul
pip install -r requirements.txt > pip_install.log 2>&1
if errorlevel 1 (
    echo Failed to install dependencies. Check pip_install.log for details.
    pause
    exit /b
)

REM Run the main script
python SuperScraper.py

pause