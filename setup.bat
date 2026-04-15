@echo off
echo ============================================
echo  Starkweather YM Prospector — First-Time Setup
echo ============================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed.
    echo Please download and install Python from https://python.org
    echo Choose the LTS version and check "Add to PATH" during install.
    pause
    exit /b 1
)

echo Python found. Installing dependencies...
echo.
pip install -r "%~dp0requirements.txt"

echo.
echo ============================================
echo  Setup complete!
echo  Run the app by double-clicking run.bat
echo ============================================
echo.
pause
