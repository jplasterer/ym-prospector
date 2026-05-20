@echo off
REM ============================================================
REM  Starkweather YM Prospector — Environment Check
REM  Called automatically by run.bat on every launch.
REM  Silent if everything is ready; installs only what's missing.
REM ============================================================

REM ── Step 1: Check for Python ─────────────────────────────────
python --version >nul 2>&1
if NOT errorlevel 1 goto :check_deps

REM Python not found — install via winget
echo.
echo ============================================================
echo  Python not found. Installing automatically...
echo  This is a one-time step and may take a few minutes.
echo ============================================================
echo.

winget install Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
if errorlevel 1 (
    echo.
    echo  ERROR: Automatic install failed.
    echo  Please install Python manually from https://python.org
    echo  Check "Add Python to PATH" during install, then run run.bat again.
    echo.
    pause
    exit /b 1
)

REM Refresh PATH so python is findable in this session
call refreshenv >nul 2>&1

REM ── Step 2: Check app dependencies ───────────────────────────
:check_deps
python -c "import streamlit" >nul 2>&1
if NOT errorlevel 1 goto :ready

REM Dependencies missing — run pip install (one-time)
echo.
echo ============================================================
echo  Installing app dependencies (one-time setup)...
echo ============================================================
echo.
pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo.
    echo  ERROR: Dependency install failed.
    echo  Please contact Joe at joe@starkweather.us
    echo.
    pause
    exit /b 1
)
echo.
echo  All dependencies installed successfully.
echo.

REM ── All good — return control to run.bat ─────────────────────
:ready
exit /b 0
