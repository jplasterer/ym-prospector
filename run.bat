@echo off
cd /d "%~dp0"
echo Starting Starkweather YM Prospector...
echo.

REM ── Ensure Python and dependencies are ready ─────────────────
call "%~dp0setup.bat"
if errorlevel 1 (
    echo Setup failed. Please resolve the error above and try again.
    pause
    exit /b 1
)

REM ── Open in Chrome if available, otherwise default browser ────
start "" /b cmd /c "timeout /t 3 /nobreak >nul && (where chrome >nul 2>&1 && start chrome http://localhost:8501 || start http://localhost:8501)"

REM ── Launch the app ────────────────────────────────────────────
python -m streamlit run app.py --browser.serverAddress localhost --server.headless true
pause
