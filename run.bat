@echo off
cd /d "%~dp0"
echo Starting Starkweather YM Prospector...
echo.

REM Try to open in Chrome after a short delay.
REM If Chrome is not installed, fall back to the default browser.
start "" /b cmd /c "timeout /t 3 /nobreak >nul && (where chrome >nul 2>&1 && start chrome http://localhost:8501 || start http://localhost:8501)"

python -m streamlit run app.py --browser.serverAddress localhost --server.headless true
pause
