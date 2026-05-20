@echo off
cd /d "%~dp0"
echo Starting Starkweather YM Prospector...
echo.

REM Open in Chrome after a short delay (Chrome required for LinkedIn enrichment)
start "" /b cmd /c "timeout /t 3 /nobreak >nul && start chrome http://localhost:8501"

python -m streamlit run app.py --browser.serverAddress localhost --server.headless true
pause
