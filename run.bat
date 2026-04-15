@echo off
cd /d "%~dp0"
echo Starting Starkweather YM Prospector...
echo.
python -m streamlit run app.py
pause
