@echo off
title Prop65 IMDS Tool Installer

echo ==========================================
echo Installing required Python packages...
echo ==========================================

python --version
if errorlevel 1 (
    echo.
    echo Python is not installed.
    echo Please install Python from:
    echo https://www.python.org/downloads/
    pause
    exit /b
)

pip install -r requirements.txt

echo.
echo ==========================================
echo Starting local web application...
echo ==========================================

start http://localhost:8501

streamlit run app.py

pause
