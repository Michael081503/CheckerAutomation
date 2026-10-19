@echo off
setlocal
cd /d "%~dp0"
python -c "import streamlit, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Python, Streamlit, and openpyxl are required.
    echo Install Python, then run: python -m pip install -r requirements.txt
    pause
    exit /b 1
)
python -m streamlit run streamlit_app.py
if errorlevel 1 pause
