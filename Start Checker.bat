@echo off
setlocal
cd /d "%~dp0"
set "CHECKER_PYTHON=python"
if exist ".venv\Scripts\python.exe" set "CHECKER_PYTHON=%~dp0.venv\Scripts\python.exe"
"%CHECKER_PYTHON%" -c "import streamlit, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Python, Streamlit, and openpyxl are required.
    echo Install Python, then run: python -m pip install -r requirements.txt
    pause
    exit /b 1
)
"%CHECKER_PYTHON%" -m streamlit run streamlit_app.py
if errorlevel 1 pause
