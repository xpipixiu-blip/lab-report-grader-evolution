@echo off
title Lab Report Grader Evolution
cd /d "%~dp0"

echo ============================================
echo   Lab Report Grader Evolution
echo ============================================
echo Starting at http://127.0.0.1:7870
echo Press Ctrl+C to stop.
echo.

python -c "import gradio, openai, docx, fitz, openpyxl, yaml" >nul 2>nul
if errorlevel 1 (
  echo Required packages are missing. Installing them now...
  python -m pip install -r requirements.txt
  if errorlevel 1 (
    echo Dependency installation failed. Check Python and network access.
    pause
    exit /b 1
  )
)

python app.py

echo.
echo Program stopped. Please keep this window when reporting an error.
pause
