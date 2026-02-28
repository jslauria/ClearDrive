@echo off
title OneDrive File Organizer - Microsoft 365 Launcher
color 0A

echo.
echo ========================================================================
echo   OneDrive File Organizer - Microsoft 365 / Business
echo   Unified Version with All Features
echo ========================================================================
echo.

cd /d "%~dp0"

echo [1/3] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    color 0C
    echo ERROR: Python not found
    echo Install from: https://www.python.org/downloads/
    pause
    exit /b 1
)
python --version
echo     OK
echo.

echo [2/3] Installing dependencies...
if exist requirements.txt (
    python -m pip install -r requirements.txt --quiet --upgrade
    echo     OK
) else (
    python -m pip install requests anthropic python-docx python-pptx PyPDF2 openpyxl --quiet
)
echo.

echo [3/3] Starting application...
echo.

python onedrive_organizer_unified_gui.py

if errorlevel 1 (
    color 0C
    echo.
    echo ERROR: Failed to start
    pause
    exit /b 1
)
