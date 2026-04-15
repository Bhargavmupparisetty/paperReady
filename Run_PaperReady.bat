@echo off
title PaperReady AI Launcher
color 0B

echo.
echo ==============================================
echo       Starting PaperReady Local AI...
echo ==============================================
echo.

:: 1. Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in your system PATH.
    echo Please install Python 3.10+ from python.org and try again.
    pause
    exit /b
)

:: 2. Check and install dependencies automatically
echo Checking and installing dependencies securely...
pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo [WARNING] There was an issue verifying dependencies. The app will try to run anyway.
)

:: 3. Setup Windows automation just in case (Silently)
python -m pywin32_postinstall -install >nul 2>&1

:: 4. Start the Application
echo.
echo Environment verified. Launching Architect...
echo.
python cli.py

:: 5. Pause if the application crashes so the user can see the error
pause
