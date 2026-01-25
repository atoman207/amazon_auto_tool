@echo off
chcp 65001 >nul
title Amazon Automation Tool
color 0A

echo ============================================================
echo     AMAZON AUTOMATION TOOL - STARTING
echo ============================================================
echo.

REM Change to script directory
cd /d "%~dp0"

REM Check if virtual environment exists
if exist "venv\Scripts\activate.bat" (
    echo [INFO] Activating virtual environment...
    call venv\Scripts\activate.bat
    echo [SUCCESS] Virtual environment activated
    echo.
) else if exist ".venv\Scripts\activate.bat" (
    echo [INFO] Activating virtual environment...
    call .venv\Scripts\activate.bat
    echo [SUCCESS] Virtual environment activated
    echo.
) else (
    echo [INFO] No virtual environment found, using system Python
    echo.
)

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo [ERROR] Please install Python 3.8 or higher
    echo.
    pause
    exit /b 1
)

echo [INFO] Python version:
python --version
echo.

REM Check if requirements are installed
echo [INFO] Checking dependencies...
python -c "import playwright" >nul 2>&1
if errorlevel 1 (
    echo [WARNING] Dependencies not installed
    echo [INFO] Installing requirements...
    python -m pip install -r requirements.txt
    echo.
    echo [INFO] Installing Playwright browsers...
    python -m playwright install chromium
    echo.
)

REM Check if .env file exists
if not exist ".env" (
    echo [WARNING] .env file not found
    echo [INFO] Please create .env file with your Amazon password
    echo [INFO] Example: AMAZON_PASSWORD=your_password_here
    echo.
    pause
    exit /b 1
)

echo ============================================================
echo     RUNNING AMAZON AUTOMATION
echo ============================================================
echo.

REM Run the automation script
python amazon_auto.py

REM Check if script ran successfully
if errorlevel 1 (
    echo.
    echo ============================================================
    echo [ERROR] Automation failed with errors
    echo ============================================================
) else (
    echo.
    echo ============================================================
    echo [SUCCESS] Automation completed successfully
    echo ============================================================
)

echo.
echo Press any key to exit...
pause >nul
