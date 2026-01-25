@echo off
chcp 65001 >nul
title Amazon Automation Tool - Setup
color 0B

echo ============================================================
echo     AMAZON AUTOMATION TOOL - INITIAL SETUP
echo ============================================================
echo.

REM Change to script directory
cd /d "%~dp0"

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo [ERROR] Please install Python 3.8 or higher from https://www.python.org/
    echo.
    pause
    exit /b 1
)

echo [INFO] Python version:
python --version
echo.

REM Create virtual environment
if not exist "venv" (
    echo [INFO] Creating virtual environment...
    python -m venv venv
    echo [SUCCESS] Virtual environment created
    echo.
) else (
    echo [INFO] Virtual environment already exists
    echo.
)

REM Activate virtual environment
echo [INFO] Activating virtual environment...
call venv\Scripts\activate.bat
echo.

REM Upgrade pip, setuptools, and wheel
echo [INFO] Upgrading pip, setuptools, and wheel...
python -m pip install --upgrade pip setuptools wheel
echo.

REM Install requirements
echo [INFO] Installing Python dependencies...
echo [INFO] This may take a few minutes...
pip install playwright python-dotenv
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client
pip install gspread --prefer-binary
echo [SUCCESS] Dependencies installed
echo.

REM Install Playwright browsers
echo [INFO] Installing Playwright browsers (this may take a few minutes)...
python -m playwright install chromium
echo.

REM Create .env file if it doesn't exist
if not exist ".env" (
    echo [INFO] Creating .env file...
    if exist ".env.example" (
        copy .env.example .env
        echo [SUCCESS] .env file created from .env.example
        echo [WARNING] Please edit .env file and add your Amazon password
    ) else (
        echo AMAZON_PASSWORD=your_password_here > .env
        echo [SUCCESS] .env file created
        echo [WARNING] Please edit .env file and add your Amazon password
    )
    echo.
) else (
    echo [INFO] .env file already exists
    echo.
)

REM Check Gmail credentials
if not exist "data\client_secret_*.json" (
    echo [WARNING] Gmail API credentials not found in data\ folder
    echo [INFO] Please add your client_secret JSON file to the data\ folder
    echo [INFO] Get it from: https://console.cloud.google.com/
    echo.
)

echo ============================================================
echo     SETUP COMPLETED
echo ============================================================
echo.
echo Next steps:
echo   1. Edit .env file and set your AMAZON_PASSWORD
echo   2. Make sure Gmail API credentials are in data\ folder
echo   3. Run the automation using run.bat
echo.
echo ============================================================
pause
