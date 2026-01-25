@echo off
chcp 65001 >nul
title Fix Installation Issues
color 0B

echo ============================================================
echo     FIXING INSTALLATION ISSUES
echo ============================================================
echo.

REM Change to script directory
cd /d "%~dp0"

REM Activate virtual environment
echo [INFO] Activating virtual environment...
call venv\Scripts\activate.bat
echo.

REM Upgrade pip, setuptools, and wheel
echo [INFO] Upgrading pip, setuptools, and wheel...
python -m pip install --upgrade pip setuptools wheel
echo.

REM Try to install greenlet with pre-built wheel
echo [INFO] Installing greenlet with pre-built wheel...
pip install greenlet --prefer-binary
echo.

REM Install other dependencies
echo [INFO] Installing remaining dependencies...
pip install playwright python-dotenv google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client gspread
echo.

REM Install Playwright browsers
echo [INFO] Installing Playwright browsers...
python -m playwright install chromium
echo.

echo ============================================================
echo     INSTALLATION FIXED
echo ============================================================
echo.
echo You can now run: run.bat
echo.
pause
