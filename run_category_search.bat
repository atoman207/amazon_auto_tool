@echo off
REM ============================================================================
REM Amazon Category Search - Run Script
REM ============================================================================

echo.
echo ========================================================================
echo          AMAZON CATEGORY SEARCH AUTOMATION
echo ========================================================================
echo.

REM Change to script directory
cd /d "%~dp0"

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH!
    echo.
    echo Please install Python 3.8 or higher from:
    echo https://www.python.org/downloads/
    echo.
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo [INFO] Python found
echo.

REM Check if .env file exists
if not exist ".env" (
    echo [ERROR] .env file not found!
    echo.
    echo Please create a .env file with your Amazon credentials:
    echo   AMAZON_EMAIL=your_email@gmail.com
    echo   AMAZON_PASSWORD=your_password
    echo.
    pause
    exit /b 1
)

echo [INFO] .env file found
echo.

REM Check if Gmail credentials exist
if not exist "data\client_secret_446842116198-amdijg8d7tb7rff25o4514r19pp1d8o9.apps.googleusercontent.com.json" (
    echo [ERROR] Gmail API credentials not found!
    echo.
    echo Please download OAuth 2.0 credentials from:
    echo https://console.cloud.google.com/apis/credentials
    echo.
    echo And save it to the 'data' folder.
    echo.
    pause
    exit /b 1
)

echo [INFO] Gmail credentials found
echo.

REM Run the category search script
echo ========================================================================
echo Starting category search automation...
echo ========================================================================
echo.

python category_search.py

REM Check exit code
if errorlevel 1 (
    echo.
    echo ========================================================================
    echo [ERROR] Category search failed!
    echo ========================================================================
    echo.
    pause
    exit /b 1
) else (
    echo.
    echo ========================================================================
    echo [SUCCESS] Category search completed!
    echo ========================================================================
    echo.
    pause
    exit /b 0
)
