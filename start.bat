@echo off
setlocal ENABLEDELAYEDEXPANSION
echo Starting BidAnalysis Tool (Python 3.12 preferred)...
echo.

:: Prefer Python 3.12 virtual environment for COM/Word support
if not exist ".venv312\Scripts\activate.bat" (
    echo Creating Python 3.12 virtual environment...
    py -3.12 -m venv .venv312
    if %errorlevel% neq 0 (
        echo.
        echo [ERROR] Python 3.12 not found. Please install Python 3.12 first.
        echo You can install via: winget install -e --id Python.Python.3.12
        pause
        exit /b 1
    )
)

:: Activate Python 3.12 environment
call .venv312\Scripts\activate.bat

:: Install dependencies
echo Installing dependencies...
pip install -r backend\requirements.txt
:: Ensure pywin32 for COM features
pip install pywin32==306 >nul 2>&1

:: Check if .env file exists
if not exist "backend\.env" (
    echo Creating .env file from template...
    copy "backend\.env.template" "backend\.env"
    echo.
    echo Please edit backend\.env file and set API keys (DASHSCOPE_API_KEY or ARK_API_KEY)
    echo Then run this script again.
    pause
    exit /b 1
)

:: Start the application
echo Starting backend server...
cd backend
python run.py
