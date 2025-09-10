@echo off
echo Starting BidAnalysis Tool...
echo.

:: Check if virtual environment exists
if not exist ".venv\Scripts\activate.bat" (
    echo Creating virtual environment...
    python -m venv .venv
)

:: Activate virtual environment
call .venv\Scripts\activate.bat

:: Install dependencies
echo Installing dependencies...
pip install -r backend\requirements.txt

:: Check if .env file exists
if not exist "backend\.env" (
    echo Creating .env file from template...
    copy "backend\.env.template" "backend\.env"
    echo.
    echo Please edit backend\.env file and add your DASHSCOPE_API_KEY
    echo Then run this script again.
    pause
    exit /b 1
)

:: Start the application
echo Starting backend server...
cd backend
python run.py
