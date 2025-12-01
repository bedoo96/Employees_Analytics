@echo off
REM Employee Attendance Analytics - Windows Startup Script

echo.
echo Starting Employee Attendance Analytics System...
echo.

REM Check if virtual environment exists
if not exist "venv\" (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Install/upgrade dependencies
echo Installing dependencies...
pip install -r requirements.txt --quiet

REM Check if .env exists
if not exist ".env" (
    echo.
    echo WARNING: .env file not found!
    echo Please copy .env.example to .env and configure your API keys
    echo.
)

REM Clear screen
cls

REM Start the application
echo.
echo ========================================
echo  Employee Attendance Analytics System
echo ========================================
echo.
echo Access the application at: http://localhost:8501
echo.
echo Press Ctrl+C to stop the server
echo.

streamlit run app.py