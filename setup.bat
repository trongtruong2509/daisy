@echo off
REM Office Automation Foundation - Setup Script
REM
REM This script creates a virtual environment and installs dependencies.
REM Run this once when setting up the project.

setlocal

echo ========================================
echo Office Automation Foundation - Setup
echo ========================================
echo.

REM Check if Python is available
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: Python not found in PATH
    echo Please install Python 3.9+ from python.org
    exit /b 1
)

REM Show Python version
echo Python version:
python --version
echo.

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    if %ERRORLEVEL% neq 0 (
        echo ERROR: Failed to create virtual environment
        exit /b 1
    )
    echo Virtual environment created.
) else (
    echo Virtual environment already exists.
)
echo.

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo ERROR: Failed to install dependencies
    exit /b 1
)
echo.

REM Check if .env exists
if not exist ".env" (
    echo Creating .env from template...
    copy .env.example .env
    echo.
    echo IMPORTANT: Please edit .env and set your OUTLOOK_ACCOUNT
) else (
    echo .env already exists.
)
echo.

REM Run a quick validation
echo Validating installation...
python -c "from core import Config; from office.outlook import OutlookClient; print('All modules imported successfully!')"
if %ERRORLEVEL% neq 0 (
    echo WARNING: Module import test failed
    echo Please check the error messages above
) else (
    echo Installation validated successfully!
)
echo.

echo ========================================
echo Setup complete!
echo ========================================
echo.
echo Next steps:
echo 1. Edit .env and set OUTLOOK_ACCOUNT to your email address
echo 2. Make sure Outlook Desktop is running
echo 3. Run: run.bat example_read_emails
echo.

call deactivate
