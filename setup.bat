@echo off
REM ============================================================================
REM Daisy Automation Platform - Setup Script
REM ============================================================================
REM Purpose: Creates virtual environment and installs dependencies
REM Usage:
REM   setup.bat           - Standard setup
REM   setup.bat --recreate - Delete and recreate venv
REM   setup.bat --upgrade  - Upgrade installed packages
REM ============================================================================

setlocal EnableDelayedExpansion

echo.
echo ============================================================================
echo  Daisy Automation Platform - Setup
echo ============================================================================
echo.

REM Parse arguments
set RECREATE=0
set UPGRADE=0

:parse_args
if "%~1"=="" goto args_done
if /i "%~1"=="--recreate" set RECREATE=1
if /i "%~1"=="--upgrade" set UPGRADE=1
shift
goto parse_args
:args_done

REM ============================================================================
REM Step 1: Check Python Installation
REM ============================================================================
echo [1/6] Checking Python installation...

where python >nul 2>nul
if !ERRORLEVEL! neq 0 (
    echo.
    echo ERROR: Python not found in PATH
    echo.
    echo Please install Python 3.14+ from:
    echo   https://www.python.org/downloads/
    echo.
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

REM Get Python version
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo Found Python version: !PYTHON_VERSION!

REM Check version (basic check for 3.x)
echo !PYTHON_VERSION! | findstr /r "^3\.1[0-9]" >nul
if !ERRORLEVEL! neq 0 (
    echo.
    echo WARNING: Python 3.14+ recommended, found !PYTHON_VERSION!
    echo You may encounter compatibility issues.
    echo.
    choice /C YN /M "Continue anyway"
    if !ERRORLEVEL! neq 1 exit /b 1
)
echo.

REM ============================================================================
REM Step 2: Handle Virtual Environment
REM ============================================================================
echo [2/6] Setting up virtual environment...

if !RECREATE! equ 1 (
    if exist "venv" (
        echo Removing existing virtual environment...
        rmdir /s /q venv
        if !ERRORLEVEL! neq 0 (
            echo ERROR: Failed to remove virtual environment
            echo Please close any programs using the venv and try again
            pause
            exit /b 1
        )
        echo Virtual environment removed.
    )
)

if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    if !ERRORLEVEL! neq 0 (
        echo ERROR: Failed to create virtual environment
        echo Please ensure you have sufficient disk space and permissions
        pause
        exit /b 1
    )
    echo Virtual environment created successfully
) else (
    echo Virtual environment already exists
)
echo.

REM ============================================================================
REM Step 3: Activate Virtual Environment
REM ============================================================================
echo [3/6] Activating virtual environment...

if not exist "venv\Scripts\activate.bat" (
    echo ERROR: Virtual environment activation script not found
    echo The venv may be corrupted. Try running: setup.bat --recreate
    pause
    exit /b 1
)

call venv\Scripts\activate.bat
echo.

REM ============================================================================
REM Step 4: Upgrade pip
REM ============================================================================
echo [4/6] Upgrading pip...

python -m pip install --upgrade pip --quiet
if !ERRORLEVEL! neq 0 (
    echo WARNING: Failed to upgrade pip, continuing anyway...
)
echo.

REM ============================================================================
REM Step 5: Install Dependencies
REM ============================================================================
echo [5/6] Installing dependencies...

if !UPGRADE! equ 1 (
    echo Upgrading all packages...
    pip install -r requirements.txt --upgrade
) else (
    pip install -r requirements.txt
)
if !UPGRADE! equ 1 (
    echo Upgrading all packages...
    pip install -r requirements.txt --upgrade
) else (
    pip install -r requirements.txt
)

if !ERRORLEVEL! neq 0 (
    echo ERROR: Failed to install dependencies
    echo Please check your internet connection and requirements.txt
    pause
    exit /b 1
)
echo.

REM ============================================================================
REM Step 6: Validation and Configuration
REM ============================================================================
echo [6/6] Validating installation...

REM Run a quick validation
python -c "from core import Config; from office.outlook import OutlookClient; print('Core modules validated successfully!')" 2>nul
if !ERRORLEVEL! neq 0 (
    echo WARNING: Module import test failed
    echo This may be normal if .env is not yet configured
) else (
    echo Core modules validated successfully!
)
echo.

REM Check if .env exists
if not exist ".env" (
    if exist ".env.example" (
        echo Creating .env from template...
        copy .env.example .env >nul
        echo.
        echo IMPORTANT: Please edit .env and configure your settings
    ) else (
        echo NOTE: .env.example not found at root level
        echo Please configure .env manually if needed
    )
) else (
    echo .env already exists.
)
echo.

REM Deactivate virtual environment
call deactivate

echo ============================================================================
echo  Setup Complete!
echo ============================================================================
echo.
echo Next steps:
echo   1. Edit .env and configure your settings (if needed)
echo   2. Run tools using: run.bat
echo   3. For help: run.bat --help
echo.
echo To update packages in the future: setup.bat --upgrade
echo To recreate environment: setup.bat --recreate
echo.
pause
