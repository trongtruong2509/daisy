@echo off
REM Office Automation Foundation - Run Script
REM
REM Usage: run.bat <script_name>
REM Example: run.bat example_read_emails
REM
REM This batch file activates the virtual environment and runs the specified script.

setlocal

REM Check if virtual environment exists
if not exist "venv\Scripts\activate.bat" (
    echo Virtual environment not found. Please run setup.bat first.
    exit /b 1
)

REM Activate virtual environment
call venv\Scripts\activate.bat

REM Check if script name was provided
if "%1"=="" (
    echo Usage: run.bat ^<script_name^>
    echo.
    echo Available scripts:
    dir /b scripts\*.py 2>nul | findstr /v "__"
    exit /b 1
)

REM Run the specified script
python scripts\%1.py %2 %3 %4 %5 %6 %7 %8 %9

REM Capture exit code
set EXIT_CODE=%ERRORLEVEL%

REM Deactivate virtual environment
call deactivate

exit /b %EXIT_CODE%
