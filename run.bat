@echo off
REM ================================================================================================
REM Daisy Automation Platform - Master Launcher
REM ================================================================================================
REM Usage:
REM   run.bat                    - Show interactive menu
REM   run.bat --help             - Show help
REM   run.bat <tool-name>        - Run a specific tool
REM   run.bat <script-name>      - Run a specific script
REM
REM Examples:
REM   run.bat                    (Shows menu)
REM   run.bat payslip-phuclong   (Runs payslip-phuclong tool)
REM   run.bat example_read_emails (Runs script from scripts/)
REM ================================================================================================

setlocal EnableDelayedExpansion

REM ================================================================================================
REM Step 1: Check Virtual Environment
REM ================================================================================================
if not exist "venv\Scripts\activate.bat" (
    echo.
    echo ERROR: Virtual environment not found!
    echo.
    echo Please run setup.bat first to create the virtual environment:
    echo   setup.bat
    echo.
    pause
    exit /b 1
)

REM ================================================================================================
REM Step 2: Activate Virtual Environment
REM ================================================================================================
call venv\Scripts\activate.bat

REM ================================================================================================
REM Step 3: Parse Arguments
REM ================================================================================================

REM If no arguments, show menu
if "%~1"=="" goto show_menu

REM If --help flag
if /i "%~1"=="--help" goto show_help
if /i "%~1"=="-h" goto show_help
if /i "%~1"=="/?" goto show_help

REM Otherwise, try to run the specified tool or script
set TARGET=%~1
shift

REM Collect remaining arguments
set ARGS=
:collect_args
if "%~1"=="" goto args_done
set ARGS=!ARGS! %1
shift
goto collect_args
:args_done

goto run_target

REM ================================================================================================
REM Interactive Menu
REM ================================================================================================
:show_menu
cls
echo.
echo ============================================================================================
echo  Daisy Automation Platform - Main Menu  
echo ============================================================================================
echo.
echo Available Tools:
echo.

REM List available tools
set TOOL_COUNT=0
if exist "tools\" (
    for /d %%D in (tools\*) do (
        if exist "tools\%%~nxD\main.py" (
            set /a TOOL_COUNT+=1
            echo   [!TOOL_COUNT!] %%~nxD
            set TOOL_!TOOL_COUNT!=%%~nxD
        )
    )
)

if !TOOL_COUNT! equ 0 (
    echo   No tools found in tools/ directory
)

echo.
echo Available Scripts:
echo.

REM List available scripts
set SCRIPT_COUNT=0
if exist "scripts\" (
    for %%F in (scripts\*.py) do (
        set FILENAME=%%~nF
        if not "!FILENAME:~0,2!"=="__" (
            set /a SCRIPT_COUNT+=1
            set /a DISPLAY_NUM=!TOOL_COUNT!+!SCRIPT_COUNT!
            echo   [!DISPLAY_NUM!] !FILENAME!
            set SCRIPT_!SCRIPT_COUNT!=!FILENAME!
        )
    )
)

if !SCRIPT_COUNT! equ 0 (
    echo   No scripts found in scripts/ directory
)

echo.
echo   [0] Exit
echo.
echo ============================================================================================
echo.

set /p CHOICE="Select an option (0-%DISPLAY_NUM%): "

if "!CHOICE!"=="0" goto cleanup_exit

REM Validate choice
set /a CHOICE_NUM=!CHOICE! 2>nul
if !CHOICE_NUM! gtr %DISPLAY_NUM% (
    echo Invalid choice!
    timeout /t 2 >nul
    goto show_menu
)
if !CHOICE_NUM! leq 0 (
    echo Invalid choice!
    timeout /t 2 >nul
    goto show_menu
)

REM Run selected tool or script
if !CHOICE_NUM! leq !TOOL_COUNT! (
    set TARGET=!TOOL_%CHOICE_NUM%!
    echo.
    echo Running tool: !TARGET!
    echo.
    python tools\!TARGET!\main.py
) else (
    set /a SCRIPT_IDX=!CHOICE_NUM!-!TOOL_COUNT!
    set TARGET=!SCRIPT_%SCRIPT_IDX%!
    echo.
    echo Running script: !TARGET!
    echo.
    python scripts\!TARGET!.py
)

echo.
echo ============================================================================================
pause
goto cleanup_exit

REM ================================================================================================
REM Help Display
REM ================================================================================================
:show_help
echo.
echo Daisy Automation Platform - Master Launcher
echo.
echo Usage:
echo   run.bat                    Show interactive menu
echo   run.bat --help             Show this help
echo   run.bat ^<tool-name^>        Run a specific tool
echo   run.bat ^<script-name^>      Run a specific script
echo.
echo Available Tools:
if exist "tools\" (
    for /d %%D in (tools\*) do (
        if exist "tools\%%~nxD\main.py" (
            echo   - %%~nxD
        )
    )
)
echo.
echo Available Scripts:
if exist "scripts\" (
    for %%F in (scripts\*.py) do (
        set FILENAME=%%~nF
        if not "!FILENAME:~0,2!"=="__" (
            echo   - %%~nF
        )
    )
)
echo.
echo Examples:
echo   run.bat                           (interactive menu)
echo   run.bat payslip-phuclong          (run tool)
echo   run.bat example_read_emails       (run script)
echo.
goto cleanup_exit

REM ================================================================================================
REM Run Target (Tool or Script)
REM ================================================================================================
:run_target

REM Check if it's a tool
if exist "tools\%TARGET%\main.py" (
    echo.
    echo Running tool: %TARGET%
    echo.
    python tools\%TARGET%\main.py !ARGS!
    set EXIT_CODE=!ERRORLEVEL!
    goto cleanup_exit
)

REM Check if it's a script
if exist "scripts\%TARGET%.py" (
    echo.
    echo Running script: %TARGET%
    echo.
    python scripts\%TARGET%.py !ARGS!
    set EXIT_CODE=!ERRORLEVEL!
    goto cleanup_exit
)

REM Target not found
echo.
echo ERROR: Tool or script '%TARGET%' not found!
echo.
echo Run 'run.bat --help' to see available tools and scripts.
echo.
set EXIT_CODE=1
pause
goto cleanup_exit

REM ================================================================================================
REM Cleanup and Exit
REM ================================================================================================
:cleanup_exit
call deactivate
if defined EXIT_CODE (
    exit /b !EXIT_CODE!
) else (
    exit /b 0
)
