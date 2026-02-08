@echo off
REM ============================================================================
REM Payslip Generator & Distributor - Phuc Long
REM ============================================================================
REM This is a convenience wrapper that calls the master launcher.
REM
REM Usage:
REM   run.bat              - Run with default settings
REM   run.bat [args]       - Pass arguments to the tool
REM ============================================================================

setlocal

REM Get the repository root (two levels up from this script)
set REPO_ROOT=%~dp0..\..
cd /d "%REPO_ROOT%"

REM Call the master launcher with this tool's name
call run.bat payslip-phuclong %*

REM Exit with the same code as the master launcher
exit /b %ERRORLEVEL%
