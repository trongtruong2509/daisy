# Virtual Environment Strategy and Tool Launcher Architecture

**Document Version:** 1.0  
**Date:** 2026-02-08  
**Status:** Proposed Recommendation  
**Author:** Platform/DevOps Team

---

## Table of Contents

1. [Executive Summary](#executive-summary)
2. [Context and Requirements](#context-and-requirements)
3. [Detailed Options Analysis](#detailed-options-analysis)
4. [Recommended Approach](#recommended-approach)
5. [User Personas and Workflows](#user-personas-and-workflows)
6. [Detailed Implementation Plan](#detailed-implementation-plan)
7. [Code Samples and Templates](#code-samples-and-templates)
8. [Testing and Validation](#testing-and-validation)
9. [Troubleshooting Guide](#troubleshooting-guide)
10. [Migration and Rollout Strategy](#migration-and-rollout-strategy)
11. [Maintenance and Best Practices](#maintenance-and-best-practices)
12. [Future Considerations](#future-considerations)
13. [Decision Matrix](#decision-matrix)
14. [Appendices](#appendices)

---

## Executive Summary

This document provides a comprehensive strategy for managing Python virtual environments and tool launchers for the Daisy automation platform. The platform hosts multiple automation tools under a single repository, targeting non-technical users (HR, managers) on Windows who need simple double-click execution.

**Key Decision:** Implement a **single shared virtual environment** at repository root with a **master launcher pattern** that supports both interactive menu and direct tool invocation.

**Benefits:**
- One-time setup reduces user friction
- Consistent dependency management across all tools
- Lower disk usage and maintenance overhead
- Easy upgrade path for all tools simultaneously
- Simple backup and restore procedures

**Trade-offs:**
- Potential dependency conflicts if tools diverge significantly (mitigated by shared codebase)
- All tools affected if venv becomes corrupted (mitigated by quick recreation)

---

## Context and Requirements

### Platform and User Environment

| Aspect | Specification |
|--------|--------------|
| **Operating System** | Windows 10/11 |
| **Python Version** | 3.14 (latest) |
| **User Profile** | Non-technical (HR, managers, admin staff) |
| **Execution Model** | Double-click `.bat` files |
| **Admin Rights** | Not required |
| **Network** | Corporate network with potential restrictions |
| **Update Frequency** | Monthly to quarterly |

### Repository Structure

```
daisy/
├── venv/                    # Virtual environment (gitignored)
├── setup.bat                # One-time setup script
├── run.bat                  # Master launcher script
├── requirements.txt         # Shared dependencies
├── .gitignore              # Excludes venv/
├── core/                   # Shared core modules
│   ├── config.py
│   ├── logger.py
│   └── state.py
├── office/                 # Office automation modules
│   └── outlook/
├── tools/                  # Tool collection
│   ├── payslip-phuclong/
│   │   ├── main.py
│   │   ├── config.py
│   │   ├── run.bat         # Per-tool wrapper
│   │   ├── .env.example
│   │   └── README.md
│   └── [future-tools]/
└── docs/
    └── venv-strategy.md    # This document
```

### Design Goals and Constraints

**Primary Goals:**
1. **Simplicity:** Non-technical users should be able to run tools without understanding Python, pip, or virtual environments
2. **Reliability:** Tools must work consistently across different user machines
3. **Maintainability:** Developers should manage dependencies in one place
4. **Portability:** Users should be able to clone, setup, and run without admin rights

**Constraints:**
1. No system-wide Python package installations allowed
2. Tools share common codebase (`core/`, `office/`, `parsing/`)
3. Virtual environment should not be committed to git
4. Setup process must be recoverable if corrupted
5. Must support air-gapped or restricted network environments (optional offline install)

---

## Detailed Options Analysis

### Option 1: Single Shared Virtual Environment (Recommended)

**Architecture:**
```
daisy/
├── venv/                    # Single shared venv
│   ├── Scripts/
│   │   ├── python.exe
│   │   ├── activate.bat
│   │   └── pip.exe
│   └── Lib/site-packages/
├── setup.bat                # Creates venv and installs all deps
└── run.bat                  # Activates venv and runs tool
```

**Detailed Pros:**
- **User Experience:** One setup command installs everything; users never think about venv
- **Disk Efficiency:** ~100-300MB for one venv vs. 100-300MB × N tools
- **Update Simplicity:** `pip install -r requirements.txt --upgrade` updates all tools
- **Consistency:** All tools guaranteed to use identical package versions
- **Developer Productivity:** Single `requirements.txt` to manage
- **Backup/Restore:** Easy to backup entire venv folder for offline deployment
- **Troubleshooting:** Single point of failure is easier to diagnose and fix

**Detailed Cons:**
- **Dependency Conflicts:** If tools later need incompatible package versions (e.g., pandas 1.x vs 2.x), this approach breaks
- **Blast Radius:** A corrupted venv affects all tools (mitigated by quick `setup.bat --recreate`)
- **Testing Isolation:** Cannot test one tool with different package versions without recreating entire venv
- **Large Dependency Set:** If some tools have heavy dependencies, all users pay the disk cost

**Risk Assessment:** **Low** - Tools share the same codebase, making conflicts unlikely in near term.

---

### Option 2: Per-Tool Virtual Environments

**Architecture:**
```
daisy/
├── tools/
│   ├── payslip-phuclong/
│   │   ├── venv/            # Tool-specific venv
│   │   ├── main.py
│   │   ├── requirements.txt # Tool-specific deps
│   │   └── setup.bat        # Tool-specific setup
│   └── excel-analysis/
│       ├── venv/
│       ├── main.py
│       ├── requirements.txt
│       └── setup.bat
```

**Detailed Pros:**
- **Complete Isolation:** Each tool's dependencies are independent
- **Version Flexibility:** Tool A can use pandas 1.5, Tool B can use pandas 2.0
- **Blast Radius Containment:** Corrupted venv only affects one tool
- **Independent Updates:** Update one tool without affecting others
- **Clear Ownership:** Each tool owns its full dependency stack
- **Testing:** Easy to test different package versions per tool

**Detailed Cons:**
- **Disk Usage:** If 10 tools exist, 1-3GB total venv storage
- **User Complexity:** Users must run setup for each tool individually
- **Maintenance Overhead:** Developers manage N `requirements.txt` files
- **Inconsistent Versions:** Core modules (`core/`) might behave differently across tools
- **Update Burden:** Security updates require N separate pip commands
- **Documentation:** Must explain multiple setup procedures

**Risk Assessment:** **Medium** - Complexity outweighs benefits unless tools truly have conflicting needs.

---

### Option 3: Central Shared Virtual Environment (Outside Repo)

**Architecture:**
```
%LOCALAPPDATA%/
└── daisy-venv/              # System-wide shared venv
    ├── Scripts/
    └── Lib/

daisy/                       # Clean repo
├── setup.bat                # Points to central venv
└── run.bat                  # Uses central venv
```

**Detailed Pros:**
- **Clean Repository:** No venv files in repo at all
- **Shared Across Clones:** Multiple repo clones share one venv
- **Corporate Standard:** Fits managed environments where IT controls venv location
- **Disk Efficiency:** One venv serves multiple repo copies
- **Centralized Updates:** IT can update venv centrally

**Detailed Cons:**
- **Path Complexity:** Scripts must handle absolute paths to external venv
- **User Confusion:** "Where is my virtual environment?" becomes a support question
- **Repo Portability:** Cannot move repo folder without breaking venv references
- **Version Conflicts:** If multiple repo versions exist, venv version becomes ambiguous
- **Corporate Restrictions:** Some users may not have write access to `%LOCALAPPDATA%`

**Risk Assessment:** **Medium** - Adds complexity without clear benefit for current use case.

---

### Option 4: pipx or Packaged Executables

**Architecture:**
```
# Each tool packaged as installable CLI
pipx install daisy-payslip
pipx install daisy-excel-analysis

# Or built as exe
tools/payslip-phuclong.exe
```

**Detailed Pros:**
- **Professional Distribution:** Tools become system commands
- **Clean Isolation:** Each tool in its own isolated environment
- **No venv Management:** Users never see or touch venv
- **Easy Updates:** `pipx upgrade tool-name`
- **Cross-Machine Portability:** Can publish to internal PyPI

**Detailed Cons:**
- **Packaging Overhead:** Must create `setup.py`, version management, build pipeline
- **Development Complexity:** Harder to iterate quickly during active development
- **Exe Limitations:** Windows executables can be flagged by antivirus
- **Size:** Each exe bundles Python runtime (~30-50MB per tool)
- **Not Appropriate for Current Stage:** Overkill for tools still under active development

**Risk Assessment:** **Low Risk, High Effort** - Worth considering once tools mature.

---

## Recommended Approach

### Selection: Option 1 - Single Shared Virtual Environment

**Rationale:**

1. **Shared Codebase Reality:** All tools use `core/`, `office/`, `parsing/` modules. They must have consistent package versions to avoid subtle bugs.

2. **User Profile:** Non-technical users need the simplest possible workflow:
   - Run `setup.bat` once
   - Double-click any tool to run

3. **Maintenance Efficiency:** Managing one `requirements.txt` is significantly easier than N files, especially for security updates.

4. **Current Scale:** With 1-5 tools expected, disk and complexity overhead of per-tool venvs is not justified.

5. **Migration Path:** If conflicts arise, we can migrate specific tools to per-tool venvs later.

### Architecture Components

```
┌─────────────────────────────────────────────────────────┐
│                    User Machine                          │
│                                                          │
│  ┌─────────────┐                                        │
│  │  setup.bat  │  (Run once)                            │
│  └──────┬──────┘                                        │
│         │                                                │
│         ├──> Creates venv/                              │
│         ├──> Installs requirements.txt                  │
│         └──> Creates .env from .env.example             │
│                                                          │
│  ┌─────────────┐         ┌─────────────────┐          │
│  │ run.bat     │◄────────│ Tool Wrappers   │          │
│  │ (Master)    │         │ (Per-tool)      │          │
│  └──────┬──────┘         └─────────────────┘          │
│         │                                                │
│         ├──> Activates venv                             │
│         ├──> Shows menu (if no args)                    │
│         ├──> Dispatches to tools/X/main.py              │
│         └──> Deactivates venv on exit                   │
│                                                          │
└─────────────────────────────────────────────────────────┘
```

---

## User Personas and Workflows

### Persona 1: HR Manager (Non-Technical)

**Profile:**
- Runs payslip generation monthly
- No Python knowledge
- Comfortable with Excel and email
- Needs clear error messages

**Workflow:**

**First Time (Setup):**
1. Receives repository folder from IT or development team
2. Opens folder in Windows Explorer
3. Double-clicks `setup.bat`
4. Waits 2-5 minutes while setup completes
5. Reads on-screen instructions about editing `.env` file

**Monthly Usage:**
1. Opens repository folder
2. Double-clicks `tools/payslip-phuclong/run.bat`
3. Follows on-screen prompts
4. Reviews output and generated payslips

**Error Recovery:**
- If error occurs: Takes screenshot, contacts support
- If "venv not found" error: Runs `setup.bat` again

---

### Persona 2: IT Support Staff (Technical)

**Profile:**
- Deploys tools to end users
- Basic Python knowledge
- Handles troubleshooting

**Workflow:**

**Deployment:**
1. Clones repository to shared network drive
2. Runs `setup.bat` to create venv
3. Tests each tool
4. Creates desktop shortcuts for end users
5. Documents any environment-specific configuration

**Troubleshooting:**
1. Reviews error messages from users
2. Checks venv exists: `venv/Scripts/python.exe --version`
3. Validates dependencies: `venv/Scripts/pip list`
4. Recreates venv if needed: `setup.bat --recreate`
5. Checks `.env` configuration

**Updates:**
1. Receives updated repository
2. Runs `setup.bat --upgrade` to update packages
3. Tests all tools
4. Notifies users of updates

---

### Persona 3: Developer (Technical)

**Profile:**
- Develops new tools
- Expert Python knowledge
- Maintains repository

**Workflow:**

**Development:**
1. Clones repository
2. Runs `setup.bat` to create dev environment
3. Activates venv manually: `venv\Scripts\activate.bat`
4. Develops and tests
5. Updates `requirements.txt` as needed
6. Commits code (venv is gitignored)

**Adding New Tool:**
1. Creates `tools/new-tool/` folder
2. Adds `main.py` and other modules
3. Creates `tools/new-tool/run.bat` wrapper
4. Updates documentation
5. Tests end-to-end with `run.bat new-tool`
6. Commits changes

**Releasing:**
1. Updates version numbers
2. Tests all tools with current venv
3. Creates release tag
4. Deploys to production environment

---

## Detailed Implementation Plan

### Phase 1: Core Infrastructure (Week 1)

#### Task 1.1: Update setup.bat
**Objective:** Enhance existing `setup.bat` with Python 3.14 checks, recreate option, and better error handling.

**Changes:**
- Add Python version check for 3.14
- Add `--recreate` flag to delete and rebuild venv
- Add `--upgrade` flag to upgrade existing packages
- Improve error messages
- Add validation checks post-install

**Estimated Time:** 2 hours

---

#### Task 1.2: Create .gitignore Entry
**Objective:** Ensure venv is never committed.

**Changes:**
- Add `venv/` to `.gitignore`
- Add `**/__pycache__/` if not present
- Add `**/*.pyc` if not present
- Add `.env` (keep `.env.example`)

**Estimated Time:** 15 minutes

---

#### Task 1.3: Create Master Launcher (run.bat)
**Objective:** Implement central launcher that handles all tool invocation.

**Features:**
- Interactive menu when run without arguments
- Direct tool invocation: `run.bat tool-name arg1 arg2`
- Auto-discovery of tools in `tools/` folder
- Error handling and user-friendly messages
- Logging support

**Estimated Time:** 4 hours

---

### Phase 2: Per-Tool Integration (Week 1-2)

#### Task 2.1: Create Tool Wrapper for payslip-phuclong
**Objective:** Create `tools/payslip-phuclong/run.bat`

**Features:**
- Calls master launcher with tool name
- Handles relative paths correctly
- Includes tool-specific help text
- Pauses on error for user to read

**Estimated Time:** 1 hour

---

#### Task 2.2: Create Tool Template
**Objective:** Standardized template for future tools.

**Deliverables:**
- `tools/_template/` folder with skeleton structure
- Template `run.bat`
- Template `README.md`
- Template `.env.example`
- Template `main.py` with argparse

**Estimated Time:** 2 hours

---

### Phase 3: Testing and Documentation (Week 2)

#### Task 3.1: End-to-End Testing
**Test Scenarios:**
1. Fresh clone → setup → run each tool
2. Corrupt venv → setup --recreate
3. Missing Python → error message
4. Wrong Python version → error message
5. Double-click without setup → clear error
6. Multiple tools run in sequence
7. Tool with missing configuration

**Estimated Time:** 4 hours

---

#### Task 3.2: Update Documentation
**Documents to Update:**
- `README.md`: Add setup and usage instructions
- `docs/venv-strategy.md`: This document
- Per-tool READMEs
- Troubleshooting guide

**Estimated Time:** 3 hours

---

### Phase 4: Rollout (Week 3)

#### Task 4.1: Pilot with IT Support
- Train 2-3 IT staff
- Deploy to test environment
- Gather feedback
- Iterate on user experience

**Estimated Time:** 1 week

---

#### Task 4.2: User Training
- Create video walkthrough
- Write one-page quick start guide
- Conduct training session
- Set up support ticketing

**Estimated Time:** 1 week

---

## Code Samples and Templates

### Enhanced setup.bat

```bat
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
    echo Please install Python 3.14 from:
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
    choice /C YN /M "Continue anyway?"
    if !ERRORLEVEL! neq 1 exit /b 1
)
echo.

REM ============================================================================
REM Step 2: Handle Virtual Environment
REM ============================================================================
echo [2/6] Setting up virtual environment...

if !RECREATE! equ 1 (
    if exist "venv" (
        echo Deleting existing virtual environment...
        rmdir /s /q venv
        if !ERRORLEVEL! neq 0 (
            echo ERROR: Failed to delete existing venv
            echo Please close any programs using files in venv\ and try again
            pause
            exit /b 1
        )
    )
)

if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    if !ERRORLEVEL! neq 0 (
        echo.
        echo ERROR: Failed to create virtual environment
        echo.
        echo Possible causes:
        echo   - Insufficient disk space
        echo   - Antivirus blocking file creation
        echo   - Corrupted Python installation
        echo.
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
    echo Try running: setup.bat --recreate
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

if !ERRORLEVEL! neq 0 (
    echo.
    echo ERROR: Failed to install dependencies
    echo.
    echo Possible causes:
    echo   - No internet connection
    echo   - Package version conflicts
    echo   - Corrupted requirements.txt
    echo.
    echo Check the error messages above for details
    pause
    exit /b 1
)
echo Dependencies installed successfully
echo.

REM ============================================================================
REM Step 6: Environment Configuration
REM ============================================================================
echo [6/6] Configuring environment...

REM Check for .env
if not exist ".env" (
    if exist ".env.example" (
        echo Creating .env from template...
        copy .env.example .env >nul
        echo.
        echo IMPORTANT: Please edit .env and configure required settings
        echo.
    ) else (
        echo WARNING: No .env.example found, skipping .env creation
    )
) else (
    echo .env already exists
)

REM Check for tool-specific .env files
for /d %%D in (tools\*) do (
    if exist "%%D\.env.example" (
        if not exist "%%D\.env" (
            echo Creating %%D\.env from template...
            copy "%%D\.env.example" "%%D\.env" >nul
        )
    )
)
echo.

REM ============================================================================
REM Validation
REM ============================================================================
echo Validating installation...

python -c "import sys; from core import config, logger; from office.outlook import client; print('✓ Core modules loaded')" 2>nul
if !ERRORLEVEL! equ 0 (
    echo ✓ Installation validated successfully
) else (
    echo ✗ WARNING: Module validation failed
    echo   Some features may not work correctly
)
echo.

REM Show installed packages
echo Installed packages:
pip list --format=freeze | findstr /v "^#" | findstr /v "^$"
echo.

call deactivate

REM ============================================================================
REM Summary
REM ============================================================================
echo ============================================================================
echo  Setup Complete!
echo ============================================================================
echo.
echo Next steps:
echo   1. Review and edit .env files with your configuration
echo   2. For Outlook tools: Make sure Outlook Desktop is running
echo   3. Run a tool:
echo      - Double-click: tools\[tool-name]\run.bat
echo      - Command line: run.bat [tool-name]
echo      - Interactive: run.bat (then select from menu)
echo.
echo For help and documentation, see README.md
echo.
pause
```

---

### Master Launcher (run.bat)

```bat
@echo off
REM ============================================================================
REM Daisy Automation Platform - Master Launcher
REM ============================================================================
REM Purpose: Central entry point for running all tools
REM Usage:
REM   run.bat                       - Show interactive menu
REM   run.bat [tool-name] [args]    - Run specific tool directly
REM Examples:
REM   run.bat payslip-phuclong
REM   run.bat payslip-phuclong --dry-run
REM ============================================================================

setlocal EnableDelayedExpansion

REM ============================================================================
REM Check Virtual Environment
REM ============================================================================
if not exist "venv\Scripts\activate.bat" (
    echo.
    echo ============================================================================
    echo  ERROR: Virtual Environment Not Found
    echo ============================================================================
    echo.
    echo The virtual environment has not been set up yet.
    echo.
    echo Please run setup.bat first:
    echo   1. Double-click setup.bat in this folder
    echo   2. Wait for setup to complete
    echo   3. Try running this tool again
    echo.
    echo If you have already run setup.bat, the virtual environment may be corrupted.
    echo Try running: setup.bat --recreate
    echo.
    pause
    exit /b 1
)

REM ============================================================================
REM Activate Virtual Environment
REM ============================================================================
call venv\Scripts\activate.bat

REM ============================================================================
REM Discover Available Tools
REM ============================================================================
set TOOL_COUNT=0
set TOOL_LIST=

for /d %%D in (tools\*) do (
    if exist "%%D\main.py" (
        set /a TOOL_COUNT+=1
        set TOOL_!TOOL_COUNT!=%%~nxD
        set TOOL_LIST=!TOOL_LIST! %%~nxD
    )
)

if !TOOL_COUNT! equ 0 (
    echo ERROR: No tools found in tools\ directory
    call deactivate
    pause
    exit /b 1
)

REM ============================================================================
REM Parse Arguments
REM ============================================================================
set TOOL_NAME=%~1
set TOOL_FOUND=0

REM If no arguments, show interactive menu
if "!TOOL_NAME!"=="" (
    goto show_menu
)

REM Check if tool exists
for /d %%D in (tools\*) do (
    if /i "%%~nxD"=="!TOOL_NAME!" (
        if exist "%%D\main.py" (
            set TOOL_FOUND=1
            set TOOL_PATH=%%D
        )
    )
)

if !TOOL_FOUND! equ 0 (
    echo.
    echo ERROR: Tool "!TOOL_NAME!" not found
    echo.
    echo Available tools:
    for /d %%D in (tools\*) do (
        if exist "%%D\main.py" echo   - %%~nxD
    )
    echo.
    call deactivate
    pause
    exit /b 1
)

goto run_tool

REM ============================================================================
REM Interactive Menu
REM ============================================================================
:show_menu
cls
echo.
echo ============================================================================
echo  Daisy Automation Platform - Tool Launcher
echo ============================================================================
echo.
echo Available tools:
echo.

set INDEX=1
for /l %%i in (1,1,!TOOL_COUNT!) do (
    echo   %%i^) !TOOL_%%i!
    set INDEX=%%i
)

echo.
echo   Q^) Quit
echo.
echo ============================================================================
set /p CHOICE="Select a tool (1-!TOOL_COUNT! or Q): "

if /i "!CHOICE!"=="Q" (
    call deactivate
    exit /b 0
)

REM Validate choice
set VALID=0
for /l %%i in (1,1,!TOOL_COUNT!) do (
    if "!CHOICE!"=="%%i" (
        set TOOL_NAME=!TOOL_%%i!
        set TOOL_PATH=tools\!TOOL_%%i!
        set VALID=1
    )
)

if !VALID! equ 0 (
    echo Invalid choice. Please try again.
    timeout /t 2 >nul
    goto show_menu
)

REM ============================================================================
REM Run Tool
REM ============================================================================
:run_tool
echo.
echo ============================================================================
echo  Running: !TOOL_NAME!
echo ============================================================================
echo.

REM Shift to get remaining arguments
shift

REM Build argument list
set ARGS=
:build_args
if not "%~1"=="" (
    set ARGS=!ARGS! %1
    shift
    goto build_args
)

REM Check for tool-specific .env
if exist "!TOOL_PATH!\.env" (
    echo Loading tool configuration...
)

REM Execute tool
python "!TOOL_PATH!\main.py" !ARGS!
set EXIT_CODE=!ERRORLEVEL!

echo.
if !EXIT_CODE! equ 0 (
    echo ============================================================================
    echo  Tool completed successfully
    echo ============================================================================
) else (
    echo ============================================================================
    echo  Tool exited with error code: !EXIT_CODE!
    echo ============================================================================
    echo.
    echo If you need help, please:
    echo   1. Take a screenshot of the error above
    echo   2. Check the tool's README.md for troubleshooting
    echo   3. Contact IT support with the screenshot
)

echo.
call deactivate

if exist "!TOOL_PATH!\run.bat" (
    REM Called from wrapper, pause for user
    pause
)

exit /b !EXIT_CODE!
```

---

### Per-Tool Wrapper Template (tools/[tool-name]/run.bat)

```bat
@echo off
REM ============================================================================
REM [Tool Name] - Quick Launcher
REM ============================================================================
REM This is a convenience wrapper for double-click execution
REM It calls the master launcher with the tool name pre-filled
REM ============================================================================

setlocal

REM Get the directory of this script
set SCRIPT_DIR=%~dp0

REM Navigate to repository root (2 levels up from tools/tool-name/)
cd /d "%SCRIPT_DIR%..\.."

REM Call master launcher with tool name
call run.bat payslip-phuclong %*

REM Preserve exit code
exit /b %ERRORLEVEL%
```

---

### .gitignore Updates

```gitignore
# Virtual Environment
venv/
env/
ENV/

# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python

# Environment Variables
.env
.env.local

# IDE
.vscode/
.idea/
*.swp
*.swo
*~

# OS
.DS_Store
Thumbs.db
desktop.ini

# Logs
*.log
logs/

# Testing
.pytest_cache/
.coverage
htmlcov/

# Tool-specific outputs
tools/*/output/
tools/*/temp/
*.pdf
*.xlsx
```

---

## Testing and Validation

### Test Suite Overview

Create a test script `tests/test_venv_setup.bat` to validate the infrastructure:

```bat
@echo off
REM Automated test suite for venv setup and launcher

setlocal EnableDelayedExpansion

echo ============================================================================
echo  Daisy Platform - Automated Tests
echo ============================================================================
echo.

set PASS=0
set FAIL=0

REM Test 1: Check Python installation
echo [TEST 1] Checking Python installation...
where python >nul 2>nul
if %ERRORLEVEL% equ 0 (
    echo   ✓ PASS: Python found in PATH
    set /a PASS+=1
) else (
    echo   ✗ FAIL: Python not found
    set /a FAIL+=1
)

REM Test 2: Check venv exists
echo [TEST 2] Checking virtual environment...
if exist "venv\Scripts\python.exe" (
    echo   ✓ PASS: Virtual environment exists
    set /a PASS+=1
) else (
    echo   ✗ FAIL: Virtual environment not found
    set /a FAIL+=1
)

REM Test 3: Activate venv and check packages
echo [TEST 3] Checking installed packages...
call venv\Scripts\activate.bat
pip show openpyxl >nul 2>nul
if %ERRORLEVEL% equ 0 (
    echo   ✓ PASS: Dependencies installed
    set /a PASS+=1
) else (
    echo   ✗ FAIL: Dependencies missing
    set /a FAIL+=1
)

REM Test 4: Import core modules
echo [TEST 4] Testing module imports...
python -c "from core import config; from office.outlook import client" 2>nul
if %ERRORLEVEL% equ 0 (
    echo   ✓ PASS: Core modules import successfully
    set /a PASS+=1
) else (
    echo   ✗ FAIL: Module import failed
    set /a FAIL+=1
)

REM Test 5: Check run.bat exists
echo [TEST 5] Checking launcher scripts...
if exist "run.bat" (
    echo   ✓ PASS: Master launcher exists
    set /a PASS+=1
) else (
    echo   ✗ FAIL: Master launcher missing
    set /a FAIL+=1
)

call deactivate

echo.
echo ============================================================================
echo  Test Results
echo ============================================================================
echo   Passed: !PASS!
echo   Failed: !FAIL!
echo.

if !FAIL! equ 0 (
    echo ✓ All tests passed!
    exit /b 0
) else (
    echo ✗ Some tests failed. Please run setup.bat
    exit /b 1
)
```

### Manual Test Checklist

- [ ] **Fresh Setup Test**
  1. Delete `venv/` folder
  2. Run `setup.bat`
  3. Verify no errors
  4. Check `venv/Scripts/python.exe` exists
  
- [ ] **Launcher Test - Interactive**
  1. Double-click `run.bat`
  2. Verify menu appears with tool list
  3. Select a tool by number
  4. Verify tool runs
  
- [ ] **Launcher Test - Direct**
  1. Open Command Prompt
  2. Navigate to repository
  3. Run `run.bat payslip-phuclong`
  4. Verify tool runs
  
- [ ] **Wrapper Test**
  1. Navigate to `tools/payslip-phuclong/`
  2. Double-click `run.bat`
  3. Verify tool runs via master launcher
  
- [ ] **Error Handling Test**
  1. Rename `venv/` to `venv_backup/`
  2. Double-click any tool wrapper
  3. Verify clear error message appears
  4. Verify instructions to run `setup.bat`
  
- [ ] **Recreate Test**
  1. Corrupt venv by deleting `venv/Scripts/python.exe`
  2. Run `setup.bat --recreate`
  3. Verify venv is fully rebuilt
  4. Test tools run correctly
  
- [ ] **Upgrade Test**
  1. Update a package version in `requirements.txt`
  2. Run `setup.bat --upgrade`
  3. Verify package upgrades
  4. Test tools still work
  
- [ ] **Multi-User Test**
  1. Copy repository to different user profile
  2. Run `setup.bat` as different user
  3. Verify tools work independently
  
- [ ] **Network Drive Test**
  1. Place repository on network drive
  2. Run `setup.bat`
  3. Test tool execution speed
  4. Verify no UNC path issues

---

## Troubleshooting Guide

### Common Issues and Solutions

#### Issue 1: "Python not found in PATH"

**Symptoms:**
```
ERROR: Python not found in PATH
```

**Solutions:**
1. Install Python 3.14 from [python.org](https://www.python.org/)
2. During installation, check "Add Python to PATH"
3. If already installed:
   - Find Python installation directory
   - Add to PATH manually:
     - Right-click "This PC" → Properties
     - Advanced system settings → Environment Variables
     - Edit PATH, add Python directory
   - Restart Command Prompt

**Verification:**
```bat
python --version
```

---

#### Issue 2: "Virtual environment not found"

**Symptoms:**
```
ERROR: Virtual Environment Not Found
Please run setup.bat first
```

**Solutions:**
1. Run `setup.bat` from repository root
2. If setup.bat was run but error persists:
   - Check if `venv/` folder exists
   - Run `setup.bat --recreate`
3. If on network drive:
   - Check write permissions
   - Try copying to local drive first

**Verification:**
```bat
dir venv\Scripts\python.exe
```

---

#### Issue 3: "Failed to install dependencies"

**Symptoms:**
```
ERROR: Failed to install dependencies
Could not find a version that satisfies the requirement...
```

**Solutions:**
1. Check internet connection
2. Check corporate proxy settings:
   ```bat
   set HTTP_PROXY=http://proxy.company.com:8080
   set HTTPS_PROXY=http://proxy.company.com:8080
   pip install -r requirements.txt
   ```
3. Use offline wheel cache:
   - Download packages on connected machine
   - Transfer wheel files
   - Install from local directory
4. Check `requirements.txt` for typos

**Verification:**
```bat
venv\Scripts\pip list
```

---

#### Issue 4: "Module import failed"

**Symptoms:**
```
ModuleNotFoundError: No module named 'core'
```

**Solutions:**
1. Ensure working directory is repository root
2. Check `PYTHONPATH`:
   ```bat
   set PYTHONPATH=%CD%
   ```
3. Reinstall dependencies:
   ```bat
   setup.bat --recreate
   ```
4. Check for circular imports in code

**Verification:**
```bat
venv\Scripts\python -c "from core import config; print('OK')"
```

---

#### Issue 5: Tool-specific configuration errors

**Symptoms:**
```
ERROR: OUTLOOK_ACCOUNT not configured
```

**Solutions:**
1. Check if `.env` exists in tool folder
2. Copy from template:
   ```bat
   copy tools\payslip-phuclong\.env.example tools\payslip-phuclong\.env
   ```
3. Edit `.env` and set required values
4. Ensure no typos in variable names

**Verification:**
```bat
type tools\payslip-phuclong\.env
```

---

#### Issue 6: Slow execution on network drive

**Symptoms:**
- Setup takes 10+ minutes
- Tool execution is sluggish

**Solutions:**
1. **Preferred:** Move repository to local drive (C:, D:)
2. **Workaround:** Use central venv on local drive:
   - Create `C:\daisy-venv\`
   - Modify scripts to point to central location
3. **Corporate IT:** Request exemption for local development

---

#### Issue 7: Antivirus blocking venv creation

**Symptoms:**
```
ERROR: Failed to create virtual environment
Access denied
```

**Solutions:**
1. Temporarily disable antivirus (if policy allows)
2. Add repository folder to antivirus exceptions
3. Contact IT to whitelist Python development
4. Use corporate-approved Python distribution

---

### Advanced Troubleshooting

#### Enable Debug Logging

Modify `run.bat` to add debug output:

```bat
@echo on  REM Enable command echo
set DEBUG=1
set PYTHONVERBOSE=1
```

#### Check Virtual Environment Integrity

```bat
venv\Scripts\python -m venv --help
venv\Scripts\pip check
```

#### Test Individual Components

```bat
REM Test Python
venv\Scripts\python --version

REM Test pip
venv\Scripts\pip --version

REM Test module imports
venv\Scripts\python -c "import sys; print(sys.path)"

REM Test specific import
venv\Scripts\python -c "from core import config; print(config.__file__)"
```

---

## Migration and Rollout Strategy

### Phase 1: Development and Testing (Week 1)

**Objective:** Implement and validate infrastructure

**Tasks:**
1. Update `setup.bat` with enhancements
2. Create `run.bat` master launcher
3. Create tool wrappers
4. Update `.gitignore`
5. Create test suite
6. Test on developer machines

**Deliverables:**
- Working setup and launcher scripts
- Passing test suite
- Initial documentation

**Success Criteria:**
- 3+ developers successfully use new setup
- All existing tools work without modification
- Zero regression in tool functionality

---

### Phase 2: IT Pilot (Week 2)

**Objective:** Validate with IT support staff

**Tasks:**
1. Deploy to IT team's test environment
2. Conduct training session
3. Gather feedback on error messages
4. Iterate on user experience
5. Create internal support documentation

**Deliverables:**
- IT training materials
- Support playbook
- Updated error messages

**Success Criteria:**
- IT can deploy and troubleshoot independently
- Clear resolution paths for common issues
- Positive feedback from IT team

---

### Phase 3: Limited User Beta (Week 3)

**Objective:** Test with 5-10 non-technical users

**Tasks:**
1. Select representative users (HR, managers)
2. Deploy to user machines
3. Observe first-time setup experience
4. Collect feedback via survey
5. Track support tickets

**Deliverables:**
- User feedback report
- List of common issues
- Quick start video/guide

**Success Criteria:**
- 90%+ users complete setup without assistance
- Average setup time < 10 minutes
- < 2 support tickets per user

---

### Phase 4: Full Rollout (Week 4-5)

**Objective:** Deploy to all users

**Tasks:**
1. Mass deployment to all users
2. Distribute quick start guide
3. Hold training sessions
4. Monitor support volume
5. Iterate on documentation

**Deliverables:**
- Organization-wide deployment
- Comprehensive documentation
- Training recordings
- Support FAQ

**Success Criteria:**
- All users operational
- Support ticket volume manageable
- Positive user sentiment

---

### Rollback Plan

If critical issues arise:

1. **Immediate:** Restore previous setup method (if any)
2. **Identify:** Root cause of failures
3. **Fix:** Address in development environment
4. **Re-test:** Complete Phase 1-3 again
5. **Re-deploy:** Once validation passes

**Rollback Triggers:**
- > 50% users cannot complete setup
- Critical tool functionality broken
- Data loss or corruption
- Unresolvable dependency conflicts

---

## Maintenance and Best Practices

### Monthly Maintenance Tasks

1. **Review Dependencies**
   - Check for security updates: `pip list --outdated`
   - Test updates in development before deploying
   - Update `requirements.txt` with new versions

2. **Validate Setup**
   - Run test suite
   - Test fresh setup on clean machine
   - Verify all tools run correctly

3. **Documentation Review**
   - Update troubleshooting with new issues
   - Refine error messages based on support tickets
   - Update screenshots and examples

---

### Quarterly Review

1. **Evaluate Tool Growth**
   - If > 10 tools: Consider per-tool venvs for some
   - If dependency conflicts emerge: Migrate affected tools
   - If disk usage becomes issue: Explore alternatives

2. **User Feedback**
   - Survey users on experience
   - Identify pain points
   - Prioritize improvements

3. **Technology Updates**
   - Evaluate new Python version
   - Consider packaging tools as executables
   - Explore alternative distribution methods

---

###Best Practices for Developers

#### Adding a New Tool

1. Create tool directory:
   ```bat
   mkdir tools\new-tool
   ```

2. Copy template files:
   ```bat
   copy tools\_template\* tools\new-tool\
   ```

3. Implement `main.py` with argparse:
   ```python
   import argparse
   from core import Config, Logger
   
   def main():
       parser = argparse.ArgumentParser(description="New Tool")
       parser.add_argument("--dry-run", action="store_true")
       args = parser.parse_args()
       
       # Tool logic here
       
   if __name__ == "__main__":
       main()
   ```

4. Create tool wrapper:
   ```bat
   REM tools\new-tool\run.bat
   @echo off
   call ..\..\run.bat new-tool %*
   ```

5. Test:
   ```bat
   run.bat new-tool
   ```

6. Document in `tools\new-tool\README.md`

---

#### Adding a New Dependency

1. Test in activated venv:
   ```bat
   venv\Scripts\activate
   pip install new-package
   python -m pytest  # Validate no conflicts
   ```

2. Add to `requirements.txt`:
   ```
   new-package>=1.2.3  # Comment: Why needed
   ```

3. Test fresh install:
   ```bat
   setup.bat --recreate
   # Test all tools
   ```

4. Commit `requirements.txt` changes

---

####Handling Dependency Conflicts

If a conflict arises:

1. **Analysis:**
   - Identify which tools need conflicting versions
   - Check if upgrading resolves conflict
   - Evaluate impact of version changes

2. **Resolution Options:**
   - **Option A:** Find compatible versions
     ```
     tool-a requires pandas>=1.5,<2.0
     tool-b requires pandas>=1.3
     → Use pandas~=1.5.0
     ```
   
   - **Option B:** Refactor code to support newer version
   
   - **Option C:** Migrate conflicting tool to per-tool venv
     ```bat
     # Create tools\tool-b\setup.bat
     # Create tools\tool-b\requirements.txt
     # Update tools\tool-b\run.bat to use local venv
     ```

3. **Testing:**
   - Test all tools with proposed resolution
   - Run full regression suite
   - Validate on multiple machines

---

### Security Best Practices

1. **Pin Dependencies:**
   ```
   # Bad
   pandas

   # Good
   pandas>=2.0.0,<3.0.0

   # Best (for stability)
   pandas==2.0.3
   ```

2. **Regular Security Audits:**
   ```bat
   pip install safety
   safety check -r requirements.txt
   ```

3. **Keep Python Updated:**
   - Monitor Python security advisories
   - Test new Python versions quarterly
   - Update recommendation in `setup.bat`

4. **Protect Secrets:**
   - Never commit `.env` files
   - Use `.env.example` templates
   - Document required secrets clearly

---

## Future Considerations

### Short-term (3-6 months)

1. **Enhanced Logging:**
   - Add centralized logging for all tools
   - Create log viewer utility
   - Implement log retention policy

2. **Config Management:**
   - Create GUI for `.env` editing
   - Validate configuration on tool start
   - Support multiple configuration profiles

3. **Update Automation:**
   - Auto-check for repository updates
   - Self-updating launcher scripts
   - Notification system for new releases

---

### Medium-term (6-12 months)

1. **Tool Packaging:**
   - Evaluate PyInstaller for standalone executables
   - Create portable tool bundles
   - Implement auto-updater for executables

2. **Web Interface:**
   - Create simple web UI for tool execution
   - Allow remote execution from browser
   - Centralized job scheduling

3. **Monitoring and Analytics:**
   - Track tool usage statistics
   - Monitor error rates
   - Performance profiling

---

### Long-term (12+ months)

1. **Cloud Integration:**
   - Migrate tools to serverless (AWS Lambda, Azure Functions)
   - Centralized execution platform
   - Web-based administration

2. **Enterprise Features:**
   - Role-based access control
   - Audit logging
   - Compliance reporting

3. **Tool Marketplace:**
   - Internal tool registry
   - Dependency resolution service
   - Automated deployment pipeline

---

## Decision Matrix

### Scoring Criteria

| Criteria | Weight | Single Venv | Per-Tool Venv | Central Venv | pipx/exe |
|----------|--------|-------------|---------------|--------------|----------|
| **User Simplicity** | 25% | 10 | 5 | 7 | 9 |
| **Developer Maintenance** | 20% | 10 | 4 | 8 | 6 |
| **Disk Efficiency** | 10% | 10 | 3 | 10 | 5 |
| **Isolation/Safety** | 15% | 6 | 10 | 6 | 10 |
| **Setup Speed** | 10% | 9 | 5 | 9 | 7 |
| **Troubleshooting Ease** | 10% | 8 | 6 | 5 | 9 |
| **Future Flexibility** | 10% | 7 | 10 | 7 | 8 |

### Weighted Scores

| Approach | Total Score | Rank |
|----------|-------------|------|
| **Single Venv** | **8.65** | **1** |
| Per-Tool Venv | 6.40 | 3 |
| Central Venv | 7.35 | 2 |
| pipx/exe | 7.85 | 2 |

**Conclusion:** Single shared venv scores highest for current requirements, with pipx/exe as future migration target.

---

## Appendices

### Appendix A: Complete Tool Template

**Directory Structure:**
```
tools/_template/
├── main.py
├── config.py
├── run.bat
├── README.md
├── .env.example
└── tests/
    └── test_main.py
```

**tools/_template/main.py:**
```python
#!/usr/bin/env python3
"""
Template Tool - Description Here
"""

import argparse
import sys
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from core import Config, Logger
from core.state import StateManager

def main():
    """Main entry point for the tool."""
    parser = argparse.ArgumentParser(
        description="Template Tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py --verbose
  python main.py --config custom.env --dry-run
        """
    )
    
    parser.add_argument(
        "--config",
        type=str,
        default=".env",
        help="Configuration file path (default: .env)"
    )
    
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Run without making changes"
    )
    
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose logging"
    )
    
    args = parser.parse_args()
    
    # Initialize logger
    logger = Logger.get_logger(__name__)
    if args.verbose:
        Logger.set_level("DEBUG")
    
    logger.info("Starting Template Tool")
    
    try:
        # Load configuration
        config = Config.from_file(args.config)
        
        # Initialize state manager
        state = StateManager("template-tool")
        
        # Tool logic here
        if args.dry_run:
            logger.info("Dry-run mode: No changes will be made")
        
        # Example processing
        logger.info("Processing...")
        result = perform_task(config, dry_run=args.dry_run)
        
        # Save state
        state.save({"last_run": "success", "result": result})
        
        logger.info("Tool completed successfully")
        return 0
        
    except Exception as e:
        logger.error(f"Tool failed: {e}", exc_info=True)
        return 1


def perform_task(config, dry_run=False):
    """
    Perform the main task of this tool.
    
    Args:
        config: Configuration object
        dry_run: If True, don't make actual changes
        
    Returns:
        Result of the operation
    """
    # Implementation here
    return {"status": "success"}


if __name__ == "__main__":
    sys.exit(main())
```

---

### Appendix B: Environment Variable Reference

**Root `.env`:**
```bash
# Global configuration for all tools
OUTLOOK_ACCOUNT=user@company.com
LOG_LEVEL=INFO
TEMP_DIR=temp/
```

**Per-Tool `.env.example`:**
```bash
# Payslip Tool Configuration

# Required: Excel file path
EXCEL_FILE_PATH=data/employee_data.xlsx

# Required: Output directory
OUTPUT_DIR=output/payslips/

# Optional: Email settings
SEND_EMAIL=true
EMAIL_SUBJECT=Your Payslip for {month}

# Optional: PDF settings
PDF_PASSWORD_PROTECT=true
PDF_OWNER_PASSWORD=admin123
```

---

### Appendix C: Support Escalation Matrix

| Issue Type | First Contact | Escalation | SLA |
|------------|---------------|------------|-----|
| Setup problems | IT Support | Developer | 4 hours |
| Tool crashes | IT Support | Developer | 8 hours |
| Wrong results | Tool Owner | Developer | 24 hours |
| Feature requests | Tool Owner | Product Owner | N/A |
| Security issues | IT Security | Developer | Immediate |

---

### Appendix D: Glossary

- **venv:** Python virtual environment, isolated Python installation
- **pip:** Python package installer
- **wrapper script:** Simple .bat file that calls another script
- **master launcher:** Central script that coordinates tool execution
- **non-technical user:** User without programming knowledge
- **activation:** Process of enabling a virtual environment
- **requirement:** Python package dependency
- **conflict:** When two packages need incompatible versions of a dependency

---

### Appendix E: Quick Reference Commands

```bat
REM Setup
setup.bat                    # Initial setup
setup.bat --recreate         # Rebuild venv from scratch
setup.bat --upgrade          # Upgrade all packages

REM Running Tools
run.bat                      # Interactive menu
run.bat tool-name            # Run specific tool
run.bat tool-name --help     # Show tool help

REM Troubleshooting
venv\Scripts\python --version           # Check Python version
venv\Scripts\pip list                   # List installed packages
venv\Scripts\pip check                  # Check for conflicts
tests\test_venv_setup.bat               # Run test suite

REM Development
venv\Scripts\activate.bat               # Activate venv manually
pip install package-name                # Install new package
pip freeze > requirements.txt           # Update requirements
deactivate                              # Deactivate venv
```

---

## Document Change History

| Version | Date | Author | Changes |
|---------|------|--------|---------|
| 1.0 | 2026-02-08 | Platform Team | Initial comprehensive strategy document |

---

**End of Document**
