# Daisy Automation Platform

**Office automation tools for Windows - making repetitive tasks easier.**

This platform helps automate office tasks like generating documents, sending emails, and processing data. Everything runs on your computer using familiar tools like Excel and Outlook.

---

## Contact point

If you have any issue or any idea to improve this plaform, please contact:
**Trong H. Truong (trongtruong2509@gmail.com)**

**Last Updated:** 09-Feb-2026

---

## What's Included

### Available Tools (On-going Development)

- **Payslip Generator** (`payslip-phuclong`) - Automatically generates and emails monthly payslips to employees
  - See [tools/payslip-phuclong-ecom/README.md](tools/payslip-phuclong-ecom/README.md) for detailed instructions

---

## What You Need

- **Computer**: Windows 10 or Windows 11
- **Software**:
  - Python 3.9 or newer ([Instruction here from MobileMingle](https://www.youtube.com/watch?v=9eYiKMCWSXQ))
  - Git for Windows: https://git-scm.com/install/windows
  - Microsoft Outlook (desktop app, not web version)
  - Microsoft Excel (for payslip tool)
- **Permissions**: Regular user account (no admin rights needed)

---

## Getting Started

### Step 1: Download the Project

If you have Git installed:

```bash
git clone <repository-url>
cd daisy
```

Or download as ZIP file and extract it to a folder.

### Step 2: Run Setup

Double-click `setup.bat` or open Command Prompt in the project folder and run:

```cmd
setup.bat
```

**First-time setup takes 2-3 minutes.** You only need to do this once.

**You'll see:**

```
Setting up Daisy Automation Platform...
Creating virtual environment...
Installing dependencies...
Setup complete!
```

### Step 3: Configure Your Tool

Each tool has its own settings. For the payslip tool, see [tools/payslip-phuclong-ecom/README.md](tools/payslip-phuclong-ecom/README.md) for configuration instructions.

### Step 4: Run a Tool

**Option 1: Interactive Menu** (Easiest)

Double-click `run.bat` or run:

```cmd
run.bat
```

You'll see a menu like this:

```
====================================================================
 Daisy Automation Platform - Main Menu
====================================================================

Available Tools:

  [1] payslip-phuclong

  [0] Exit

====================================================================

Select an option (0-1):
```

Just type `1` and press Enter to run the payslip tool.

**Option 2: Direct Command** (Faster if you know the tool name)

```cmd
run.bat payslip-phuclong
```

**That's it!** The tool will guide you through the rest.

---

## Using the Master Script

The `run.bat` file is your main entry point for all tools.

### How to Use

**Show interactive menu:**

```cmd
run.bat
```

**Run a specific tool directly:**

```cmd
run.bat payslip-phuclong
```

**Get help:**

```cmd
run.bat --help
```

### What Happens When You Run It

1. The script activates the Python environment
2. Shows you available tools (if no tool specified)
3. Runs the tool you selected
4. Closes Python environment when done

You don't need to understand what happens behind the scenes - just run the command and follow the prompts!

---

## Configuration

### Tool-Specific Configuration

Each tool has its own `.env` file with settings specific to that tool.

For the payslip tool, see [tools/payslip-phuclong-ecom/README.md](tools/payslip-phuclong-ecom/README.md) for configuration options.

---

## Understanding Test Mode (Dry Run)

**Test mode (`DRY_RUN=true`) is very important!**

When test mode is ON:

- ✓ The tool runs normally
- ✓ Shows you what it would do
- ✓ Creates log files
- ✗ **Does NOT send real emails**
- ✗ **Does NOT modify files**

**Always test with `DRY_RUN=true` first** to make sure everything works correctly!

When you're ready to run for real:

1. Open the `.env` file in the tool's folder
2. Change `DRY_RUN=true` to `DRY_RUN=false`
3. Save the file
4. Run the tool again

---

## Troubleshooting

### "Python is not recognized" error

**Problem:** Python is not installed or not in PATH.

**Solution:**

1. Download Python from https://www.python.org/downloads/
2. During installation, **check the box "Add Python to PATH"**
3. Restart your computer
4. Run `setup.bat` again

### "setup.bat failed" error

**Problem:** Something went wrong during setup.

**Solution:**

1. Close all command prompt windows
2. Delete the `venv` folder in the project directory
3. Run `setup.bat` again
4. If it still fails, contact IT support

### "Outlook not found" error

**Problem:** Outlook desktop app is not running or not installed.

**Solution:**

1. Make sure Outlook **desktop app** is installed (not just web version)
2. Open Outlook before running the tool
3. Make sure you're logged into your account in Outlook

### Tool runs but doesn't do anything

**Problem:** Probably in test mode (dry run).

**Solution:**

1. Check if `DRY_RUN=true` in the tool's `.env` file
2. Change to `DRY_RUN=false` when ready to run for real
3. Check the log files in `logs/` folder for details

### "Permission denied" errors

**Problem:** The tool can't write files to the folder.

**Solution:**

1. Make sure you have write permissions to the project folder
2. Don't run the tool from a network drive or read-only location
3. Try moving the project to your Documents folder

### Tool stops in the middle

**Problem:** Something went wrong during processing.

**Solution:**

1. Check the log file in `logs/` folder for error messages
2. Run the tool again - it can usually resume where it stopped
3. If the same error happens again, note the error message and contact IT support

---

## Where to Find Things

After running tools, you'll find:

| What              | Where                        | Description                         |
| ----------------- | ---------------------------- | ----------------------------------- |
| **Log files**     | `logs/`                      | Detailed record of what happened    |
| **Output files**  | `output/`                    | Results, PDFs, CSV reports          |
| **State files**   | `state/` or `tools/*/state/` | Tracks progress (so you can resume) |
| **Configuration** | `.env` files                 | Settings for each tool              |

---
