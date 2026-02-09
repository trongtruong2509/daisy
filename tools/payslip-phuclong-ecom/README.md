# Payslip Generator & Email Sender

**Automated monthly payslip generation and email distribution for Phuc Long employees.**

This tool automatically:

1. Reads employee data from Excel
2. Generates individual payslip files
3. Creates password-protected PDFs
4. Sends personalized emails via Outlook with payslip attachments

---

## Quick Start

### First Time Setup

1. **Prepare the Excel File**
   - Update the payslip Excel file with current month's data
   - Make sure employee information is complete (ID, Name, Email)

2. **Configure Settings** (Optional)
   - Copy `.env.example` to `.env`
   - Edit values if needed (see Configuration section below)

3. **Run the Tool**

   ```cmd
   run.bat
   ```

   Or from the project root:

   ```cmd
   run.bat payslip-phuclong
   ```

4. **Follow the Prompts**
   - The tool will ask for Excel file path if not configured
   - Enter the payroll month (MM/YYYY format)
   - Select your Outlook account from the list
   - Review the summary and confirm

---

## Configuration

You can set these values in `.env` file, or the tool will prompt you when needed:

| Setting              | What It Does                 | Example                              |
| -------------------- | ---------------------------- | ------------------------------------ |
| `PAYSLIP_EXCEL_PATH` | Path to Excel file           | `D:\repos\path\to\TBKQ-phuclong.xls` |
| `DATE`               | Payroll month                | `01/2026`                            |
| `OUTLOOK_ACCOUNT`    | Your email address           | `hr@phuclong.com`                    |
| `DRY_RUN`            | Test mode (no emails sent)   | `true` or `false`                    |
| `KEEP_PDF_PAYSLIPS`  | Keep PDF files after sending | `true` or `false`                    |

**Tip:** Always test with `DRY_RUN=true` first to make sure everything works correctly!

---

## How to Use

### Step 1: Start the Tool

From the project root directory:

```cmd
run.bat payslip-phuclong
```

Or double-click `run.bat` inside the `tools/payslip-phuclong-ecom` folder.

### Step 2: Provide Required Information

If you haven't set up `.env`, the tool will ask you:

**Excel File Path:**

```
Enter the path to your Excel file (e.g., D:\repos\trongtruong2509\daisy\excel-files\TBKQ-phuclong.xls):
```

**Payroll Date:**

```
Enter payroll date in MM/YYYY format (e.g., 01/2026):
```

**Outlook Account:**

```
Choose your Outlook account:
  [1] hr@phuclong.com
  [2] accounting@phuclong.com
Select: 1
```

### Step 3: Review and Confirm

The tool shows a summary:

```
Configuration Summary
---------------------
Excel file      : TBKQ-phuclong.xls
Payroll date    : 01/2026
Employees       : 250
Outlook account : hr@phuclong.com
Dry run         : No
```

Type `yes` to proceed or `no` to cancel.

### Step 4: Wait for Completion

The tool will:

- ✓ Generate 250 payslips
- ✓ Convert to PDF (with password protection)
- ✓ Send 250 emails

You'll see progress updates as it works.

### Step 5: Check Results

After completion:

- Log file: `logs/payslip_YYYYMMDD_HHMMSS.log`
- Result file: `output/MMYYYY/sent_results_MMYYYY.csv`
- PDF files: `output/MMYYYY/` (if KEEP_PDF_PAYSLIPS=true)

---

## Important Notes

### Before Running for Real

1. **Test with Dry Run First**
   - Set `DRY_RUN=true` in `.env`
   - Run the tool to verify everything works
   - No emails will actually be sent

2. **Check Employee Data**
   - Make sure all employees have valid email addresses
   - Verify employee IDs and names are correct
   - The tool will validate data before sending

3. **Outlook Must Be Running**
   - Open Outlook Desktop before running the tool
   - Make sure you're logged into the correct account

### During Execution

- **Don't close Outlook** - The tool uses it to send emails
- **Don't close Excel** - The tool may open Excel in the background
- **Be patient** - Processing 2000+ employees takes 60-90 minutes
- **Check progress** - The tool shows updates every 25 employees

### If Something Goes Wrong

**The tool stops or crashes:**

- Don't worry! Your progress is saved
- Run the tool again - it will ask if you want to resume
- Choose "yes" to continue from where it stopped

**"Employee already processed" messages:**

- This is normal if you're resuming after a crash
- The tool prevents sending duplicate emails

**Validation errors:**

- The tool will list all problems before starting
- Fix the Excel file and run again

---

## Common Questions

**Q: Can I test without sending real emails?**  
A: Yes! Set `DRY_RUN=true` in `.env` and run the tool. It will simulate everything without sending emails.

**Q: What if I need to stop in the middle?**  
A: Just close the window. Next time you run, the tool will offer to resume from where you stopped.

**Q: How do I know which employees got their emails?**  
A: Check the CSV file in `output/MMYYYY/sent_results_MMYYYY.csv` - it lists every employee and their status.

**Q: Can I run this for multiple months?**  
A: Yes! Change the `DATE` in `.env` or enter a different date when prompted. Each month creates its own output folder.

**Q: What's the PDF password?**  
A: Each employee's PDF is protected with their employee ID number.

**Q: Will employees get duplicate emails if I run twice?**  
A: No. The tool tracks who already received their payslip and skips them automatically.

---

## Troubleshooting

### "Excel file not found"

- Check the path in `.env` or the path you entered
- Use relative path: `../../excel-files/file.xls`
- Or full path: `D:/Documents/excel-files/file.xls`

### "Outlook account not found"

- Make sure Outlook Desktop is running (not Outlook Web)
- Check that your account is configured in Outlook
- Try restarting Outlook

### "Validation failed"

- Read the error messages carefully
- Common issues:
  - Missing email addresses
  - Duplicate email addresses
  - Invalid date format
- Fix the Excel file and run again

### "Excel is already open"

- Close all Excel windows
- Wait 10 seconds
- Run the tool again

### Emails not sending

- Verify `DRY_RUN=false` in `.env`
- Check Outlook is running and connected to internet
- Make sure you confirmed with "yes" when prompted

---

## Need Help?

If you encounter problems:

1. Check the log file in `logs/` folder
2. Look for error messages in red text
3. Try the troubleshooting steps above
4. Contact IT support if the issue persists

---
