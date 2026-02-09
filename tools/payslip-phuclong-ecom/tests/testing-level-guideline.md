# Recommended Testing Levels (Windows Python Automation Tool)

## 1️⃣ Unit Testing (Mandatory)

### Purpose

Validate individual functions in isolation.

### Scope

- Email validation logic
- Duplicate detection logic
- DATE parsing
- File naming logic
- Idempotency logic
- Summary calculations
- Configuration loading (.env precedence logic)

### Rules

- No real Excel
- No real Outlook
- Use mocks for:
  - win32com
  - filesystem
  - environment variables

### Tools

- `pytest`
- `unittest.mock`
- `pytest-mock`

### Target Coverage

> 70–85% logical coverage (not COM lines)

---

## 2️⃣ Component / Service-Level Tests (Mocked COM Layer)

### Purpose

Test integration between internal modules while mocking external dependencies.

Example:

- Email service with mocked Outlook COM
- Excel service with mocked COM object
- PDF generation wrapper

### Scope

- Validate that correct COM methods are called
- Validate correct parameters passed to:
  - `ExportAsFixedFormat`
  - `Send()`
  - Workbook open/close

- Simulate COM failures

### Why Important

COM automation is fragile. You must test:

- Error handling branches
- Exception propagation
- Resource cleanup logic

---

## 3️⃣ Integration Testing (Real Excel + Real Outlook) — Manual or Controlled

### Purpose

Verify real Windows environment behavior.

This is CRITICAL for COM-based tools.

### Scope

- Real Excel instance
- Real Outlook profile
- Sample Excel file (2 employees)
- DRY_RUN=false

Validate:

- Excel opens correctly
- Formulas recalculate
- PDF generated
- Password applied
- Email sent
- No orphan Excel processes remain

### Execution

- Manual trigger only
- Clearly marked:

  ```
  @pytest.mark.integration
  ```

### Run Frequency

- Before production release
- After major refactor
- After Office updates

---

## 4️⃣ End-to-End (E2E) Operational Test

### Purpose

Simulate real production run (1000+ employees).

### Scope

- Full run with large dataset
- Measure:
  - Execution time
  - Memory growth
  - Excel process stability
  - Log correctness
  - Idempotency behavior (rerun test)

### Validate

- No duplicate sending
- Correct summary output
- Stable Excel lifecycle

This is usually done in a staging environment.

---

# Optional (Recommended for Enterprise-Level Tools)

## 5️⃣ Non-Functional Testing

### Performance Testing

- 1000+ employees
- Measure:
  - Execution time
  - Memory usage
  - Excel handle leakage

### Resilience Testing

- Kill Excel mid-process
- Disconnect Outlook
- Lock output folder

### Security Testing

- Verify password protection actually works
- Attempt to open PDF without password

---

# Minimum Acceptable Levels

For your specific Payslip Tool, the minimum professional setup should be:

| Level                  | Required               |
| ---------------------- | ---------------------- |
| Unit                   | ✅ Mandatory           |
| Component (mocked COM) | ✅ Mandatory           |
| Integration (real COM) | ✅ Mandatory           |
| E2E large dataset      | ⚠ Strongly recommended |

---

# Why This Matters for Windows COM Tools

Windows COM introduces risks not present in pure Python apps:

- Orphan Excel processes
- Threading issues
- Hidden Excel UI popups
- Office version differences
- Outlook security prompts
- File locks

Without integration testing, these issues will only appear in production.

---

# Practical Structure Example

```
tests/
├── unit/
├── component/
├── integration/
│   └── test_real_excel_outlook.py
└── e2e/
```

---
