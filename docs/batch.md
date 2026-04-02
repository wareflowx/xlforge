# Batch Execution

## run

Executes commands from a script file in a single COM session.

```bash
xlforge run <script.xlf>
xlforge run <script.xlf> --dry-run        # Preview without executing
xlforge run <script.xlf> --transaction    # Atomic: all or nothing
xlforge run <script.xlf> --debug         # Visual HUD overlay
```

---

## Transaction Mode

If a command fails, the file is rolled back to its pre-script state:
```bash
xlforge run <script.xlf> --transaction
# If command 5 fails, commands 1-4 are undone
```

---

## Debug Mode (Visual HUD)

Adds a semi-transparent overlay on the Excel window showing:
- Current command being executed
- Red border around active cell
- Tooltip with operation description
- Progress indicator

```
┌─────────────────────────────────────┐
│ [3/15] cell formula Summary!B3     │
│ ━━━━━━━━━━━━━━━━━━━░░░░░░░░░░░░░░ │
│ Cell: Summary!B3 =SUM(Data!B:B)    │
└─────────────────────────────────────┘
```

---

## .xlf Script Format

**Must include version header:**
```
version 1.0

# report.xlf - Generate Q1 report
file open report.xlsx
sheet create "Summary"
sheet use report.xlsx "Summary"
cell set A1 "Q1 2026 Performance"
format cell A1 --bold --size 16 --color "44,72,116"
cell formula B3 "=SUM(Data!B:B)"
format cell B3 --number-format "€#,##0"
import csv report.xlsx sales.csv --sheet "Data" --cell A1 --has-headers
table create report.xlsx "Data!A1" sales.csv --style zebra --freeze-header
column auto-fit report.xlsx "Data!A:A"
freeze report.xlsx "Data!A2"
file save report.xlsx
```

**Version header** ensures old scripts don't silently fail on syntax changes.

---

## Fast-Path Mode (Direct-to-OpenXML)

For large datasets, xlforge bypasses COM entirely when the file is closed.

### How it works

```
Closed File ──► xlsxwriter/openpyxl ──► Stream to .xlsx ──► COM Refresh
     │                                                    │
     │         (Direct XML write)                        │
     └────────────────────────────────────────────────────┘
                      (Single COM call for refresh)
```

### Usage

Fast-Path is automatic when:
1. Target file is closed
2. Data size > 1000 rows
3. No active Excel instance editing the file

```bash
# This automatically uses Fast-Path
xlforge import csv report.xlsx large_data.csv --sheet Data --cell A1
# Writes 100k rows in < 1 second
```

### Force modes

```bash
xlforge import csv report.xlsx data.csv --sheet Data --cell A1 --fast-path true
xlforge import csv report.xlsx data.csv --sheet Data --cell A1 --fast-path false
```
