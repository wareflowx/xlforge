# Data Commands

---

## import csv

Imports a CSV file into a sheet.

```bash
xlforge import csv <file.xlsx> <data.csv> --sheet <name> --cell <cell> [options]
```

**Accepts stdin:**
```bash
cat data.csv | xlforge import csv <file.xlsx> --sheet Data --cell A1
curl -s "https://api.example.com/data" | xlforge import csv <file.xlsx> --sheet Data --cell A1
sql-cli query "..." | xlforge import csv <file.xlsx> --sheet Data --cell A1
```

---

### Core Options

### sheet, cell

Target location.

```bash
--sheet <name>   # Target sheet (creates if not exists)
--cell <cell>    # Top-left anchor (e.g., A1)
```

### mode

Import behavior.

```bash
--mode <mode>    # append or overwrite (default: overwrite)
```

| Mode | Behavior |
|------|----------|
| `overwrite` | Replace cells from anchor (default) |
| `append` | Find last row with data, insert after |

```bash
xlforge import csv data.csv --sheet Logs --cell A1 --mode append
# Automatically finds last row and appends
```

---

### Data Type Options

### encoding

CSV character encoding.

```bash
--encoding <encoding>    # Default: utf-8-sig (Excel preferred)
```

**Options:** `utf-8`, `utf-8-sig`, `latin-1`, `ascii`

**Why `utf-8-sig`?** Excel prefers UTF-8 with BOM for proper character display.

### types

Force column types to prevent Excel auto-formatting.

```bash
--types "<col1>:<type>,<col2>:<type>,..."
```

**Types:** `string`, `number`, `date`, `bool`

```bash
xlforge import csv data.csv --sheet Data --cell A1 \
    --types "ID:string,Amount:number,Date:date"
```

**Prevents:**
- `00123` becoming `123` (use `string`)
- `1-3` becoming `Mar 1` (use `string`)
- Leading zeros being stripped

### strings-only

Force all columns to text format.

```bash
--strings-only
```

---

### Table Creation

### to-table

Immediately convert imported range to an Excel Table (ListObject).

```bash
--to-table <name>
```

```bash
xlforge import csv data.csv --sheet Data --cell A1 \
    --to-table "SalesTable" \
    --style zebra
```

**Benefits:**
- Adds auto-filter
- Banded rows (if style applied)
- Named reference for Pivot Tables
- Structured references (`=SUM(SalesTable[Amount])`)

### style

Table style when using `--to-table`.

```bash
--style <style>    # zebra, bordered, plain, grid (default: plain)
```

---

### Data Shaping

### columns

Import only specific columns from CSV.

```bash
--columns "<col1>,<col2>,..."
```

```bash
xlforge import csv data.csv --sheet Data --cell A1 \
    --columns "Date,Region,Revenue"
```

### limit

Limit number of rows to import.

```bash
--limit <n>
```

```bash
xlforge import csv big_data.csv --sheet Data --cell A1 --limit 100
```

### skip

Skip first N rows (e.g., metadata rows).

```bash
--skip <n>
```

```bash
xlforge import csv data.csv --sheet Data --cell A1 --skip 2
```

---

### Validation

### validate

Check if CSV structure matches existing data before importing.

```bash
--validate
--validate --strict
```

**Default behavior:** Warning with diff
**With `--strict`:** Error on mismatch

**Example output:**
```
WARNING: Header mismatch
  Expected: [Date, Revenue, Region]
  Got:      [Date, Sales, Region]
  Mapping: Revenue → Sales (renamed)
```

---

### Performance

### fast-path

Direct-to-OpenXML writing for large files.

```bash
--fast-path <auto|true|false>
```

| Rows | Default Behavior |
|------|------------------|
| < 1,000 | COM (live update) |
| > 1,000 | Auto Fast-Path |
| Any | Force with `--fast-path true` |

**Fast-Path:** Bypasses COM, writes XML directly. 50k rows in < 1 second.

---

## export csv

Exports a sheet to CSV.

```bash
xlforge export csv <file.xlsx> <sheet>
xlforge export csv <file.xlsx> <sheet> --file <output.csv>
```

**If `--file` omitted:** Outputs to stdout.

---

### Core Options

### headers

Include header row.

```bash
--headers
```

### formatted

Export displayed values instead of raw values.

```bash
--formatted
```

| Without `--formatted` | With `--formatted` |
|-----------------------|---------------------|
| `1` | `$1.00` |
| `0.5` | `50%` |
| `44566` | `2026-01-15` |

### encoding

Output encoding.

```bash
--encoding <encoding>    # Default: utf-8
```

### range

Export specific range instead of entire sheet.

```bash
--range <range>
```

```bash
xlforge export csv report.xlsx "Data" --range A1:Z100 --file output.csv
```

---

## Complete Examples

### Basic import
```bash
xlforge import csv sales.csv --sheet Data --cell A1 --has-headers
```

### Type-safe import with table
```bash
xlforge import csv data.csv \
    --sheet Sales \
    --cell A1 \
    --types "ID:string,Amount:number,Date:date" \
    --to-table "SalesData" \
    --style zebra
```

### Append daily data
```bash
xlforge import csv today_sales.csv \
    --sheet Logs \
    --cell A1 \
    --mode append \
    --encoding utf-8-sig
```

### Import with column selection
```bash
xlforge import csv full_export.csv \
    --sheet Summary \
    --cell A1 \
    --columns "Date,Revenue,Expenses" \
    --limit 100
```

### Export with formatting
```bash
xlforge export csv report.xlsx "Data" --headers --formatted --file output.csv
```

### Pipeline usage
```bash
sql-cli query "SELECT * FROM sales" | \
    xlforge import csv report.xlsx --sheet Data --cell A1 \
    --types "ID:string,Amount:number"
```

---

## Performance Notes

### Engine Selection

| File Size | Engine | Speed |
|-----------|--------|-------|
| < 1,000 rows | COM (xlwings) | Live update |
| > 1,000 rows | Fast-Path (OpenXML) | < 1 second |

### Fast-Path Process

```
1. Detect file closed + size > threshold
2. Close COM handle
3. Write to .xlsx via OpenXML (xlsxwriter)
4. Re-open COM for refresh
```

### Type Preservation

**Problem:** Excel auto-formats imported data (e.g., `00123` → `123`).

**Solution:** Use `--types` or `--strings-only` to force text format before import.

---

## Error Codes

| Code | Meaning |
|------|---------|
| `40` | CSV not found |
| `41` | Encoding error |
| `42` | Type coercion failed |
| `43` | Header mismatch (use `--validate`) |
| `44` | Sheet not found during export |
| `45` | Invalid CSV format |
