# Cell Commands

Cell references use `<sheet>!A1` notation. Examples: `Data!A1`, `Summary!B2:C10`

---

## cell get

Reads a cell value.

```bash
xlforge cell get <file.xlsx> <sheet!cell>
xlforge cell get <file.xlsx> <sheet!cell> --json
xlforge cell get <file.xlsx> <sheet!cell> --formula   # Get formula, not value
xlforge cell get <file.xlsx> <sheet!cell> --calculate  # Force recalculation first
```

**JSON output:**
```json
{
  "cell": "Data!A1",
  "value": "Sales Report",
  "type": "string",
  "formula": null
}
```

With `--formula`:
```json
{
  "cell": "Data!B3",
  "value": 123456,
  "type": "number",
  "formula": "=SUM(Data!B:B)"
}
```

**Dates** are returned in ISO 8601 format: `"2026-01-15T00:00:00"` not Python datetime objects.

---

## cell set

Sets a cell value.

```bash
xlforge cell set <file.xlsx> <sheet!cell> <value>
xlforge cell set <file.xlsx> <sheet!cell> <value> --type <type>
```

**Options:**
```
--type <type>     # Force type: string, number, date, bool, formula
--as <type>       # Alias for --type
```

**Type examples:**
```bash
xlforge cell set report.xlsx "Data!A1" "00123" --type string  # Preserves leading zeros
xlforge cell set report.xlsx "Data!B1" "2026-01-15" --type date
xlforge cell set report.xlsx "Data!C1" "=SUM(A:A)" --type formula
```

**Stdin support:**
```bash
cat long_text.txt | xlforge cell set report.xlsx "Data!A1" -
curl -s "https://api.example.com/value" | xlforge cell set report.xlsx "Data!A1" -
```

**Why `--type` matters:** Excel often mangles data (e.g., "1-1" → Date, "00123" → 123). Forcing the type prevents 90% of "Excel ruined my data" issues.

---

## cell formula

Sets a formula.

```bash
xlforge cell formula <file.xlsx> <sheet!cell> <formula>
```

**Example:**
```bash
xlforge cell formula report.xlsx "Summary!B2" "=SUM(Data!B:B)"
xlforge cell formula report.xlsx "Summary!C3" "=AVERAGE(Data!C:C)"
```

---

## cell clear

Clears a cell content and formatting.

```bash
xlforge cell clear <file.xlsx> <sheet!cell>
xlforge cell clear <file.xlsx> <sheet!cell> --format-only   # Keep value
xlforge cell clear <file.xlsx> <sheet!cell> --value-only   # Keep formatting
```

---

## cell copy

Copies a cell to another location.

```bash
xlforge cell copy <file.xlsx> <src-cell> <dst-cell>
```

**Cross-workbook copy:**
```bash
xlforge cell copy report.xlsx "Template!A1" "Report.xlsx!Data!A1"
```

---

## cell bulk

Bulk operations on ranges. **Uses array operations internally for performance.**

```bash
xlforge cell bulk <file.xlsx> <pattern> [options]
```

**Pattern syntax:**
- `<sheet>!*` → All cells in UsedRange (not 17 billion empty cells)
- `<sheet>!A*` → All cells in column A that have data
- `<sheet>!1:*` → All cells in row 1
- `<sheet>!A1:Z100` → Explicit range

**Options:**
```
--filter <filter>      # empty, non-empty, formula, value
--format <transform>   # uppercase, lowercase, trim, proper
--set <value>          # Set all matching cells to value
--clear                # Clear matching cells
```

**Examples:**
```bash
# Uppercase all text in column A
xlforge cell bulk report.xlsx "Data!A*" --filter non-empty --format uppercase

# Clear all empty cells in a range
xlforge cell bulk report.xlsx "Data!A1:Z100" --filter empty --clear

# Set all formula cells to their calculated value
xlforge cell bulk report.xlsx "Data!*" --filter formula --set evaluated
```

**Performance:** Internally uses range array operations (pull → transform → push), not individual cell COM calls. 100x faster than cell-by-cell iteration.

---

## cell search

Finds a cell by its content. Essential for agents that don't know coordinates.

```bash
xlforge cell search <file.xlsx> <query>
xlforge cell search <file.xlsx> <query> --json
xlforge cell search <file.xlsx> <query> --sheet <name>  # Limit to sheet
```

**Output:**
```bash
xlforge cell search report.xlsx "Total Revenue"
# Found: Summary!B42 (contains "Total Revenue")
```

**JSON output:**
```json
{
  "query": "Total Revenue",
  "cell": "Summary!B42",
  "sheet": "Summary",
  "value": "Total Revenue",
  "type": "string"
}
```

---

## cell fill

Excel's Auto-fill feature exposed to CLI.

```bash
xlforge cell fill <file.xlsx> <range> --direction <direction>
```

**Options:**
```
--direction <dir>      # down, up, left, right (default: down)
--stop <value>         # Stop when reaching this value
```

**Examples:**
```bash
# Fill a date series
xlforge cell fill report.xlsx "Data!A1:A10" --direction down

# Fill a number series
xlforge cell fill report.xlsx "Data!B1:B5" --direction down

# Fill formula across
xlforge cell fill report.xlsx "Data!C1:H1" --direction right
```

---

## Range Operations

### cell get (range)

Reads a range of cells.

```bash
xlforge cell get <file.xlsx> <sheet!range>
xlforge cell get <file.xlsx> <sheet!range> --json
xlforge cell get <file.xlsx> <sheet!range> --headers   # First row as keys
```

**With `--headers`, returns keyed objects:**
```bash
xlforge cell get report.xlsx "Data!A1:B2" --headers --json
```
```json
[
  {"Name": "Alice", "Score": 95},
  {"Name": "Bob", "Score": 88}
]
```

**Without `--headers`, returns array of arrays:**
```json
{
  "range": "Data!A1:B2",
  "data": [["Name", "Score"], ["Alice", 95], ["Bob", 88]]
}
```

---

## Data Type Handling

### Type Inference

When setting values, xlforge infers types by default:

| Input | Inferred Type | Example |
|-------|---------------|---------|
| `123` | number | `123` |
| `"text"` | string | `"text"` |
| `2026-01-15` | date | Serial date |
| `TRUE`/`FALSE` | bool | `TRUE` |
| `=FORMULA` | formula | `=SUM(A:A)` |

### Type Coercion

Use `--type` to force a specific type:

| Type | Behavior |
|------|---------|
| `string` | Preserves leading zeros, prevents date conversion |
| `number` | Converts to numeric value |
| `date` | Converts to Excel serial date |
| `bool` | Converts to TRUE/FALSE |
| `formula` | Treats value as formula |

**Examples preventing data mangling:**
```bash
# Preserve leading zeros
xlforge cell set report.xlsx "Data!A1" "00123" --type string

# Prevent "1-3" becoming "Mar 1"
xlforge cell set report.xlsx "Data!B1" "1-3" --type string

# Explicit date
xlforge cell set report.xlsx "Data!C1" "2026-01-15" --type date
```

---

## Stale Value Problem

When reading after writing, Excel may return cached values.

### Solution: --calculate

```bash
xlforge cell set report.xlsx "A1" 10
xlforge cell formula report.xlsx "B1" "=A1*2"
xlforge cell get report.xlsx "B1"              # May return stale value
xlforge cell get report.xlsx "B1" --calculate   # Forces recalc, returns 20
```

---

## Error Codes

| Code | Meaning |
|------|---------|
| `4` | Cell not found |
| `5` | Invalid syntax |
| `11` | Type coercion failed |
| `12` | Range too large (use `cell bulk` instead) |
