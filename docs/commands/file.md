# File Commands

File commands manage the Excel file lifecycle, process isolation, and environment health.

---

## file open

Creates a new file or opens an existing one.

```bash
xlforge file open <file.xlsx>
xlforge file open <file.xlsx> --engine <engine>
xlforge file open <file.xlsx> --template <name>
```

**Options:**
```
--engine <engine>       # xlwings (default) or openpyxl
--visible <true|false>  # Show/hide Excel window
--template <name>       # Use a saved template
```

**Implicit behavior:** If you run a command on a file that isn't open, xlforge automatically opens it, performs the action, and closes. Use explicit `file open` for:
- Setting visibility (`--visible false` for batch)
- Using a template
- Pre-warming COM for faster subsequent commands

**Exit codes:**
- `0` - Success
- `1` - File locked or permission denied
- `2` - File not found (for explicit open with `--must-exist`)
- `50` - Engine mismatch (file opened with different engine)

---

## file save

Saves the workbook.

```bash
xlforge file save <file.xlsx>
xlforge file save <file.xlsx> --output <new-file.xlsx>
```

**Options:**
```
--output <file>    # Save to new file
--dry-run          # Validate without writing
```

**Examples:**
```bash
xlforge file save report.xlsx
xlforge file save report.xlsx --output "backup/report_v2.xlsx"
```

Auto-save is enabled by default. Use `--no-save` on other commands to disable.

---

## file close

Closes the workbook and releases COM resources.

```bash
xlforge file close <file.xlsx>
xlforge file close <file.xlsx> --force
```

**Options:**
```
--force    # Force close even if unsaved changes
```

---

## file info

Displays workbook metadata.

```bash
xlforge file info <file.xlsx>
xlforge file info <file.xlsx> --json
```

**JSON output:**
```json
{
  "file": "report.xlsx",
  "path": "C:/Users/name/report.xlsx",
  "absolute_path": "C:/Users/name/report.xlsx",
  "sheets": ["Summary", "Data", "Analysis"],
  "active": "Summary",
  "is_dirty": false,
  "is_open": true,
  "engine": "xlwings",
  "pid": 12345,
  "version": "16.0 (Microsoft Excel 365)",
  "size_bytes": 1048576
}
```

**New fields:**
- `is_dirty`: Boolean indicating unsaved changes
- `absolute_path`: Always absolute (resolves `./` and `../`)

---

## file kill

Force-kills the specific Excel process holding the file handle.

```bash
xlforge file kill <file.xlsx>
xlforge file kill <file.xlsx> --force
xlforge file kill --pid <pid>
```

**Behavior:**
1. Finds the PID that has `file.xlsx` open
2. Kills **only** that process
3. Leaves other Excel windows untouched

**Options:**
```
--force    # Kill without confirmation
--pid      # Kill specific PID directly
```

**Safety:** Prefer `file recover` first. Use `file kill` only when Excel is hung.

---

## file recover

Recovers from a hung Excel instance by killing and reopening to last save.

```bash
xlforge file recover <file.xlsx>
xlforge file recover <file.xlsx> --force
```

**Behavior:**
1. Kills the Excel process holding the file
2. Reopens the file to the last saved state
3. Preserves checkpoint history

**Exit codes:**
- `0` - Successfully recovered
- `1` - Recovery failed

**Warning:** Unsaved changes may be lost.

---

## file check

Analyzes file health and reports issues.

```bash
xlforge file check <file.xlsx>
xlforge file check <file.xlsx> --json
xlforge file check <file.xlsx> --repair
```

**JSON output:**
```json
{
  "file": "report.xlsx",
  "health": "good",
  "issues": [],
  "size_bytes": 1048576,
  "optimized_size_bytes": 851234,
  "savings_percent": 18.7,
  "checks": {
    "corruption": "passed",
    "unused_styles": "warning",
    "phantom_objects": "passed",
    "linked_files": "passed"
  }
}
```

**Issues detected:**
- Unused styles bloating file
- Broken links
- Phantom shapes/objects
- Corruption markers

**Options:**
```
--repair    # Automatically fix issues
```

---

## file monitor

Streams file change events (for agent waiting).

```bash
xlforge file monitor <file.xlsx>
xlforge file monitor <file.xlsx> --timeout <seconds>
```

**Output:**
```
[10:01:02] Cell Summary!B2 changed by "User"
[10:01:05] Sheet "Data" added
[10:02:30] Formula updated in Data!C5
```

**Use case:** Agent waits for human to make manual changes before proceeding.

---

## file template

Manages template library.

```bash
xlforge file template list
xlforge file template add <name> <file.xlsx>
xlforge file template delete <name>
```

**Example:**
```bash
xlforge file template add "monthly_budget" templates/budget.xlsx
xlforge file open report.xlsx --template "monthly_budget"
```

---

## Path Resolution

xlforge automatically resolves all paths to absolute paths before passing to Excel COM.

**Why:** Excel COM hates relative paths. `./report.xlsx` often fails.

**What xlforge does:**
```bash
./reports/../report.xlsx  →  C:/Users/name/report.xlsx
~/Documents/report.xlsx   →  C:/Users/name/Documents/report.xlsx
```

**Always stored as absolute in context.**

---

## Process Scoping

xlforge pins commands to specific Excel instances by PID.

**How it works:**
1. `file open` records the PID of the Excel instance
2. All subsequent commands for that file use the same PID
3. Multiple Excel instances are isolated

**Scenario:**
```bash
# User has 3 Excel files open (PIDs 100, 101, 102)
xlforge file open report1.xlsx  # Pins to PID 100
xlforge file open report2.xlsx  # Pins to PID 101
# Each command targets the correct instance
```

---

## Engine Strategy

### xlwings (default on Windows/Mac)

Full feature support:
- Macros
- Charts
- Formatting
- Live Excel interaction

### openpyxl (headless/Linux)

Limited but cross-platform:
- Read/write cells
- Sheet operations
- Basic formatting

**Engine mismatch:** If file is opened with `openpyxl` and you try to use `chart create` (requires xlwings), xlforge returns error 50 with a clear message.

---

## Complete Examples

### Basic workflow (implicit open)
```bash
xlforge cell set report.xlsx "A1" "Hello"
# Automatically opens, sets, and closes
```

### Batch mode (explicit open with hidden window)
```bash
xlforge file open report.xlsx --visible false
xlforge cell bulk report.xlsx "Data!*" --filter empty --clear
xlforge file save report.xlsx
xlforge file close report.xlsx
```

### Safe recovery
```bash
xlforge file check report.xlsx --repair
xlforge file recover report.xlsx
```

### Template workflow
```bash
xlforge file template add "budget" budget_template.xlsx
xlforge file open monthly.xlsx --template "budget"
```

### Agent waiting for human
```bash
xlforge file monitor report.xlsx --timeout 3600
# Script waits until human makes changes or timeout
```

---

## Error Codes

| Code | Meaning |
|------|---------|
| `2` | File not found |
| `50` | Engine mismatch |
| `51` | File corrupted |
| `52` | Cannot kill (file in use by another process) |
| `53` | Template not found |
| `54` | Recovery failed |
