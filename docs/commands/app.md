# App Commands (Real-time Control)

App commands manage the Excel process and provide feedback during automation.

---

## app status

Health check - returns the state of the Excel engine before running batch operations.

```bash
xlforge app status <file.xlsx>
xlforge app status <file.xlsx> --json
```

**JSON output:**
```json
{
  "is_visible": true,
  "calculation_mode": "automatic",
  "calculation_state": "done",
  "is_ready": true,
  "is_dirty": true,
  "screen_updating": true,
  "pid": 12345,
  "version": "16.0 (Microsoft Excel 365)"
}
```

**Use before running a 100-line batch script to ensure Excel is ready.**

---

## app show / app hide

Shows or hides the Excel window.

```bash
xlforge app show <file.xlsx>
xlforge app hide <file.xlsx>
```

**Why separate commands?** Avoids accidental `true`/`false` typos.

**Example:**
```bash
xlforge app hide report.xlsx   # Hide window (faster for batch)
xlforge app show report.xlsx   # Show window
```

---

## app visible

Legacy form with explicit boolean (still supported).

```bash
xlforge app visible <file.xlsx> <true|false>
```

---

## app calculate-mode

Controls Excel's calculation mode. Essential for large SQL pushes.

```bash
xlforge app calculate-mode <file.xlsx> <automatic|manual>
```

**Example:**
```bash
# Set to manual before bulk operations
xlforge app calculate-mode report.xlsx manual

# ... push 50,000 rows ...

# Set back to automatic when done
xlforge app calculate-mode report.xlsx automatic
```

**Why?** Prevents Excel from recalculating after every cell write.

---

## app calculate

Forces Excel to recalculate all formulas.

```bash
xlforge app calculate <file.xlsx>
```

---

## app idle

Waits for Excel to be in idle state (not in edit mode). **Critical for reliability.**

```bash
xlforge app idle <file.xlsx> [--timeout <seconds>]
```

**Exit codes:**
- `0` - Excel is idle and ready
- `8` - Timeout reached (Excel still busy)

**Example:**
```bash
xlforge app idle report.xlsx --timeout 60
if [ $? -eq 8 ]; then
    echo "Excel still busy - killing and recovering"
    xlforge app recover report.xlsx
fi
```

**Checks:** `App.Ready` and `App.CalculationState` to ensure Excel isn't locked.

---

## app focus

Scrolls Excel to and selects a specific cell, then activates the sheet.

```bash
xlforge app focus <file.xlsx> <sheet!cell>
```

**Example:**
```bash
xlforge app focus report.xlsx "Data!B10"  # Switches to Data sheet, scrolls to B10
```

**Behavior:**
1. Activates the target sheet (switches tab if needed)
2. Scrolls to the cell
3. Selects the cell (shows visual highlight)

---

## app alert

Displays a native Excel message box.

```bash
xlforge app alert <file.xlsx> <message>
xlforge app alert <file.xlsx> <message> --timeout <seconds>
```

**Options:**
```
--timeout <seconds>    # Auto-dismiss after N seconds (default: wait forever)
--type <type>          # info, warning, error (default: info)
```

**Example:**
```bash
xlforge app alert report.xlsx "Processing complete!" --timeout 5
```

**Warning:** Without `--timeout`, this blocks indefinitely. Always use timeout in CI/CD pipelines.

---

## app silence

Disables Excel's native dialogs and popups that could block automation.

```bash
xlforge app silence <file.xlsx>
xlforge app unsilence <file.xlsx>
```

**What it silences:**
- Privacy warnings
- Links update prompts
- External data refresh warnings
- "File modified" dialogs

**Example:**
```bash
xlforge app silence report.xlsx
# ... batch operations ...
xlforge app unsilence report.xlsx
```

---

## app wait-idle

Alias for `app idle` (see above).

---

## screen-update

Controls screen updating during batch operations. **Can improve performance by 3-5x.**

```bash
xlforge screen-update <file.xlsx> <true|false>
```

**Safety:** `xlforge run` automatically restores screen-update to `true` on exit, even if the script crashes.

**Example:**
```bash
xlforge screen-update report.xlsx false
# ... many operations (faster) ...
xlforge screen-update report.xlsx true
```

---

## app recover

If Excel hangs, kills the process and reopens the file to the last save point.

```bash
xlforge app recover <file.xlsx>
xlforge app recover --pid <pid>
```

**Exit codes:**
- `0` - Successfully recovered
- `1` - Recovery failed

**Warning:** Unsaved changes may be lost.

---

## app kill

Force-kills the Excel process. Use as last resort.

```bash
xlforge app kill <file.xlsx>
xlforge app kill --pid <pid>
```

**Unlike `app recover`:** Does not attempt to reopen the file.

---

## Process Scoping

xlforge automatically pins itself to the specific Excel instance that has the target file open.

**How it works:**
1. When `xlforge file open <file.xlsx>` is called, xlforge records the Excel process PID
2. All subsequent commands for that file use the same PID
3. If the file isn't open, xlforge starts a new isolated instance

**Why?** Prevents confusion when multiple Excel instances are running.

**Example scenario:**
```bash
# User has 3 Excel files open
xlforge file open report1.xlsx  # Pins to Excel instance A
xlforge file open report2.xlsx  # Pins to Excel instance B
xlforge file open report3.xlsx  # Pins to Excel instance C
# Each command targets the correct instance automatically
```

---

## Safety Wrappers

### xlforge run

The `run` command automatically wraps operations with safety measures:

```bash
xlforge run <script.xlf> [options]
```

**Automatic safety:**
- Sets `screen-update false` at start
- Restores `screen-update true` on exit (even on crash)
- Sets `calculation-mode manual` at start
- Restores `calculation-mode automatic` on exit
- Cleans up COM references on failure

---

## Error Codes

| Code | Meaning |
|------|---------|
| `7` | COM error |
| `8` | Excel busy (timeout in `app idle`) |
| `10` | Excel hung (use `app recover`) |
