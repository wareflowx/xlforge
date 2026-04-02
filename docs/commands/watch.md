# Watch Commands (Reactive Triggers)

The watch module is the "Nervous System" of xlforge. It transitions from a **Utility** (something you call) to a **Platform** (something that calls you).

---

## watch

Runs a background daemon that watches for changes and triggers actions.

```bash
xlforge watch <file.xlsx> [options]
```

### Options

```
--on-change <pattern>    # Cell/range pattern to watch (supports globbing)
--on-condition <formula> # Fire only when condition is met
--run <command>          # Local script to execute
--webhook <url>          # HTTP POST webhook
--poll <seconds>         # Poll interval (default: 2)
--debounce <duration>    # Coalesce rapid changes (default: 200ms)
--headless               # Watch closed files via file system events
--log <file>            # Log all changes to file
--all                    # Watch entire sheet
```

---

## Watch Patterns

### Cell Patterns

```bash
# Watch specific cell
xlforge watch report.xlsx --on-change "Summary!B2"

# Watch entire column
xlforge watch report.xlsx --on-change "Data!B:B"

# Watch entire sheet
xlforge watch report.xlsx --on-change "Summary!*"

# Watch multiple patterns
xlforge watch report.xlsx \
    --on-change "Summary!Revenue" \
    --on-change "Summary!Status"
```

### Glob Support

| Pattern | What It Watches |
|---------|-----------------|
| `Summary!B2` | Single cell B2 |
| `Data!B:B` | All cells in column B |
| `Data!2:2` | All cells in row 2 |
| `Summary!*` | All cells on Summary sheet |
| `Data!A1:Z100` | Range A1:Z100 |

---

## Debounce (Coalescing)

Prevents script flooding when a user pastes 100 cells or a formula triggers chain reactions.

```bash
# Wait 500ms after last change before firing
xlforge watch report.xlsx \
    --on-change "Data!*:" \
    --debounce 500ms \
    --run "./process_batch.sh"
```

**Logic:** If 10 changes happen in 500ms, only fire once with the final state.

**Durations:** `100ms`, `500ms`, `1s`, `2s`, etc.

---

## Edit Mode Safety

**Critical:** If the poll hits while the user is typing in a cell (Edit Mode), Excel will crash or beep.

### App.Ready Check

Before polling, the daemon checks `Application.Ready`. If Excel is busy:
1. **Skip this poll** (don't crash)
2. **Wait for next interval** (gentle polling)

```bash
# Default behavior is safe (checks App.Ready)
xlforge watch report.xlsx --on-change "Summary!B2"
```

### Force Mode (Not Recommended)

```bash
# Only use if Excel is guaranteed closed
xlforge watch report.xlsx --on-change "Summary!B2" --force
```

---

## Trigger Actions

### Local Script

```bash
xlforge watch report.xlsx \
    --on-change "Summary!B2" \
    --run "./notify.sh"
```

### Webhook

```bash
xlforge watch report.xlsx \
    --on-change "Summary!Revenue" \
    --webhook "https://hooks.slack.com/services/xxx/..."
```

### Multiple Triggers

```bash
xlforge watch report.xlsx \
    --on-change "Summary!Status" \
    --run "./local_handler.sh" \
    --webhook "https://api.example.com/webhook"
```

---

## Environment Variables

When `--run` executes, these env vars are available:

| Variable | Description |
|----------|-------------|
| `XLFORGE_FILE` | Full path to the watched file |
| `XLFORGE_SHEET` | Sheet name where change occurred |
| `XLFORGE_CELL` | Cell coordinate (e.g., `Summary!B2`) |
| `XLFORGE_VALUE` | New cell value |
| `XLFORGE_PREV_VALUE` | Previous cell value (before change) |
| `XLFORGE_CHANGE_COUNT` | Number of cells changed in this event |

### notify.sh Example

```bash
#!/bin/bash
echo "File: $XLFORGE_FILE"
echo "Cell: $XLFORGE_CELL"
echo "Changed: $XLFORGE_PREV_VALUE → $XLFORGE_VALUE"
echo "Sheet: $XLFORGE_SHEET"

# Example: Alert Slack
curl -X POST "$WEBHOOK_URL" \
    -d "{\"text\": \"Cell $XLFORGE_CELL changed to $XLFORGE_VALUE\"}"
```

---

## Webhook Payload

The webhook sends a POST request with JSON body:

```json
{
  "event": "cell_change",
  "file": "/path/to/report.xlsx",
  "sheet": "Summary",
  "cell": "B2",
  "value": 50000,
  "prev_value": 45000,
  "timestamp": "2026-03-31T10:30:00Z",
  "triggered_by": "human"
}
```

### Webhook Headers

```
Content-Type: application/json
X-XLForge-Event: cell_change
X-XLForge-File: report.xlsx
```

---

## Logic-Based Triggers (Conditions)

Fire only when a specific condition is met:

```bash
xlforge watch report.xlsx \
    --on-change "Summary!Total_Revenue" \
    --condition "val > 10000" \
    --run "./celebrate.sh"
```

### Condition Syntax

| Condition | Description |
|-----------|-------------|
| `val > 1000` | Numeric comparison |
| `val == "Approved"` | String comparison |
| `prev != val` | Value changed |
| `val == ""` | Cell is empty |

### Complex Conditions

```bash
# Trigger when revenue exceeds threshold
xlforge watch report.xlsx \
    --on-change "Summary!Revenue" \
    --condition "val > 10000 AND prev < 10000" \
    --run "./alert_supervisor.sh"
```

---

## Headless Watching

Watch files without opening Excel GUI. Uses file system events (`inotify`/`fsevents`) and `openpyxl`.

```bash
xlforge watch report.xlsx \
    --headless \
    --on-change "A1" \
    --run "./process.sh"
```

### Headless vs. GUI Mode

| Mode | Use Case | Engine |
|------|----------|--------|
| **GUI** (default) | Open Excel, real-time events | xlwings COM |
| **Headless** | Closed file on shared drive | openpyxl + inotify |

### When to Use Headless

- File is on OneDrive/Dropbox/SharePoint
- Server environment (no Excel installed)
- Watch must survive Excel close
- Low-CPU background monitoring

---

## Audit Logging

Record every change to a log file:

```bash
# JSON Lines format (one JSON per line)
xlforge watch report.xlsx --all --log changes.jsonl

# CSV format
xlforge watch report.xlsx --all --log changes.csv
```

### Log Output (JSONL)

```json
{"ts":"2026-03-31T10:30:00Z","sheet":"Summary","cell":"B2","prev":45000,"value":50000,"triggered_by":"human"}
{"ts":"2026-03-31T10:31:15Z","sheet":"Data","cell":"C5","prev":"Active","value":"Closed","triggered_by":"human"}
```

### Audit with All Changes

```bash
# Watch entire sheet and log everything
xlforge watch report.xlsx --all --log audit.jsonl
```

---

## Daemon Lifecycle

### PID File Management

Watchers store PID in `.xlforge/watchers/<file_hash>.pid`:

```bash
# Daemon is automatically assigned a PID
xlforge watch report.xlsx --on-change "A1" &

# Find the PID
cat .xlforge/watchers/abc123.pid

# Stop it
xlforge watch stop report.xlsx
```

### Graceful Self-Termination

If Excel closes while watching:
1. Daemon detects Excel process is gone
2. Waits `timeout` seconds for Excel to reopen (user might re-open)
3. If timeout exceeded, **self-terminates** (no zombie watchers)

```bash
# Wait 60 seconds before auto-kill
xlforge watch report.xlsx --timeout 60
```

### Safe Stop

```bash
# Gracefully stops the watcher
xlforge watch stop report.xlsx

# Force kill (if graceful fails)
xlforge watch stop report.xlsx --force
```

---

## watch status

Check if a watcher is active.

```bash
xlforge watch status <file.xlsx>
```

### Output

```json
{
  "file": "report.xlsx",
  "watching": true,
  "pid": 12345,
  "pattern": "Summary!*",
  "started_at": "2026-03-31T10:00:00Z",
  "trigger_count": 42,
  "mode": "gui"
}
```

---

## watch list

List all active watchers.

```bash
xlforge watch list
```

### Output

```json
{
  "watchers": [
    {"file": "report.xlsx", "pid": 12345, "pattern": "Summary!B2"},
    {"file": "dashboard.xlsx", "pid": 12346, "pattern": "Data!*"}
  ]
}
```

---

## watch log

Read the change log.

```bash
xlforge watch log <file.xlsx>
xlforge watch log <file.xlsx> --last 10
xlforge watch log <file.xlsx> --since "2026-03-31T10:00:00"
```

---

## Complete Examples

### Real-time dashboard update

```bash
# Watch revenue, update dashboard when it changes
xlforge watch report.xlsx \
    --on-change "Summary!Revenue" \
    --debounce 1s \
    --run "./refresh_dashboard.sh"

# ./refresh_dashboard.sh
#!/bin/bash
xlforge table refresh dashboard.xlsx "SalesData" --sync
xlforge run update_charts.xlf
```

### Slack alerts for approval workflow

```bash
# Watch status column for approvals
xlforge watch report.xlsx \
    --on-change "Data!Status" \
    --condition "val == 'Approved'" \
    --webhook "https://hooks.slack.com/services/xxx" \
    --run "./approved_handler.sh"
```

### Headless file monitoring (server)

```bash
# On a shared drive, watch without opening Excel
xlforge watch /shared/reports/sales.xlsx \
    --headless \
    --on-change "Summary!Total" \
    --debounce 2s \
    --log /var/log/xlforge/sales_changes.jsonl
```

### Audit trail

```bash
# Create audit log of all changes
xlforge watch report.xlsx \
    --all \
    --log audit_$(date +%Y%m%d).jsonl

# Query the audit log
xlforge watch log report.xlsx --last 100 | jq '.[] | select(.cell | startswith("Data!"))'
```

### Multi-condition trigger

```bash
xlforge watch report.xlsx \
    --on-change "Summary!*:" \
    --debounce 500ms \
    --condition "changed_cells >= 5" \
    --run "./batch_import_handler.sh"
```

---

## Error Codes

| Code | Meaning |
|------|---------|
| `120` | Watcher already active for this file |
| `121` | No active watcher to stop |
| `122` | Watcher PID not found (stale pid file) |
| `123` | Headless mode not supported on this platform |
| `124` | File does not exist |
| `125` | Condition syntax error |
| `126` | Watcher timeout (Excel closed, no re-open) |
| `127` | App.Ready check failed (Excel in bad state) |
