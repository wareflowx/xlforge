# Reference

## Error Codes

| Code | Meaning |
|------|---------|
| `0` | Success |
| `1` | General error |
| `2` | File not found |
| `50` | Engine mismatch |
| `51` | File corrupted |
| `52` | Cannot kill (file in use) |
| `53` | Template not found |
| `54` | Recovery failed |
| `3` | Sheet not found |
| `4` | Cell not found |
| `5` | Invalid syntax |
| `6` | File locked (retry with backoff) |
| `7` | COM error |
| `8` | Excel is busy (timeout in `app idle`) |
| `9` | Feature unavailable (e.g., chart in headless mode) |
| `10` | Excel hung (use `app recover`) |
| `11` | Type coercion failed |
| `12` | Range too large (use `cell bulk` instead) |
| `13` | Chart not found |
| `14` | Invalid chart type |
| `15` | Chart name already exists (use `--replace`) |
| `20` | Checkpoint not found |
| `21` | Checkpoint restore failed |
| `22` | Branch not found |
| `23` | Branch merge conflict |
| `24` | Cannot delete active branch |
| `30` | Column not found |
| `31` | Row not found |
| `32` | Invalid unit (use px, pt, or excel) |
| `33` | Column/row hidden |
| `40` | CSV not found |
| `41` | Encoding error |
| `42` | Type coercion failed |
| `43` | Header mismatch |
| `44` | Sheet not found during export |
| `45` | Invalid CSV format |
| `60` | Invalid style string |
| `61` | Invalid number format |
| `62` | Named style not found |
| `63` | Conditional format not supported |
| `64` | Range too complex for conditional format |
| `70` | Sheet is protected |
| `71` | Password required |
| `72` | Invalid password |
| `73` | Workbook is protected |
| `74` | Cannot unhide very-hidden sheet |
| `75` | Cell is locked |
| `76` | Invalid protection option |
| `77` | Cannot delete last sheet (workbook must have at least one) |
| `78` | Circular sheet reference in move |
| `79` | Cannot move sheet that doesn't exist |
| `80` | Index not found (run `index create` first) |
| `81` | LLM provider error (quota, network) |
| `82` | Privacy check failed (sensitive data detected) |
| `83` | Confidence below threshold |
| `84` | Recording already active |
| `85` | No active recording to stop |
| `86` | Watchdog timeout (Excel connection lost) |
| `87` | Invalid rule syntax for semantic-check |
| `88` | Type coercion failed (use `--strict` for details) |
| `89` | Database connection failed |
| `90` | Query timeout |
| `91` | File is locked (Excel has it open) |
| `92` | Upsert key column not found |
| `93` | Schema mismatch (use `--strict` or manual CAST) |
| `94` | Extension not available (e.g., postgres_scanner) |
| `95` | Virtual view connection failed |
| `96` | Pivot refresh failed (no matching pivot found) |
| `100` | Table not found |
| `101` | Table already exists (use `--replace`) |
| `102` | Invalid table name |
| `103` | Link connection failed |
| `104` | Refresh timeout (data not loaded in time) |
| `105` | Writeback key column not found |
| `106` | No dirty rows to writeback |
| `107` | Schema drift detected (use `--strict` or `--prune`) |
| `108` | Pivot creation failed (source table empty) |
| `109` | Formula column syntax error |
| `110` | Value violates validation (strict mode) |
| `111` | Validation type not supported |
| `112` | Invalid formula syntax |
| `113` | Dependent validation map not found |
| `114` | Parent cell validation not found |
| `115` | Circular dependency in dependent validation |
| `116` | Validation range is too large |
| `120` | Watcher already active for this file |
| `121` | No active watcher to stop |
| `122` | Watcher PID not found (stale pid file) |
| `123` | Headless mode not supported on this platform |
| `124` | File does not exist |
| `125` | Condition syntax error |
| `126` | Watcher timeout (Excel closed, no re-open) |
| `127` | App.Ready check failed (Excel in bad state) |

---

## Global Flags

These flags work with all commands:

```
--json          # JSON output
--json-errors   # Return errors as JSON instead of stderr
--dry-run       # Preview without executing
--engine <name> # Force engine (xlwings|openpyxl)
--verbose       # Verbose logging
```

**JSON errors example:**
```bash
xlforge cell set report.xlsx "A1" "value" --json-errors
# On success: {"success": true}
# On error: {"success": false, "code": 3, "message": "Sheet not found"}
```

---

## Retry Mechanism

File locks use exponential backoff:
```
Attempt 1: immediate
Attempt 2: 1 second delay
Attempt 3: 2 second delay
```

After 3 attempts, returns exit code 6.

For "Excel busy" (edit mode), use `app wait-idle` before running scripts:
```bash
xlforge app wait-idle report.xlsx --timeout 60
xlforge run script.xlf
```

---

## Implementation Structure

```
xlforge/
├── src/xlforge/
│   ├── __init__.py
│   ├── cli.py              # Click/Typer entry point
│   ├── core.py             # Engine abstraction
│   ├── engines/
│   │   ├── xlwings.py      # xlwings implementation
│   │   ├── openpyxl.py     # openpyxl implementation
│   │   └── duckdb.py       # DuckDB SQL engine
│   ├── batch.py            # Script execution + debug HUD
│   ├── recorder.py          # Macro recording
│   ├── semantic.py          # Index + query
│   ├── fastpath.py          # Direct-to-OpenXML writer
│   └── commands/
│       ├── file.py          # file open, save, close, info, kill
│       ├── sheet.py         # sheet list, create, delete, rename, copy, use
│       ├── cell.py          # cell get, set, formula, clear, copy, bulk
│       ├── format.py        # format cell, range, apply
│       ├── column.py        # column width, auto-fit
│       ├── row.py           # row height
│       ├── data.py          # import csv, export csv
│       ├── table.py         # table create, link, sync-schema, refresh
│       ├── chart.py         # chart create
│       ├── validation.py    # validation create
│       ├── protection.py    # freeze, protect, unprotect
│       ├── app.py           # app visible, calculate, focus, alert, wait-idle
│       ├── checkpoint.py    # checkpoint create, list, restore, delete
│       ├── branch.py        # branch create, list, merge, delete
│       ├── watch.py         # watch start, stop
│       ├── sql.py           # sql query, push, pull, connect
│       └── pro/
│           ├── diff.py      # diff command
│           └── template.py  # template command
```
