# Checkpoint Commands (Git for Excel)

Checkpoints are lightweight snapshots that enable safe experimentation and rollback.

---

## checkpoint create

Creates a snapshot of the current state.

```bash
xlforge checkpoint create <file.xlsx> <name>
xlforge checkpoint create <file.xlsx> <name> --message <message>
xlforge checkpoint create <file.xlsx> <name> --max <n>
```

**Options:**
```
--message <msg>      # Description of what changed
--max <n>            # Retention limit (auto-prune older checkpoints)
```

**What is stored:**
- Cell values and formulas
- Basic formatting
- Sheet structure
- Metadata (version, timestamp, SHA-256)

**Location:** `.xlforge/checkpoints/<name>.zip`

**Manifest contents:**
```json
{
  "name": "pre-agent",
  "created": "2026-03-31T10:00:00Z",
  "message": "Before AI modifications",
  "version": "1.0.0",
  "sha256": "a1b2c3d4e5f6...",
  "file_size": 1048576
}
```

---

## checkpoint list

Lists all checkpoints for a file.

```bash
xlforge checkpoint list <file.xlsx>
xlforge checkpoint list <file.xlsx> --json
xlforge checkpoint list <file.xlsx> --tags    # Show tags too
```

**JSON output:**
```json
{
  "file": "report.xlsx",
  "checkpoints": [
    {
      "name": "initial-load",
      "created": "2026-03-31T10:00:00Z",
      "message": "Initial data load",
      "tags": ["v1.0-baseline"]
    },
    {
      "name": "pre-agent",
      "created": "2026-03-31T11:30:00Z",
      "message": "Before AI modifications",
      "tags": []
    }
  ]
}
```

---

## checkpoint restore

Restores the file to a previous checkpoint. **Atomic operation.**

```bash
xlforge checkpoint restore <file.xlsx> <name>
xlforge checkpoint restore <file.xlsx> <name> --force   # Skip confirmation
xlforge checkpoint restore <file.xlsx> <name> --dry-run  # Preview
```

**Atomic swap pattern:**
1. Close Excel COM instance
2. Unzip checkpoint to `.tmp`
3. Rename original to `.old`
4. Rename `.tmp` to original
5. Delete `.old` on success

**Safety:** If restore fails, `.old` is preserved for manual recovery.

---

## checkpoint diff

Shows differences between two checkpoints.

```bash
xlforge checkpoint diff <file.xlsx> <name1> <name2>
xlforge checkpoint diff <file.xlsx> <name1> <name2> --json
```

**JSON output:**
```json
{
  "from": "pre-agent",
  "to": "post-agent",
  "changes": [
    {"cell": "Summary!A1", "old": "Sales", "new": "Revenue"},
    {"cell": "Summary!B5", "old": 100, "new": 150},
    {"cell": "Summary!C3", "old": null, "new": "=SUM(Data!B:B)"}
  ]
}
```

**Use case:** Verify agent changed exactly what was intended.

---

## checkpoint tag

Tags a checkpoint for easy reference.

```bash
xlforge checkpoint tag <file.xlsx> <name> <tag>
xlforge checkpoint tag <file.xlsx> <name> --add <tag>
xlforge checkpoint tag <file.xlsx> <name> --remove <tag>
```

**Examples:**
```bash
xlforge checkpoint tag report.xlsx "initial-load" "v1.0-baseline"
xlforge checkpoint tag report.xlsx "pre-agent" --add "stable"
xlforge checkpoint tag report.xlsx "pre-agent" --remove "unstable"
```

---

## checkpoint prune

Deletes old checkpoints based on retention policy.

```bash
xlforge checkpoint prune <file.xlsx>
xlforge checkpoint prune <file.xlsx> --keep <n>      # Keep last n
xlforge checkpoint prune <file.xlsx> --before <date>  # Delete before date
xlforge checkpoint prune <file.xlsx> --all            # Delete all
```

---

## checkpoint delete

Deletes a specific checkpoint.

```bash
xlforge checkpoint delete <file.xlsx> <name>
```

---

## checkpoint commit

Saves checkpoint and commits to Git.

```bash
xlforge checkpoint commit <file.xlsx> <name> -m <message>
```

**What it does:**
1. Creates checkpoint with name
2. Cleans temporary metadata
3. Runs `git add <file.xlsx>`
4. Runs `git commit -m "<message>"`

**Requires:** File must be tracked by Git.

---

## Branch Commands

Branches are "shadow copies" of sheets for safe experimentation.

### branch create

Creates a hidden copy of a sheet.

```bash
xlforge branch create <file.xlsx> <sheet> <branch-name>
xlforge branch create <file.xlsx> <sheet> <branch-name> --hidden
```

**Options:**
```
--hidden     # Keep branch hidden from user (for AI testing)
```

**Example:**
```bash
# Create hidden branch for AI work
xlforge branch create report.xlsx "Summary" "AI_Update" --hidden

# Create visible branch for human work
xlforge branch create report.xlsx "Summary" "Dev"
```

**Hidden vs Visible:**
- `--hidden`: User sees original sheet, AI manipulates hidden copy
- Default: Creates visible "Summary_Copy" sheet

---

### branch list

Lists all branches for a file.

```bash
xlforge branch list <file.xlsx>
xlforge branch list <file.xlsx> --json
```

**JSON output:**
```json
{
  "file": "report.xlsx",
  "branches": [
    {
      "name": "AI_Update",
      "source": "Summary",
      "hidden": true,
      "created": "2026-03-31T10:00:00Z"
    },
    {
      "name": "Dev",
      "source": "Summary",
      "hidden": false,
      "created": "2026-03-31T09:00:00Z"
    }
  ]
}
```

---

### branch merge

Merges a branch back to its source sheet.

```bash
xlforge branch merge <file.xlsx> <branch-name>
xlforge branch merge <file.xlsx> <branch-name> <target-sheet>
```

**Options:**
```
--strategy <strategy>   # overwrite, fill-empty, formulas-only
--delete                 # Delete branch after merge
```

**Merge strategies:**

| Strategy | Behavior |
|----------|----------|
| `overwrite` | Branch version wins completely |
| `fill-empty` | Only fill cells null in original |
| `formulas-only` | Bring back formulas, keep other data |

**Examples:**
```bash
# Simple merge (overwrite)
xlforge branch merge report.xlsx "AI_Update"

# Fill only empty cells
xlforge branch merge report.xlsx "AI_Update" --strategy fill-empty

# Merge and delete branch
xlforge branch merge report.xlsx "AI_Update" --delete
```

---

### branch delete

Deletes a branch.

```bash
xlforge branch delete <file.xlsx> <branch-name>
```

---

## Auto-Checkpoint Feature

### Global --checkpoint flag

Automatically creates a checkpoint before running a script.

```bash
xlforge run <script.xlf> --checkpoint
xlforge run <script.xlf> --checkpoint --checkpoint-name <name>
```

**Behavior:**
1. Creates checkpoint named `auto-<timestamp>` before execution
2. If script fails (exit code ≠ 0), offers auto-restore
3. User can run `xlforge checkpoint restore <name>` manually

**Example workflow:**
```bash
xlforge run update_script.xlf --checkpoint
# Script fails at line 15

# Offer restore
xlforge checkpoint restore report.xlsx "auto-20260331-103045"
```

---

## Complete Workflow Examples

### Agent Safe Workflow
```bash
# 1. Create a safe space to work
xlforge branch create report.xlsx "Summary" "AI_Update" --hidden

# 2. Set context to use the branch
xlforge use report.xlsx --sheet "AI_Update"

# 3. Perform risky operations
xlforge cell formula "B2" "=SUM(Data!C:C)"
xlforge format cell "B2" --bold

# 4. Compare with original
xlforge diff report.xlsx "Summary" "AI_Update"

# 5. Looks good? Merge back
xlforge branch merge report.xlsx "AI_Update" --delete
```

### Pre-Flight Checkpoint Workflow
```bash
# Run dangerous script with safety net
xlforge run risky_updates.xlf --checkpoint

# Script succeeds - checkpoint preserved
# Script fails - restore with one command
xlforge checkpoint restore report.xlsx "auto-20260331-104500" --force
```

### Stable Release Workflow
```bash
# Create baseline checkpoint
xlforge checkpoint create report.xlsx "release-v1" --message "v1.0 release"
xlforge checkpoint tag report.xlsx "release-v1" "stable,v1.0"

# Make changes...

# If problems arise, restore stable
xlforge checkpoint restore report.xlsx "release-v1"
```

---

## Storage Architecture

### Location
```
.xlforge/
└── checkpoints/
    └── report.xlsx/
        ├── manifest.json      # Index of all checkpoints
        ├── pre-agent.zip     # Checkpoint data
        ├── initial-load.zip
        └── ...

```

### Retention
- Default retention: unlimited (use `checkpoint prune`)
- `--max 10`: Keep last 10 checkpoints
- Auto-cleanup of `.old` recovery files after 7 days

### Deduplication (v1.1)
Future versions may store only XML deltas to reduce storage bloat.

---

## Error Codes

| Code | Meaning |
|------|---------|
| `20` | Checkpoint not found |
| `21` | Checkpoint restore failed (file corrupted) |
| `22` | Branch not found |
| `23` | Branch merge conflict |
| `24` | Cannot delete active branch |
