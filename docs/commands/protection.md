# Protection Commands

---

## freeze

Freezes panes at a specific cell.

```bash
xlforge freeze <file.xlsx> <sheet!cell>
xlforge freeze <file.xlsx> <sheet> --header
xlforge freeze <file.xlsx> <sheet> --first-col
xlforge freeze <file.xlsx> <sheet> --top-left
```

**Semantic shortcuts:**
| Flag | Effect |
|------|--------|
| `--header` | Freezes first row (freeze at row 2) |
| `--first-col` | Freezes first column (freeze at column B) |
| `--top-left` | Freezes row 1 and column A |

**Examples:**
```bash
xlforge freeze report.xlsx "Data!A2"       # Freezes row 1
xlforge freeze report.xlsx "Data" --header  # Same, shorter
xlforge freeze report.xlsx "Data" --top-left  # Freezes R1 + Col A
```

**Behavior:** Internally activates the target sheet, scrolls to A1, then applies freeze.

---

## unfreeze

Removes freeze panes.

```bash
xlforge unfreeze <file.xlsx> <sheet>
```

---

## protect

Protects a sheet.

```bash
xlforge protect <file.xlsx> <sheet>
xlforge protect <file.xlsx> <sheet> --password <pwd>
```

### Security Options

**Password from environment variable:**
```bash
export XL_PASS="company-secret"
xlforge protect report.xlsx "Summary" --password-env XL_PASS
```

**Password from stdin:**
```bash
echo "Secret123" | xlforge protect report.xlsx "Summary" --password-stdin
```

### Granular Permissions

Allow specific actions while protected:

```bash
xlforge protect <file.xlsx> <sheet> \
    --allow-filters \
    --allow-sorting \
    --allow-formatting-cells \
    --allow-formatting-columns \
    --allow-formatting-rows \
    --allow-insert-columns \
    --allow-insert-rows \
    --allow-delete-columns \
    --allow-delete-rows \
    --allow-select-locked-cells \
    --allow-select-unlocked-cells
```

**Common presets:**
| Use Case | Flags |
|-----------|-------|
| Data entry | `--allow-filters --allow-sorting --allow-select-unlocked-cells` |
| Manager review | `--allow-filters --allow-sorting` |
| Fully locked | (no flags) |

---

## unprotect

Removes protection from a sheet.

```bash
xlforge unprotect <file.xlsx> <sheet>
xlforge unprotect <file.xlsx> <sheet> --password <pwd>
```

---

## protect-workbook

Protects the workbook structure (prevents delete/rename/hide of sheets).

```bash
xlforge protect-workbook <file.xlsx>
xlforge protect-workbook <file.xlsx> --password <pwd>
```

**Options:**
```
--password <pwd>       # Workbook password
--windows             # Also protect windows (window positions/sizes)
```

---

## unprotect-workbook

Removes workbook protection.

```bash
xlforge unprotect-workbook <file.xlsx>
xlforge unprotect-workbook <file.xlsx> --password <pwd>
```

---

## cell unlock

Unlocks specific cells before sheet protection.

```bash
xlforge cell unlock <file.xlsx> <sheet!cell>
xlforge cell unlock <file.xlsx> <sheet!range>
```

**Use case:** Create editable regions in a locked sheet.

```bash
# Unlock only the data entry area
xlforge cell unlock report.xlsx "Summary!B2:D10"

# Lock everything, then unlock specific cells
xlforge protect report.xlsx "Summary"

# Now user can only edit B2:D10
```

---

## cell lock

Locks specific cells (default state, but explicit).

```bash
xlforge cell lock <file.xlsx> <sheet!cell>
```

---

## protection status

Checks protection status of a sheet or workbook.

```bash
xlforge protection status <file.xlsx> <sheet>
xlforge protection status <file.xlsx> --workbook --json
```

**JSON output:**
```json
{
  "sheet": "Summary",
  "is_protected": true,
  "has_password": true,
  "permissions": {
    "allow_filters": true,
    "allow_sorting": true,
    "allow_formatting": false,
    "allow_insert_columns": false
  }
}
```

---

## Sheet Visibility

### sheet hide

Hides a sheet.

```bash
xlforge sheet hide <file.xlsx> <sheet>
xlforge sheet hide <file.xlsx> <sheet> --very-hidden
```

**Options:**
```
--very-hidden    # Prevents user from unhiding via right-click menu
```

**Why `--very-hidden`?** Standard hidden sheets can be unhidden via Excel's UI. `xlSheetVeryHidden` can only be unhidden by macros/CLI.

```bash
xlforge sheet hide report.xlsx "BackendData"
xlforge sheet hide report.xlsx "RawData" --very-hidden
```

### sheet unhide

Unhides a sheet.

```bash
xlforge sheet unhide <file.xlsx> <sheet>
```

---

## visibility list

Lists all sheets with visibility status.

```bash
xlforge sheet visibility <file.xlsx>
xlforge sheet visibility <file.xlsx> --json
```

**JSON output:**
```json
{
  "sheets": [
    {"name": "Summary", "visible": true},
    {"name": "Data", "visible": true},
    {"name": "BackendData", "visible": "very-hidden"},
    {"name": "RawData", "visible": false}
  ]
}
```

---

## watermark

Adds a watermark to a sheet (via header image).

```bash
xlforge watermark <file.xlsx> <sheet> <text>
xlforge watermark <file.xlsx> <sheet> <text> --color "#FF0000"
```

**Example:**
```bash
xlforge watermark report.xlsx "Summary" "CONFIDENTIAL" --color "#FF0000"
```

---

## Complete Examples

### Locked data entry form
```bash
# Unlock only input cells
xlforge cell unlock report.xlsx "Entry!B2:D100"

# Lock headers and formulas
xlforge protect report.xlsx "Entry" \
    --allow-select-unlocked-cells \
    --allow-filters
```

### Secure report delivery
```bash
# Hide raw data
xlforge sheet hide report.xlsx "RawData" --very-hidden

# Protect structure
xlforge protect-workbook report.xlsx --password-env WB_PASS

# Protect sensitive sheet
xlforge protect report.xlsx "Summary" --password-env SH_PASS \
    --allow-filters --allow-sorting

# Add watermark
xlforge watermark report.xlsx "Summary" "DRAFT"
```

### Reset for re-editing
```bash
# Remove all protection
xlforge unprotect-workbook report.xlsx --password-env WB_PASS
xlforge unprotect report.xlsx --password-env SH_PASS
xlforge sheet unhide report.xlsx "RawData"
xlforge unfreeze report.xlsx "Summary"
```

---

## Error Codes

| Code | Meaning |
|------|---------|
| `70` | Sheet is protected |
| `71` | Password required |
| `72` | Invalid password |
| `73` | Workbook is protected |
| `74` | Cannot unhide very-hidden sheet |
| `75` | Cell is locked |
| `76` | Invalid protection option |
