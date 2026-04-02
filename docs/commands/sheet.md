# Sheet Commands

Manage workbook structure: sheets, ordering, visibility, and organization.

---

## sheet list

Lists all sheets in the workbook.

```bash
xlforge sheet list <file.xlsx>
xlforge sheet list <file.xlsx> --json
xlforge sheet list <file.xlsx> --match "Sales*"
```

### Pattern Matching

Filter sheets by name pattern:

```bash
xlforge sheet list report.xlsx --match "Q1*"     # All Q1 sheets
xlforge sheet list report.xlsx --match "*Data*"  # Contains "Data"
xlforge sheet list report.xlsx --match "Summary" # Exact match
```

### Enhanced JSON Output

```json
{
  "sheets": [
    {"name": "Summary", "index": 1, "visibility": "visible", "tab_color": null, "is_active": true},
    {"name": "Data", "index": 2, "visibility": "visible", "tab_color": "#00B050", "is_active": false},
    {"name": "Analysis", "index": 3, "visibility": "hidden", "tab_color": null, "is_active": false}
  ]
}
```

**Fields:**
| Field | Type | Description |
|-------|------|-------------|
| `name` | string | Sheet name |
| `index` | number | Position (1-based) |
| `visibility` | string | `visible`, `hidden`, `very-hidden` |
| `tab_color` | string | Hex color or null |
| `is_active` | boolean | Currently selected sheet |

---

## sheet create

Creates a new sheet.

```bash
xlforge sheet create <file.xlsx> <name>
xlforge sheet create <file.xlsx> <name> --before <other-sheet>
xlforge sheet create <file.xlsx> <name> --after <other-sheet>
xlforge sheet create <file.xlsx> <name> --if-not-exists
```

### Name Validation

Sheet names are validated *before* calling Excel:

| Rule | Limit |
|------|-------|
| Max length | 31 characters |
| Illegal chars | `\ / ? * [ ] :` |

```bash
# Auto-truncate long names
xlforge sheet create report.xlsx "Sales_Report_For_Q1_2026_Final_Final_v2"
# Creates: "Sales_Report_For_Q1_2026_Fi" (31 chars)

# Explicit error for illegal chars
xlforge sheet create report.xlsx "Sales/Data"
# Error 5: Invalid sheet name (contains illegal character '/')
```

### Idempotency Flag

Use `--if-not-exists` to prevent errors when running scripts multiple times:

```bash
xlforge sheet create report.xlsx "Data" --if-not-exists
# Does nothing if "Data" already exists (exit 0)
```

---

## sheet delete

Deletes a sheet.

```bash
xlforge sheet delete <file.xlsx> <name>
xlforge sheet delete <file.xlsx> <name> --if-exists
xlforge sheet delete <file.xlsx> --all --keep "Summary"
```

### Idempotency Flag

Use `--if-exists` to prevent errors when running scripts multiple times:

```bash
xlforge sheet delete report.xlsx "TempData" --if-exists
# Does nothing if "TempData" doesn't exist (exit 0)
```

### Bulk Delete

Delete multiple sheets with `--all` and `--keep`:

```bash
# Delete all sheets except Summary
xlforge sheet delete report.xlsx --all --keep "Summary"

# Delete all sheets except multiple
xlforge sheet delete report.xlsx --all --keep "Summary" --keep "Data"

# Delete all sheets (dangerous!)
xlforge sheet delete report.xlsx --all
# Error 77: Cannot delete last sheet (workbook must have at least one sheet)
```

**Note:** Cannot delete the last sheet in a workbook.

---

## sheet rename

Renames a sheet.

```bash
xlforge sheet rename <file.xlsx> <old-name> <new-name>
xlforge sheet rename <file.xlsx> <old-name> <new-name> --if-exists
```

**Name validation applies** (31 chars, illegal chars).

```bash
xlforge sheet rename report.xlsx "Sheet1" "Summary"
# Error 5: Invalid sheet name (contains illegal character ':')
```

---

## sheet copy

Copies a sheet to a new name.

```bash
xlforge sheet copy <file.xlsx> <from> <to>
xlforge sheet copy <file.xlsx> <from> <to> --with-data
```

### Options

```bash
--with-data    # Copy data and formatting (default)
--format-only # Copy formatting only, leave cells empty
```

---

## sheet use

Sets the active sheet for context-based commands AND activates it in the Excel UI.

```bash
xlforge sheet use <file.xlsx> <name>
```

**Behavior:**
1. Sets CLI default context (subsequent commands don't need `--sheet`)
2. Activates the sheet in Excel (visible to user)

```bash
xlforge sheet use report.xlsx "Data"
xlforge cell set "A1" "Value"  # Uses Data sheet
xlforge sheet use report.xlsx "Summary"
xlforge cell get "B7"           # Now uses Summary sheet
```

---

## sheet hide

Hides a sheet from the tab bar.

```bash
xlforge sheet hide <file.xlsx> <name>
xlforge sheet hide <file.xlsx> <name> --very-hidden
```

### Visibility Levels

| Level | User Can Unhide? | Use Case |
|-------|------------------|----------|
| `hidden` | Yes (right-click menu) | Temporary hiding |
| `very-hidden` | No (only macro/CLI) | Backend sheets, calculation sheets |

**Example:**
```bash
xlforge sheet hide report.xlsx "Calculation_Sheet"
xlforge sheet hide report.xlsx "RawData" --very-hidden
# RawData cannot be unhidden via Excel UI
```

---

## sheet unhide

Makes a hidden sheet visible.

```bash
xlforge sheet unhide <file.xlsx> <name>
```

**Note:** Cannot unhide `very-hidden` sheets. Use `sheet hide` without `--very-hidden` first:

```bash
xlforge sheet hide report.xlsx "RawData"  # Remove very-hidden
xlforge sheet unhide report.xlsx "RawData" # Now it appears
```

---

## sheet move

Reorders a sheet within the workbook.

```bash
xlforge sheet move <file.xlsx> <name> --to-start
xlforge sheet move <file.xlsx> <name> --to-end
xlforge sheet move <file.xlsx> <name> --index 3
xlforge sheet move <file.xlsx> <name> --before <other-sheet>
xlforge sheet move <file.xlsx> <name> --after <other-sheet>
```

**Examples:**
```bash
xlforge sheet move report.xlsx "Summary" --to-start   # Move to first position
xlforge sheet move report.xlsx "Appendix" --to-end     # Move to last position
xlforge sheet move report.xlsx "Data" --index 2       # Move to position 2
xlforge sheet move report.xlsx "Data" --before "Summary"  # Before Summary
```

---

## sheet color

Sets the tab color of a sheet.

```bash
xlforge sheet color <file.xlsx> <name> <color>
xlforge sheet color <file.xlsx> <name> <color> --json
```

### Color Formats

```bash
xlforge sheet color report.xlsx "Data" "green"      # Named color
xlforge sheet color report.xlsx "Data" "#00B050"    # Hex color
xlforge sheet color report.xlsx "Data" "auto"      # Excel auto (no color)
```

### Named Colors

`red`, `green`, `blue`, `yellow`, `orange`, `purple`, `cyan`, `white`, `black`, `gray`

### JSON Output

```json
{
  "sheet": "Data",
  "tab_color": "#00B050",
  "previous_color": null
}
```

**Use case:** Color-coding sheets (Green = Inputs, Yellow = Processing, Red = Outputs)

---

## sheet isolate

Exports a sheet as a standalone workbook.

```bash
xlforge sheet isolate <file.xlsx> <name>
xlforge sheet isolate <file.xlsx> <name> --output <new-file.xlsx>
```

### Link Handling

By default, `isolate` breaks links to prevent `#REF!` errors:

```bash
--break-links     # Paste values, remove external references (default)
--keep-links     # Preserve formulas that reference other sheets (may cause #REF!)
--values-only    # Export data only, no formulas
```

**Example:**
```bash
xlforge sheet isolate report.xlsx "Summary" --output "summary_only.xlsx"
# Creates standalone workbook with Summary's data intact
```

---

## sheet clear

Deletes all cell values and formatting from a sheet, but keeps the sheet itself.

```bash
xlforge sheet clear <file.xlsx> <name>
xlforge sheet clear <file.xlsx> <name> --confirm
```

**Warning:** This is destructive. Use `--confirm` to acknowledge:

```bash
xlforge sheet clear report.xlsx "Data" --confirm
# Clears all data and formatting, sheet structure remains
```

**Use case:** Resetting a sheet for a new data run while preserving table definitions, named ranges, and sheet-level settings.

---

## Complete Examples

### Creating an organized workbook structure
```bash
# Create sheets with organization
xlforge sheet create report.xlsx "Summary" --if-not-exists
xlforge sheet create report.xlsx "Data" --after "Summary" --if-not-exists
xlforge sheet create report.xlsx "Analysis" --after "Data" --if-not-exists

# Color-code by function
xlforge sheet color report.xlsx "Summary" "green"    # Inputs
xlforge sheet color report.xlsx "Data" "yellow"     # Processing
xlforge sheet color report.xlsx "Analysis" "red"    # Outputs

# Hide backend sheets
xlforge sheet hide report.xlsx "Calculation_Sheet" --very-hidden
```

### Clean slate for new data run
```bash
# Backup current data
xlforge checkpoint create report.xlsx --tag "pre-reset"

# Reset data sheets
xlforge sheet clear report.xlsx "Data" --confirm
xlforge sheet clear report.xlsx "RawData" --confirm

# Import new data
xlforge import csv report.xlsx new_data.csv --sheet Data --cell A1
```

### Agent discovery workflow
```bash
# List all sheets with metadata
xlforge sheet list report.xlsx --json

# Find specific sheets
xlforge sheet list report.xlsx --match "Sales*"

# Activate sheet for context
xlforge sheet use report.xlsx "Sales_Q1_2026"
```

### Export a sheet for sharing
```bash
# Isolate Summary sheet as standalone file
xlforge sheet isolate report.xlsx "Summary" --output "summary_export.xlsx" --break-links

# Verify no broken references
xlforge file check summary_export.xlsx
```

---

## Error Codes

| Code | Meaning |
|------|---------|
| `3` | Sheet not found |
| `5` | Invalid sheet name (too long, illegal chars) |
| `33` | Sheet is hidden (operation requires visible sheet) |
| `74` | Cannot unhide very-hidden sheet (use `sheet hide` first) |
| `77` | Cannot delete last sheet (workbook must have at least one) |
| `78` | Circular sheet reference in move |
| `79` | Cannot move sheet that doesn't exist |
