# Column & Row Commands

---

## column width

Sets column width.

```bash
xlforge column width <file.xlsx> <sheet!col> <value>
xlforge column width <file.xlsx> <sheet!col> <value> --unit <unit>
```

**Arguments:**
- `sheet!col`: Column reference (e.g., `Data!A:A` or `Data!A:C`)
- Supports multi-column ranges: `Data!A:Z` (single COM call)

**Options:**
```
--unit <unit>     # px (pixels), pt (points), or excel (character units)
```

**Examples:**
```bash
xlforge column width report.xlsx "Data!A:A" 100 --unit px
xlforge column width report.xlsx "Data!B:B" 15       # Excel character units
xlforge column width report.xlsx "Data!A:Z" 50     # Entire range, one call
```

**Unit conversion:**
- `px`: Pixels (converted based on DPI)
- `pt`: Points (1/72 inch)
- `excel`: Default Excel character width (default)

---

## column auto-fit

Auto-fits column width to content.

```bash
xlforge column auto-fit <file.xlsx> <sheet!col>
xlforge column auto-fit <file.xlsx> <sheet!col> --max <px>
```

**Options:**
```
--max <px>    # Cap width at this value
```

**Examples:**
```bash
# Standard auto-fit
xlforge column auto-fit report.xlsx "Data!A:A"

# With max-width (prevents column sprawl from long text)
xlforge column auto-fit report.xlsx "Data!A:Z" --max 250
```

**Why `--max`?** Standard auto-fit can make columns unreasonably wide for paragraphs. Cap prevents ugliness.

---

## column best-fit

Data-aware sizing that analyzes content type.

```bash
xlforge column best-fit <file.xlsx> <sheet!col>
```

**Behavior by type:**
- **Dates**: Fixed width for `YYYY-MM-DD` format
- **Currency**: Buffer for `€` sign and decimals
- **Numbers**: Padding for comma separators
- **Text**: Longest string + buffer

---

## column hide / column unhide

Hides or unhides columns.

```bash
xlforge column hide <file.xlsx> <sheet!col>
xlforge column unhide <file.xlsx> <sheet!col>
```

**Example:**
```bash
xlforge column hide report.xlsx "Data!B:C"    # Hide columns B-C
xlforge column unhide report.xlsx "Data!B:B"  # Unhide column B
```

---

## column visibility list

Lists hidden columns in a sheet.

```bash
xlforge column visibility <file.xlsx> <sheet>
xlforge column visibility <file.xlsx> <sheet> --json
```

**JSON output:**
```json
{
  "sheet": "Data",
  "hidden": ["B", "D", "F"]
}
```

---

## row height

Sets row height.

```bash
xlforge row height <file.xlsx> <sheet!row> <value>
xlforge row height <file.xlsx> <sheet!row> <value> --unit <unit>
```

**Arguments:**
- `sheet!row`: Row reference (e.g., `Data!1:1` or `Data!1:10`)
- Supports multi-row ranges: `Data!1:100` (single COM call)

**Options:**
```
--unit <unit>     # px (pixels) or pt (points, default)
```

**Examples:**
```bash
xlforge row height report.xlsx "Data!1:1" 30 --unit pt
xlforge row height report.xlsx "Data!2:100" 15  # pt default
```

---

## row auto-fit

Auto-fits row height to content.

```bash
xlforge row auto-fit <file.xlsx> <sheet!row>
```

---

## row hide / row unhide

Hides or unhides rows.

```bash
xlforge row hide <file.xlsx> <sheet!row>
xlforge row unhide <file.xlsx> <sheet!row>
```

**Example:**
```bash
xlforge row hide report.xlsx "Data!10:20"   # Hide rows 10-20
xlforge row unhide report.xlsx "Data!10:10"  # Unhide row 10
```

---

## row stripe

Applies alternating row colors for readability.

```bash
xlforge row stripe <file.xlsx> <sheet!range> [options]
```

**Options:**
```
--color <color>     # Stripe color (default: light gray)
--every <n>         # Stripe every n rows (default: 1)
```

**Examples:**
```bash
xlforge row stripe report.xlsx "Data!2:100"
xlforge row stripe report.xlsx "Data!2:100" --color "217,225,242" --every 2
```

---

## Semantic Presets

### size preset

Quick size presets instead of pixel values.

```bash
xlforge column width <file.xlsx> <sheet!col> --size <preset>
xlforge row height <file.xlsx> <sheet!row> --size <preset>
```

**Column presets:**
| Preset | Excel Units | Use Case |
|--------|-------------|----------|
| `tiny` | 5 | ID columns |
| `small` | 10 | Short text |
| `medium` | 15 | Standard text |
| `large` | 25 | Long text |
| `xlarge` | 40 | Paragraphs |

**Row presets:**
| Preset | Points | Use Case |
|--------|--------|----------|
| `compact` | 12 | Data rows |
| `normal` | 15 | Standard |
| `header` | 30 | Header rows |
| `tall` | 45 | Multi-line |

**Examples:**
```bash
xlforge column width report.xlsx "Data!A:A" --size tiny
xlforge row height report.xlsx "Data!1:1" --size header
```

---

## Complete Examples

### Header Formatting
```bash
# Set up professional headers
xlforge row height report.xlsx "Data!1:1" --size header
xlforge column width report.xlsx "Data!A:Z" --size medium
xlforge row stripe report.xlsx "Data!2:100" --every 2
```

### Multi-column Layout
```bash
# Set up a data entry form
xlforge column width report.xlsx "Data!A:A" --size large  # Name
xlforge column width report.xlsx "Data!B:B" --size medium  # Email
xlforge column width report.xlsx "Data!C:C" --size small   # Phone
```

### Auto-fit with Guardrails
```bash
xlforge column auto-fit report.xlsx "Data!A:Z" --max 250
```

### Hide/Show Workflow
```bash
xlforge column hide report.xlsx "Data!B:D"   # Group columns
# ... user work ...
xlforge column unhide report.xlsx "Data!B:D"  # Reveal
```

---

## Error Codes

| Code | Meaning |
|------|---------|
| `30` | Column not found |
| `31` | Row not found |
| `32` | Invalid unit (use px, pt, or excel) |
| `33` | Column/row hidden (unhide first) |
