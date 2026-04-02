# Format Commands

---

## format cell

Applies formatting to a single cell (or range).

```bash
xlforge format cell <file.xlsx> <sheet!cell> [options]
xlforge format <file.xlsx> <sheet!cell> [options]    # Shorthand
```

**Performance:** Internally uses range operations. `format cell A1` = `format range A1:A1`.

---

## format range

Applies formatting to a range of cells.

```bash
xlforge format range <file.xlsx> <sheet!range> [options]
```

---

## Tailwind-Style Shorthand

The fastest way to format. Space-separated, CSS-like syntax.

```bash
--style "bold size-14 text-#2C4874 bg-#D9E1F2 align-center"
```

**Shorthand tokens:**
| Token | Meaning |
|-------|---------|
| `bold` | Font weight bold |
| `italic` | Font style italic |
| `underline` | Underline |
| `size-N` | Font size N points |
| `text-#RRGGBB` | Text color (hex) |
| `bg-#RRGGBB` | Background color (hex) |
| `font-Name` | Font name |
| `align-left\|center\|right` | Horizontal alignment |
| `valign-top\|center\|bottom` | Vertical alignment |
| `wrap` | Text wrap |
| `merge` | Merge cells |

**Examples:**
```bash
xlforge format "Summary!A1" --style "bold size-16 text-#2C4874"
xlforge format "Data!A1:E10" --style "bg-#D9E1F2 border-all"
```

---

## Color Format

Supports both formats:

```bash
--color "44,72,116"      # RGB integers (legacy)
--color "#2C4874"         # Hex (preferred)
```

**Hex is recommended** for agents and modern tooling.

---

## Number Formatting

### number-format

Sets the number format.

```bash
--number-format <format>
```

**Preset shortcuts:**
| Preset | Excel Format |
|--------|--------------|
| `currency` | Local currency (e.g., `$#,##0.00`) |
| `percent` | `0.00%` |
| `date` | `YYYY-MM-DD` |
| `datetime` | `YYYY-MM-DD HH:MM:SS` |
| `time` | `HH:MM:SS` |
| `comma` | `#,##0.00` |
| `integer` | `0` |

**Custom formats:**
```bash
--number-format "€#,##0.00"
--number-format "0.00%"
--number-format "YYYY-MM-DD"
```

---

## Border Options

Simplified border syntax.

```bash
--border <type>
```

| Type | Effect |
|------|--------|
| `all` | Outside + inside borders |
| `box` | Outside border only |
| `bottom` | Bottom border only |
| `top` | Top border only |
| `left` | Left border only |
| `right` | Right border only |
| `inside` | Inside borders only |
| `none` | Remove all borders |

**With style:**
```bash
--border thin|medium|thick
--border bottom:medium   # Specific border + style
```

---

## Named Styles

Use Excel's built-in styles for clean, reusable formatting.

```bash
--named-style <style-name>
```

**Available styles:**
```bash
xlforge format "Summary!A1" --named-style "Heading 1"
xlforge format "Summary!B2" --named-style "Currency"
xlforge format "Data!A1" --named-style "Header"
```

**Benefits:**
- Uses Excel's native style system
- Prevents "style bloat"
- Theme changes propagate automatically

---

## format apply

Applies a predefined table style.

```bash
xlforge format apply <file.xlsx> <sheet!range> <style>
```

**Styles:** `zebra`, `zebra-striped`, `bordered`, `plain`, `grid`

```bash
xlforge format apply report.xlsx "Data!A1:E10" zebra
```

---

## format table

Converts a range into a Native Excel Table (ListObject).

```bash
xlforge format table <file.xlsx> <sheet!range> --name <table-name> [options]
```

**Options:**
```
--name <table-name>     # Required: Excel table name
--style <style>         # Table style (default: plain)
--headers              # First row is headers
--no-headers           # No header row
```

**Styles:**
```bash
xlforge format table report.xlsx "Data!A1:E100" \
    --name "SalesData" \
    --style zebra \
    --headers
```

**Benefits:**
- Auto-filters
- Banded rows
- Structured references: `=SUM(SalesData[Revenue])`
- Pivot Table friendly

---

## format condition

Conditional formatting for dynamic visual rules.

```bash
xlforge format condition <file.xlsx> <sheet!range> --rule <rule> --style <style>
```

### Rule types

**Data bars:**
```bash
xlforge format condition report.xlsx "Data!B2:B100" \
    --type data-bar \
    --color "#00FF00"
```

**Color scale:**
```bash
xlforge format condition report.xlsx "Data!C2:C100" \
    --type color-scale \
    --min "#FF0000" \
    --max "#00FF00"
```

**Icon set:**
```bash
xlforge format condition report.xlsx "Data!D2:D100" \
    --type icon-set \
    --icons traffic-light
```

**Formula-based:**
```bash
xlforge format condition report.xlsx "Data!A2:A100" \
    --type formula \
    --formula "=A2>100" \
    --style "bold text-#FF0000"
```

### Rule options

| Option | Description |
|--------|-------------|
| `--type` | data-bar, color-scale, icon-set, formula |
| `--rule` | greater-than, less-than, between, equal, contains |
| `--style` | Style string for matching cells |
| `--color` | Bar/icon color |
| `--min`, `--max` | Scale range |

---

## format copy-to

Format painter - copies formatting from one cell/range to another.

```bash
xlforge format copy-to <file.xlsx> <source> <destination>
```

**Examples:**
```bash
# Copy A1 format to B2:B10
xlforge format copy-to report.xlsx "Sheet1!A1" "Sheet1!B2:B10"

# Copy entire range format
xlforge format copy-to report.xlsx "Summary!A1:F1" "Summary!A2:F10"
```

---

## Complete Examples

### Professional header
```bash
xlforge format "Data!A1:F1" --style "bold size-12 text-white bg-#2C4874 align-center"
```

### Currency formatting
```bash
xlforge format "Data!B2:B100" --number-format currency --align right
```

### Zebra table with headers
```bash
xlforge format table report.xlsx "Data!A1:E100" \
    --name "SalesData" \
    --style zebra-striped \
    --headers
```

### Conditional data bars
```bash
xlforge format condition report.xlsx "Data!B2:B100" \
    --type data-bar \
    --color "#00B050"
```

### Copy manual format to new range
```bash
xlforge format copy-to report.xlsx "Template!B2" "Report!B2:B50"
```

---

## Error Codes

| Code | Meaning |
|------|---------|
| `60` | Invalid style string |
| `61` | Invalid number format |
| `62` | Named style not found |
| `63` | Conditional format not supported |
| `64` | Range too complex for conditional format |
