# Chart Commands

## chart create

Creates a chart from CSV data or a range.

```bash
xlforge chart create <file.xlsx> <anchor-cell> <source> [options]
```

**Arguments:**
- `anchor-cell`: Top-left position (e.g., `Analysis!A1`)
- `source`: CSV file path OR range reference (e.g., `Data!A1:B10`)

**Example with CSV:**
```bash
xlforge chart create report.xlsx "Analysis!A1" sales.csv \
    --type bar \
    --x "Month" \
    --y "Sales,Forecast" \
    --name "SalesChart"
```

**Example with existing range:**
```bash
xlforge chart create report.xlsx "Analysis!A1" "Data!A1:F20" \
    --type line \
    --x "Month" \
    --y "Revenue"
```

---

## Core Options

### type

Chart type.

```bash
--type <type>
```

Types: `bar`, `bar-grouped`, `bar-stacked`, `column`, `column-grouped`, `column-stacked`, `line`, `line-marker`, `pie`, `doughnut`, `area`, `scatter`, `bubble`, `radar`

### name

Chart name for idempotency. If a chart with this name exists, **updates** its data source instead of creating a duplicate.

```bash
--name <chart-name>
```

**Why?** Prevents chart stacking when running scripts multiple times.

```bash
xlforge chart create report.xlsx "A1" sales.csv \
    --type bar \
    --name "RegionalSales" \
    --replace   # Update if exists, create if not
```

---

## Data Source Options

### hidden-data

Creates a hidden sheet to store CSV data, preventing broken links if the CSV is deleted.

```bash
--hidden-data <sheet-prefix>
```

```bash
xlforge chart create report.xlsx "A1" sales.csv \
    --hidden-data "_xlf_data"
# Creates hidden sheet "_xlf_data_sales" with CSV content
# Chart links to this hidden sheet
```

**Without `--hidden-data`:** Chart links directly to the CSV path (breaks if CSV moves).

### x, y

Column mappings for chart data.

```bash
--x <column>           # X-axis column (header name)
--y <columns>          # Y-axis column(s), comma-separated
```

**Header-based (default):**
```bash
--x "Date" --y "Sales,Cost,Profit"
```

**Index-based (for CSVs without headers):**
```bash
--x 0 --y 1,2,3
```

**Auto-detection:** If CSV has headers and `--x` is omitted, uses first column. If `--y` is omitted and CSV has 2+ columns, second column is used.

---

## Sizing Options

### width, height

Chart dimensions in pixels.

```bash
--width <px>
--height <px>
```

Default: `--width 600 --height 400`

### to-cell

Sets bottom-right corner of chart using a cell reference. More intuitive than pixels.

```bash
--to-cell <cell>
```

```bash
xlforge chart create report.xlsx "Analysis!A1" data.csv \
    --type bar \
    --to-cell "Analysis!G20"   # Chart spans A1:G20
```

**Why?** Snaps chart to Excel grid instead of floating at arbitrary pixels.

---

## Style Options

### style

Excel chart style number (1-48) or preset name.

```bash
--style <style>
```

**Style numbers:**
- 1-8: Simple bar/column
- 9-16: Line charts
- 17-24: Pie/doughnut
- 25-32: Area charts
- 33-40: Scatter/bubble
- 41-48: Modern styles

**Preset names (more memorable):**
```bash
--style modern      # Style 46
--style minimal    # Style 42
--style colorful   # Style 10
```

### title

Chart title.

```bash
--title <text>
```

### legend

Show or hide legend.

```bash
--legend <auto|true|false>
```

**Default:** `auto` (shown if multiple Y columns).

---

## Update & Replace

### replace

If a chart with `--name` exists, update its data source instead of creating a duplicate.

```bash
--replace
```

---

## Chart Management

### chart list

Lists all charts in a workbook.

```bash
xlforge chart list <file.xlsx>
xlforge chart list <file.xlsx> --json
xlforge chart list <file.xlsx> --sheet <name>
```

**JSON output:**
```json
[
  {
    "name": "RegionalSales",
    "type": "bar",
    "sheet": "Analysis",
    "range": "A1:G15"
  },
  {
    "name": "TrendLine",
    "type": "line",
    "sheet": "Analysis",
    "range": "A20:G35"
  }
]
```

### chart export

Exports a chart as an image file.

```bash
xlforge chart export <file.xlsx> <chart-name> --to <output-file>
```

**Formats:** `.png`, `.jpg`, `.svg`, `.pdf`

```bash
xlforge chart export report.xlsx "RegionalSales" --to sales.png
xlforge chart export report.xlsx "RegionalSales" --to sales.pdf
```

**Why?** Allows agents to share charts via Slack/Teams without opening Excel.

### chart delete

Deletes a chart by name.

```bash
xlforge chart delete <file.xlsx> <chart-name>
```

### chart update

Updates an existing chart's data source.

```bash
xlforge chart update <file.xlsx> <chart-name> <source> [options]
```

**Options:**
```
--x <column>      # Change X-axis
--y <columns>     # Change Y columns
--type <type>     # Change chart type
--replace         # Alias for update (same chart name)
```

---

## Complete Examples

### Basic bar chart
```bash
xlforge chart create report.xlsx "Analysis!A1" sales.csv \
    --type bar \
    --x "Region" \
    --y "Sales" \
    --title "Sales by Region"
```

### Multi-series line chart
```bash
xlforge chart create report.xlsx "Analysis!A1" data.csv \
    --type line \
    --x "Month" \
    --y "Sales,Forecast" \
    --name "SalesTrend" \
    --legend true
```

### Idempotent chart (safe to run multiple times)
```bash
xlforge chart create report.xlsx "A1" sales.csv \
    --type bar \
    --x "Region" \
    --y "Sales" \
    --name "RegionalSales" \
    --replace \
    --style modern
```

### Chart with hidden data sheet
```bash
xlforge chart create report.xlsx "Dashboard!A1" quarterly.csv \
    --type column \
    --x "Quarter" \
    --y "Revenue,Expenses" \
    --name "QuarterlyOverview" \
    --hidden-data "_xlforge" \
    --title "Q1-Q4 Performance"
```

### Export chart for sharing
```bash
xlforge chart export report.xlsx "SalesChart" --to chart.png
# Then share via Slack/Teams
```

---

## Implementation Notes

### Data Persistence Strategy

1. **CSV input + no `--hidden-data`:** Chart links to file path (fragile)
2. **CSV input + `--hidden-data`:** Creates hidden sheet `_xlforge_<name>`, links to it (persistent)
3. **Range input:** Chart links directly to existing cells (always persistent)

### Performance

For CSV-to-chart:
1. Fast-writer dumps CSV to temporary sheet (Fast-Path mode)
2. `Shapes.AddChart2` creates chart
3. Chart is named immediately for future updates

### Xlwings Note

Use `chart.set_source_data()` to link ranges reliably. For CSV, write to hidden sheet first, then link.

---

## Error Codes

| Code | Meaning |
|------|---------|
| `13` | Chart not found |
| `14` | Invalid chart type |
| `15` | Chart name already exists (use `--replace`) |
