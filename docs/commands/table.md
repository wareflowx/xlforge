# Table Commands

Excel Tables (ListObjects) provide structured references, automatic formatting, and serve as the "Source of Truth" for Pivot Tables.

---

## table create

Creates an Excel Table (ListObject) from data.

```bash
xlforge table create <file.xlsx> <sheet!cell> <data.csv> [options]
xlforge table create <file.xlsx> <sheet!cell> --query "<sql>" [options]
```

### Options

```
--name <table-name>     # Table name (required for referencing later)
--style <style>         # zebra, zebra-striped, bordered, plain, grid (default: plain)
--headers               # Use first row as headers
--no-headers            # No header row
--freeze-header         # Freeze header row (default: true)
--auto-filter           # Enable auto-filter (default: true)
--column <formula>      # Define calculated column (can be specified multiple times)
```

### Calculated Columns

Define formula columns during creation:

```bash
xlforge table create report.xlsx "Data!A1" sales.csv \
    --name "SalesData" \
    --column "Tax=[@Revenue]*0.15" \
    --column "Profit=[@Revenue]-[@Cost]"
```

**Formula syntax:** `=[@ColumnName]` for structured references.

### Defaults

By default, `table create` automatically applies:
- `freeze-header` - Header row stays visible when scrolling
- `auto-filter` - Enables column filtering

Use `--no-freeze` or `--no-filter` to disable.

### Examples

```bash
# Create table from CSV
xlforge table create report.xlsx "Summary!A1" summary.csv \
    --name "SummaryData" \
    --style zebra-striped \
    --headers

# Create table from SQL query
xlforge table create report.xlsx "Data!A1" \
    --query "SELECT * FROM sales WHERE date > '2026-01-01'" \
    --db prod \
    --name "SalesData" \
    --freeze-header
```

---

## table link

Creates a live ODBC/OLEDB connection inside Excel.

```bash
xlforge table link <file.xlsx> <table-name> [options]
```

### Options

```
--connection <dsn>      # DSN name or connection string
--query <sql>           # SQL query to execute on refresh
--timeout <seconds>     # Query timeout (default: 300)
--strip-passwords       # Remove passwords before saving file
--sync                  # Immediately sync after creation
```

### DSN Support

Use DSN names for security (passwords stored in ODBC admin, not in file):

```bash
xlforge table link report.xlsx "LiveInventory" \
    --connection "DSN=WarehouseDB" \
    --query "SELECT * FROM stock_levels"
```

### Security

**Warning:** Connection strings are stored in the `.xlsx` file. Use `--strip-passwords` before distributing:

```bash
xlforge table link report.xlsx "ProdData" \
    --connection "Server=prod.db.com;Database=main;User=admin;Password=secret" \
    --query "SELECT * FROM metrics"

xlforge table link report.xlsx "ProdData" --strip-passwords
# Sanitizes connection string in the saved file
```

**Best practice:** Use DSN names instead of connection strings when possible.

### Example

```bash
xlforge table link report.xlsx "LiveSales" \
    --connection "DSN=SalesDB" \
    --query "SELECT region, SUM(amount) FROM sales GROUP BY region" \
    --timeout 600
```

**Result:** User opens Excel and clicks **Data > Refresh** to fetch fresh data.

---

## table sync-schema

Updates Excel Table headers to match a database schema without destroying user formatting.

```bash
xlforge table sync-schema <file.xlsx> <table-name> \
    --db <connection-url> \
    [options]
```

### Options

```
--db <connection-url>     # Database connection
--schema <name>          # Schema name (default: public)
--dry-run                # Preview changes without applying
--prune                  # Remove columns that don't exist in DB
--strict                 # Error on schema mismatches instead of warning
--patch-formulas         # Attempt to patch formulas referencing renamed columns
```

### Prune Mode

Use `--prune` to remove columns that no longer exist in the database:

```bash
xlforge table sync-schema report.xlsx "SalesData" \
    --db prod \
    --prune
```

**Without `--prune`:** Extra columns in Excel are preserved.
**With `--prune`:** Extra columns are deleted.

### Dry Run

Preview changes before applying:

```bash
xlforge table sync-schema report.xlsx "SalesData" --dry-run --db prod
```

**Output:**
```
DRY RUN: Schema sync for "SalesData"
+---------------------+---------------------+
| Action              | Column              |
+---------------------+---------------------+
| RENAME              | revenue → Revenue   |
| ADD                 | Discount (DECIMAL)  |
| PRUNE               | OldColumn           |
+---------------------+---------------------+
```

### Safety

`sync-schema` preserves:
- Cell formatting
- Formulas in adjacent columns
- Pivot Tables referencing the table
- User's manual changes outside the table

**Patch Formulas:** Use `--patch-formulas` to update formulas that reference renamed columns:

```bash
xlforge table sync-schema report.xlsx "SalesData" \
    --db prod \
    --patch-formulas
# Detects =[@revenue] becomes =[@Revenue] after rename
```

---

## table refresh

Refreshes a linked Excel Table.

```bash
xlforge table refresh <file.xlsx> <table-name>
xlforge table refresh <file.xlsx> <table-name> --sync
xlforge table refresh <file.xlsx> --all
```

### Synchronous Execution

**Critical:** Excel's default is background refresh (`BackgroundQuery = True`), which means the command returns immediately while data is still loading.

The CLI forces **synchronous execution** by setting `BackgroundQuery = False`:

```bash
xlforge table refresh report.xlsx "LiveInventory" --sync
# Waits until data is actually in the sheet before returning
```

**Why:** If you run `table refresh` followed by `sql push`, but the refresh is still running in background, your push will read stale data.

### Refresh with Data Push

Complete workflow for live data updates:

```bash
# 1. Refresh the linked table
xlforge table refresh report.xlsx "LiveSales" --sync

# 2. Now push to another database (data is guaranteed current)
xlforge sql push "SELECT * FROM 'report.xlsx!LiveSales'" --db warehouse
```

### Example

```bash
# Refresh single table
xlforge table refresh report.xlsx "SalesData"

# Refresh all tables in workbook
xlforge table refresh report.xlsx --all

# Force sync mode (wait for completion)
xlforge table refresh report.xlsx "Inventory" --sync
```

---

## table export

Exports table data in various formats.

```bash
xlforge table export <file.xlsx> <table-name>
xlforge table export <file.xlsx> <table-name> --json
xlforge table export <file.xlsx> <table-name> --csv
xlforge table export <file.xlsx> <table-name> --to <output.file>
```

### JSON Export

Exports with headers as keys (more agent-friendly than arrays):

```bash
xlforge table export report.xlsx "SalesData" --json
```

**Output:**
```json
[
  {"Date": "2026-01-01", "Region": "North", "Revenue": 15000},
  {"Date": "2026-01-02", "Region": "South", "Revenue": 12000}
]
```

### CSV Export

```bash
xlforge table export report.xlsx "SalesData" --to sales_backup.csv
```

---

## table pivot

Creates a Pivot Table based on an existing Table.

```bash
xlforge table pivot <file.xlsx> <source-table> [options]
```

### Options

```
--sheet <name>          # Target sheet (default: new sheet)
--name <pivot-name>     # Pivot Table name
--rows <field>          # Row field (can specify multiple)
--columns <field>       # Column field
--values <aggregation>  # Value field with aggregation (e.g., SUM:Revenue, COUNT:ID)
--filters <field>       # Filter field
--style <style>         # Pivot Style (default: plain)
--replace               # Replace if pivot with same name exists
```

### Aggregations

```
SUM:<field>     # Sum of values
COUNT:<field>   # Count of values
AVERAGE:<field> # Average
MIN:<field>     # Minimum
MAX:<field>     # Maximum
```

### Examples

```bash
# Create regional sales summary
xlforge table pivot report.xlsx "SalesData" \
    --sheet "Dashboard" \
    --name "RegionalSales" \
    --rows "Region" \
    --values "SUM:Revenue" \
    --values "SUM:Cost" \
    --style zebra

# Multi-dimensional pivot
xlforge table pivot report.xlsx "SalesData" \
    --rows "Region" \
    --columns "Quarter" \
    --values "SUM:Revenue" \
    --filters "Year"
```

---

## table writeback

Detects user edits in an Excel Table and pushes them back to the database.

```bash
xlforge table writeback <file.xlsx> <table-name> \
    --db <connection-url> \
    --key <key-column> \
    [options]
```

### Options

```
--db <connection-url>     # Target database
--key <column>           # Primary key column for matching rows
--dry-run                # Preview changes without applying
--batch-size <n>         # Rows per batch (default: 100)
```

### How It Works

1. Scans the Excel Table for "dirty" rows (edited since last writeback)
2. Compares with database using `--key` column
3. Generates UPDATE or INSERT statements
4. Executes in batch

### Example

```bash
# Setup inventory table with writeback
xlforge table link report.xlsx "Inventory" \
    --connection "DSN=Warehouse" \
    --query "SELECT item_id, name, stock FROM inventory"

# User manually edits stock levels in Excel...
# User runs writeback

xlforge table writeback report.xlsx "Inventory" \
    --db warehouse \
    --key item_id
```

### Use Case

Turns Excel into a **CRUD interface** for production data. Users can:
1. Refresh live data
2. Manually adjust values
3. Writeback only the changed rows

---

## table list

Lists all Tables in the workbook.

```bash
xlforge table list <file.xlsx>
xlforge table list <file.xlsx> --json
xlforge table list <file.xlsx> --sheet "Data"
```

### JSON Output

```json
{
  "tables": [
    {
      "name": "SalesData",
      "sheet": "Data",
      "range": "A1:E1000",
      "has_link": true,
      "link_connection": "DSN=SalesDB",
      "refreshed_at": "2026-03-31T10:30:00"
    }
  ]
}
```

---

## table rename

Renames a table.

```bash
xlforge table rename <file.xlsx> <old-name> <new-name>
```

**Note:** Updates all Pivot Tables and formulas that reference the old name.

---

## table delete

Deletes a table (not the data).

```bash
xlforge table delete <file.xlsx> <table-name>
xlforge table delete <file.xlsx> <table-name> --keep-data
```

### Options

```
--keep-data    # Delete table formatting only, keep cell values
```

---

## Table Shorthand

Once a Table is named, use the table name directly instead of sheet references:

```bash
# Instead of specifying sheet and range:
xlforge cell get report.xlsx "Data!A1"
xlforge cell set report.xlsx "Data!B5" "Value"

# Use table name directly (xlforge finds it):
xlforge cell get "SalesData[Revenue]"           # Get Revenue column header
xlforge cell set "SalesData[Status]" "Active"   # Set Status column value
xlforge table refresh "SalesData"                # Refresh by table name
```

### Structured References

Table shorthand supports Excel's structured references:

```
SalesData[#All]          # Entire table including headers
SalesData[#Data]         # Data body only
SalesData[#Headers]      # Header row only
SalesData[#Totals]       # Total row
SalesData[@Revenue]      # Current row, Revenue column
SalesData[@]             # Current row
```

---

## Complete Examples

### Create a fully-configured table
```bash
xlforge table create report.xlsx "Data!A1" sales.csv \
    --name "SalesData" \
    --style zebra-striped \
    --headers \
    --freeze-header \
    --auto-filter \
    --column "Tax=[@Revenue]*0.15" \
    --column "Margin=([@Revenue]-[@Cost])/[@Revenue]"
```

### Live dashboard workflow
```bash
# 1. Create and link table
xlforge table link report.xlsx "LiveSales" \
    --connection "DSN=ProductionDB" \
    --query "SELECT * FROM sales WHERE date >= CURRENT_DATE" \
    --sync

# 2. Build pivot from linked data
xlforge table pivot report.xlsx "LiveSales" \
    --sheet "Dashboard" \
    --name "SalesSummary" \
    --rows "Region" \
    --columns "Quarter" \
    --values "SUM:Revenue"

# 3. Refresh before meetings
xlforge table refresh report.xlsx --all --sync
```

### Data entry with writeback
```bash
# 1. Setup editable table
xlforge table create report.xlsx "Entry!A1" template.csv \
    --name "DataEntry" \
    --style zebra

# 2. User enters data...
# 3. Writeback changes to database
xlforge table writeback report.xlsx "DataEntry" \
    --db prod \
    --key record_id \
    --dry-run

# Review changes
xlforge table writeback report.xlsx "DataEntry" \
    --db prod \
    --key record_id
```

### Schema migration with formula patching
```bash
# Sync schema and patch formulas
xlforge table sync-schema report.xlsx "SalesData" \
    --db prod \
    --prune \
    --patch-formulas \
    --dry-run

# Apply changes
xlforge table sync-schema report.xlsx "SalesData" \
    --db prod \
    --prune \
    --patch-formulas
```

---

## Error Codes

| Code | Meaning |
|------|---------|
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
