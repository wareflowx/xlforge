# SQL Bridge (DuckDB Integration)

xlforge uses **DuckDB** as its internal SQL engine for high-performance data operations. This enables set-based operations against Excel files without triggering Excel's UI thread.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                      xlforge SQL Bridge                     │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│   Excel File ──► DuckDB ──► SQL Query ──► Results          │
│        │                            │                       │
│        │         (Read Path)       │                       │
│        │                            ▼                       │
│        │                      CSV/Parquet                   │
│        │                            │                       │
│        │         (Write Path)      │                       │
│        ▼                            │                       │
│   Native Excel ◄────────────────────┘                       │
│   (ListObject/Table)                                         │
│                                                              │
└─────────────────────────────────────────────────────────────┘
```

**Performance:**
- Read: ~100ms for 100k rows (direct file pointer, no COM)
- Write: <1 second for 50k rows (Direct-to-OpenXML mode)

---

## Concurrency Manager

The SQL Bridge automatically handles the "open vs. closed" file conflict.

### Write Strategy

| File State | Method | Why |
|------------|--------|-----|
| **Closed** | Direct-to-OpenXML | Ultra-fast, bypasses COM |
| **Open in Excel** | COM `Range.Value2` + Temp CSV | Prevents corruption |
| **File locked** | Retry with backoff (3 attempts) | Wait for release |

```bash
xlforge sql push "SELECT * FROM sales" --db prod --to report.xlsx "Sales"
# Automatically detects file state and uses optimal method
```

### Silent Refresh

When writing to an **open file**, the bridge automatically:
1. Writes to a temporary CSV
2. Uses COM `Range.Value2 = array` for high-speed insertion
3. Calls `Table.Refresh()` on any affected Excel Tables
4. Triggers Pivot Table refresh if `--refresh-pivots` is specified

```bash
xlforge sql push "SELECT * FROM live_data" \
    --db prod \
    --to dashboard.xlsx "SalesData" \
    --refresh-pivots
# After push: all Pivot Tables showing SalesData are auto-refreshed
```

---

## sql query

Runs SQL against Excel files, CSVs, Parquet, and any DuckDB-mounted data source.

```bash
xlforge sql query "<query>"
xlforge sql query "<query>" --json
xlforge sql query "<query>" --to <output.csv>
xlforge sql query "<query>" --explain
```

### Virtual Table Syntax

- Excel sheet: `'file.xlsx!SheetName'` or `'file.xlsx![SheetName$]'`
- CSV file: `'file.csv'`
- Parquet: `'file.parquet'`
- Glob patterns: `'*.csv'` or `'sales_*.xlsx'`

### Glob Support

DuckDB's glob support allows querying multiple files at once:

```bash
# Query all CSV files in current directory
xlforge sql query "SELECT * FROM '*.csv'" --json

# Query all sales workbooks
xlforge sql query "
    SELECT region, SUM(total) as revenue
    FROM 'sales_*.xlsx!Data'
    GROUP BY region
" --json
```

### Explain Feature

Debug complex queries by viewing the execution plan:

```bash
xlforge sql explain "
    SELECT * FROM 'orders.xlsx!Data' o
    JOIN 'products.xlsx!Data' p ON o.product_id = p.id
"
```

**Output:**
```
┌─────────────────────┐
│     Query Plan       │
├─────────────────────┤
│ HashJoin (o.product_│
│   │   id = p.id)    │
│ ├── Scan: orders    │
│ └── Scan: products  │
│                     │
│ Est. rows: 1.2M     │
│ Warning: Large join  │
└─────────────────────┘
```

### AI Filter (Semantic Enhancement)

Combine SQL with semantic search:

```bash
xlforge sql query "SELECT * FROM 'transactions.xlsx!Data'" \
    --ai-filter "Find only rows that look like fraudulent transactions"
```

**How it works:**
1. Uses `DESCRIBE` to get table schema
2. Applies semantic filter via LLM (local or cloud)
3. Returns filtered results via SQL

---

## sql push

Pushes SQL results directly into an Excel Table (ListObject).

```bash
xlforge sql push "<query>" \
    --db <connection-url> \
    --to <file.xlsx> <table-name> \
    [options]
```

### Options

```
--db <url>              # Database connection (postgres://, mysql://, etc.)
--to <file> <name>     # Target Excel file and table name
--mode <mode>           # insert, upsert, replace (default: replace)
--key-col <column>      # For upsert mode: key column to match
--format <style>        # zebra, bordered, plain, grid (default: plain)
--freeze-header          # Freeze header row
--auto-filter           # Enable auto-filter
--refresh-pivots        # Refresh all Pivot Tables using this data
--strict                # Error on type mismatches instead of casting
```

### Modes

| Mode | Behavior |
|------|----------|
| `replace` | Drop table and recreate (default) |
| `upsert` | Update existing rows by key, insert new ones |
| `insert` | Append new rows only |
| `append` | Alias for insert |

### Upsert Example

```bash
xlforge sql push "SELECT * FROM inventory WHERE updated_at > $last_sync" \
    --db warehouse \
    --to inventory.xlsx "StockLevels" \
    --mode upsert \
    --key-col item_id \
    --format zebra \
    --refresh-pivots
```

**How upsert works:**
1. Pulls existing Excel Table data into DuckDB
2. Joins incoming SQL data with existing data on `--key-col`
3. Generates merged dataset
4. Overwrites Excel Table with merged data
5. Refreshes dependent Pivot Tables

### Pivot Refresh

After pushing data, automatically refresh all Pivot Tables that depend on the target Table:

```bash
xlforge sql push "SELECT * FROM live_sales" \
    --db prod \
    --to dashboard.xlsx "SalesData" \
    --refresh-pivots

# All Pivot Tables using "SalesData" are automatically refreshed
```

---

## sql pull

Extracts data from Excel and loads it into a database.

```bash
xlforge sql pull <file.xlsx> <sheet!range> \
    --into <connection-url> \
    [options]
```

### Options

```
--into <url>            # Target database connection
--table <name>          # Target table name (default: derived from sheet)
--mode <mode>           # insert, upsert, replace, append (default: append)
--key-col <column>      # For upsert: key column to match
--batch-size <n>        # Rows per batch (default: 10000)
--strict                # Error on type mismatches instead of casting to NULL
```

### Schema Inference

DuckDB automatically infers column types using `sniff`:

```bash
xlforge sql pull report.xlsx "Data!A1:Z100000" \
    --into postgres://user:pass@host/dbname \
    --table sales_data

# Automatically generates:
# CREATE TABLE sales_data (
#   id INTEGER,
#   date DATE,
#   amount DECIMAL(10,2),
#   region VARCHAR,
#   ...
# )
```

### Strict Mode

Handle "vibes-based" Excel columns (where a column has mixed types):

```bash
# Without --strict (default): casts errors to NULL
xlforge sql pull report.xlsx "Prices" --into warehouse --table prices

# With --strict: errors and reports the issue
xlforge sql pull report.xlsx "Prices" --into warehouse --table prices --strict
# Error 88: Type coercion failed - Prices!B5 contains "N/A" in numeric column
```

---

## sql connect

Creates a named database connection for reuse.

```bash
xlforge sql connect <name> <connection-url>
xlforge sql connect list
xlforge sql connect delete <name>
```

### Secret Management

Connections are stored in `~/.xlforge/connections.json` (encrypted) or use environment variables:

```bash
# Store encrypted connection
xlforge sql connect prod postgres://user:pass@host/dbname

# Use environment variable (recommended for CI/CD)
export DATABASE_URL="postgres://user:pass@host/dbname"
xlforge sql connect prod $DATABASE_URL
```

### Connection Examples

```bash
xlforge sql connect prod postgres://user:pass@prod-host/prod
xlforge sql connect warehouse postgres://user:pass@warehouse-host/warehouse
xlforge sql connect analytics mysql://user:pass@analytics-host/reporting

# Now use by name
xlforge sql push "SELECT * FROM sales" --db prod --to report.xlsx "Sales"
xlforge sql pull report.xlsx "Data" --into warehouse --table raw_sales
```

---

## sql view (Virtual Sheets)

Creates an Excel "Power Query" connection without writing data to cells.

```bash
xlforge sql view <file.xlsx> <view-name> --query "<sql-query>"
xlforge sql view <file.xlsx> <view-name> --query "<sql-query>" --refresh-interval 60
xlforge sql view list <file.xlsx>
xlforge sql view delete <file.xlsx> <view-name>
```

### How It Works

Creates an Excel connection definition (not data). User hits "Refresh" in Excel to run the query.

```bash
# Create a live connection to production data
xlforge sql view report.xlsx "LiveRevenue" \
    --query "SELECT region, SUM(amount) FROM prod.sales GROUP BY region" \
    --db prod

# User can now refresh this in Excel whenever they want
# Data is fetched directly from the database, not stored in cells
```

### Auto-Refresh

Set a refresh interval for automatic updates:

```bash
xlforge sql view report.xlsx "LiveInventory" \
    --query "SELECT * FROM warehouse.stock_levels" \
    --db warehouse \
    --refresh-interval 300  # Refresh every 5 minutes
```

---

## Zero-ETL Joins

DuckDB extensions enable joining Excel data with live production databases and S3 files.

### Automatic Extension Loading

xlforge auto-loads DuckDB extensions when needed:

```bash
# No explicit extension installation needed
xlforge sql query "
    SELECT o.order_id, o.amount, p.name, p.stock
    FROM 'orders.xlsx!Data' o
    JOIN postgres.prodb.products p ON o.product_id = p.id
    JOIN 's3://warehouse/sales.parquet' s ON o.region = s.region
" --json
```

### Supported Extensions

| Extension | Syntax | Use Case |
|----------|--------|----------|
| PostgreSQL | `postgres.<db>.<schema>.<table>` | Live production joins |
| MySQL | `mysql.<db>.<table>` | Read replica joins |
| S3/Parquet | `'s3://bucket/path/*.parquet'` | Data lake joins |
| SQLite | `sqlite.<db>.<table>` | Local database joins |

### Cross-Source Join Example

```bash
xlforge sql query "
    SELECT
        e.name,
        e.salary,
        d.department_head,
        p.avg_salary as industry_avg
    FROM 'employees.xlsx!Data' e
    JOIN 'departments.csv' d ON e.dept_id = d.id
    JOIN postgres.hr.departments p ON e.dept_id = p.id
    WHERE e.salary > p.avg_salary * 1.5
" --json
```

---

## Complete Examples

### Live Dashboard Workflow

```bash
# Setup: Create connection once
xlforge sql connect prod postgres://user:pass@prod-host/main

# Daily sync with pivot refresh
xlforge sql push "SELECT * FROM sales WHERE date >= CURRENT_DATE - INTERVAL '30 days'" \
    --db prod \
    --to dashboard.xlsx "RecentSales" \
    --mode replace \
    --format zebra-striped \
    --freeze-header \
    --refresh-pivots

# Check for anomalies using AI filter
xlforge sql query "SELECT * FROM 'dashboard.xlsx!RecentSales'" \
    --ai-filter "Find rows with unusual patterns"
```

### Multi-File Aggregation

```bash
# Aggregate all quarterly sales files
xlforge sql query "
    SELECT
        'Q1' as quarter, region, SUM(amount)
    FROM 'sales_q1.xlsx!Data'
    GROUP BY region

    UNION ALL

    SELECT
        'Q2' as quarter, region, SUM(amount)
    FROM 'sales_q2.xlsx!Data'
    GROUP BY region
" --to annual_summary.csv

# Or use glob for cleaner syntax
xlforge sql query "
    SELECT quarter, region, SUM(amount)
    FROM 'sales_q*.xlsx!Data'
    GROUP BY quarter, region
" --to annual_summary.csv
```

### Upsert for Real-Time Data

```bash
# First sync: full replace
xlforge sql push "SELECT * FROM inventory" \
    --db warehouse \
    --to stock.xlsx "Inventory" \
    --mode replace

# Subsequent syncs: upsert only changed rows
xlforge sql push "SELECT * FROM inventory WHERE updated_at > NOW() - INTERVAL '1 hour'" \
    --db warehouse \
    --to stock.xlsx "Inventory" \
    --mode upsert \
    --key-col item_id

# Verify
xlforge sql query "SELECT COUNT(*) FROM 'stock.xlsx!Inventory'"
```

### Debug Query Performance

```bash
# Check if query will be efficient
xlforge sql explain "
    SELECT * FROM 'huge_file.xlsx!Data' a
    JOIN 'another_large.xlsx!Data' b ON a.id = b.id
"

# If you see "CrossJoin" warnings, add filters
xlforge sql query "
    SELECT a.id, a.total, b.category
    FROM 'huge_file.xlsx!Data' a
    JOIN 'another_large.xlsx!Data' b ON a.id = b.id
    WHERE a.date >= '2026-01-01'  -- Add filter to reduce rows
" --json
```

---

## Error Codes

| Code | Meaning |
|------|---------|
| `88` | Type coercion failed (use `--strict` for details) |
| `89` | Database connection failed |
| `90` | Query timeout |
| `91` | File is locked (Excel has it open) |
| `92` | Upsert key column not found |
| `93` | Schema mismatch (pull requires `--strict` or manual CAST) |
| `94` | Extension not available (e.g., postgres_scanner) |
| `95` | Virtual view connection failed |
| `96` | Pivot refresh failed (no matching pivot found) |
