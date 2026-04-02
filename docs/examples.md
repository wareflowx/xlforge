# Examples

## Complete Report Creation

### Single Command Mode
```bash
xlforge file open report.xlsx
xlforge sheet create report.xlsx "Summary"
xlforge sheet create report.xlsx "Data"
xlforge sheet rename report.xlsx "Sheet1" "Analysis"
xlforge cell set report.xlsx "Summary!A1" "Q1 2026 Performance"
xlforge format cell report.xlsx "Summary!A1" --bold --size 16 --color "44,72,116"
xlforge cell formula report.xlsx "Summary!B3" "=SUM(Data!B:B)"
xlforge format cell report.xlsx "Summary!B3" --number-format "€#,##0"
xlforge import csv report.xlsx sales.csv --sheet "Data" --cell A1 --has-headers
xlforge table create report.xlsx "Data!A1" sales.csv --style zebra --freeze-header
xlforge chart create report.xlsx "Analysis!A1" regional.csv \
    --type bar --x Region --y Sales --title "Sales by Region"
xlforge column auto-fit report.xlsx "Data!A:A"
xlforge freeze report.xlsx "Data!A2"
xlforge file save report.xlsx
```

### Safe Mode (with checkpoints)
```bash
xlforge checkpoint create report.xlsx "pre-agent"
xlforge run script.xlf --transaction
# If anything goes wrong:
xlforge checkpoint restore report.xlsx "pre-agent"
```

### Agent Mode (semantic query)
```bash
TARGET=$(xlforge query report.xlsx "Where is the total revenue cell?")
xlforge cell set report.xlsx "$TARGET" "999999"
```

---

## SQL-Agent Workflow

Complete enterprise data pipeline:

```bash
# 1. Query data across sources (Excel + CSV + Database)
xlforge sql query "
    SELECT e.name, e.salary, d.department_head
    FROM 'employees.xlsx!Data' e
    JOIN 'departments.csv' d ON e.dept_id = d.id
    WHERE e.salary > 100000
" --to high_earners.csv

# 2. Push SQL results to Excel Table with formatting
xlforge sql push "SELECT region, SUM(total) as revenue FROM orders GROUP BY region" \
    --db $PROD_DB \
    --to dashboard.xlsx "RevenueByRegion" \
    --format zebra --freeze-header

# 3. Update live inventory table
xlforge sql push "SELECT * FROM inventory WHERE status='active'" \
    --db $WAREHOUSE_DB \
    --to inventory.xlsx "LiveStock" \
    --mode upsert --key-col item_id

# 4. Pull Excel data to database
xlforge sql pull forecast.xlsx "ForecastData!A1:D500" \
    --into $DATA_WAREHOUSE \
    --table forecasts \
    --mode upsert --key-col forecast_id

# 5. Visual feedback
xlforge app focus dashboard.xlsx "RevenueByRegion"
xlforge app alert dashboard.xlsx "Data sync complete. 4,200 rows updated."
```

---

## Data Warehouse Example

Join multiple sources and push to Excel:

```bash
# Connect to enterprise databases
xlforge sql connect prod postgres://user:pass@prod-host/dbname
xlforge sql connect dw postgres://user:pass@dw-host/warehouse

# Complex multi-source query
xlforge sql query "
    SELECT
        p.product_name,
        p.category,
        SUM(s.quantity) as total_qty,
        SUM(s.amount) as total_revenue,
        i.current_stock
    FROM 'sales.xlsx!LineItems' s
    JOIN 'products.xlsx!Catalog' p ON s.product_id = p.id
    JOIN prod.inventory i ON p.sku = i.sku
    JOIN dw.returns r ON s.order_id = r.order_id
    WHERE s.sale_date >= '2026-01-01'
    GROUP BY p.product_name, p.category, i.current_stock
    HAVING SUM(s.amount) > 10000
    ORDER BY total_revenue DESC
" --to top_products.csv

# Push to executive dashboard
xlforge sql push "SELECT * FROM top_products.csv" \
    --to executive_dashboard.xlsx "TopProducts" \
    --format bordered --freeze-header
```

---

## Macro Recording Workflow

```bash
# Start recording your manual actions
xlforge record start report.xlsx --output my_actions.xlf

# Now manually:
# - Bold cell A1
# - Set value in B2
# - Create a chart
# - etc.

# Stop recording
xlforge record stop

# Now you have a reusable script!
xlforge run my_actions.xlf
```
