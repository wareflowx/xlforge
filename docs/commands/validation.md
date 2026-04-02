# Validation Commands

Data validation is the "UI Schema" of Excel. Setting validation creates a "Safe Landing Zone" for human users to interact with the spreadsheet without breaking automation logic.

---

## validation create

Adds data validation to a range.

```bash
xlforge validation create <file.xlsx> <sheet!range> --type <type> [options]
```

### Validation Types

| Type | Options |
|------|---------|
| `list` | `--formula "A,B,C"` or `--range <sheet!range>` or `--from-table <table[column]>` |
| `number` | `--min <n>` `--max <n>` |
| `decimal` | `--min <n>` `--max <n>` |
| `date` | `--min <date>` `--max <date>` |
| `time` | `--min <time>` `--max <time>` |
| `text-length` | `--min <n>` `--max <n>` (character count) |
| `boolean` | No options (preset: True/False) |
| `toggle` | No options (preset: Yes/No) |
| `formula` | `--formula "<excel-formula>"` |
| `dependent` | `--parent <cell>` `--map <file.json>` |

### Global Options

```
--title <text>           # Tooltip/title when cell is selected
--error <text>          # Error message when invalid value entered
--alert-type <type>     # stop, warning, information (default: stop)
--strict                # CLI respects validation (see below)
--allow-blank           # Empty cells are valid (default: true)
```

---

## Semantic Presets

### Boolean (True/False)

```bash
xlforge validation create report.xlsx "Data!Approved" --type boolean
# Creates dropdown: True, False
```

### Toggle (Yes/No)

```bash
xlforge validation create report.xlsx "Data!IsActive" --type toggle
# Creates dropdown: Yes, No
```

---

## Human Feedback

### Title (Tooltip)

Sets the tooltip that appears when the user selects the cell:

```bash
xlforge validation create report.xlsx "Data!B2:B20" \
    --type number \
    --min 1 --max 10 \
    --title "Performance Score" \
    --error "Score must be between 1 and 10"
```

### Error Messages

```bash
xlforge validation create report.xlsx "Data!Region" \
    --type list --formula "North,South,East,West" \
    --title "Select Region" \
    --error "Please select a valid region from the dropdown" \
    --alert-type stop
```

### Alert Types

| Type | Behavior |
|------|----------|
| `stop` | **Block input** - User must enter valid value |
| `warning` | **Warn but allow** - User can override |
| `information` | **Notify** - Just inform the user |

---

## List Validation

### Inline Formula

```bash
xlforge validation create report.xlsx "Data!Status" \
    --type list --formula "Active,Pending,Closed"
```

### Range Reference

Reference an existing range for the list:

```bash
xlforge validation create report.xlsx "Data!Category" \
    --type list --range "Lookup!A1:A10"
```

### Pipe from Stdin

Pull list values from another command:

```bash
sql-cli "SELECT name FROM categories" | \
    xlforge validation create report.xlsx "Data!Category" \
    --type list --formula -
```

### From Table Column (Dynamic)

Link to a Table column using structured references. **Automatically updates** when the table changes:

```bash
xlforge validation create report.xlsx "OrderSheet!Product" \
    --type list --from-table "Products[ProductName]"
```

**Why:** Uses Excel's structured references. Adding a product to `Products` table automatically updates the dropdown.

---

## Number/Date Validation

### Number Range

```bash
xlforge validation create report.xlsx "Data!Age" \
    --type number --min 0 --max 150

xlforge validation create report.xlsx "Data!Score" \
    --type number --min 0 --max 100 --alert-type warning
```

### Decimal (Allow Fractional)

```bash
xlforge validation create report.xlsx "Data!Price" \
    --type decimal --min 0.01 --max 999999.99
```

### Date Range

```bash
xlforge validation create report.xlsx "Data!StartDate" \
    --type date --min 2026-01-01 --max 2026-12-31
```

### Text Length

```bash
xlforge validation create report.xlsx "Data!Code" \
    --type text-length --min 3 --max 10
```

---

## Formula Validation

Custom Excel formula for complex rules:

```bash
# Must be unique in the range
xlforge validation create report.xlsx "Data!ID" \
    --type formula --formula "=COUNTIF($A$1:$A$100, A1) = 1"

# Must be uppercase
xlforge validation create report.xlsx "Data!Code" \
    --type formula --formula "=EXACT(A1, UPPER(A1))"
```

---

## Dependent Dropdowns (The "Holy Grail")

Multi-level dropdowns where selecting one value filters the next:

```bash
# Create dependent dropdowns
xlforge validation create report.xlsx "Category" \
    --type list --formula "Fruit,Vegetable,Meat"

xlforge validation create report.xlsx "Item" \
    --type dependent --parent "Category" --map category_map.json
```

### category_map.json

```json
{
  "Fruit": ["Apple", "Banana", "Orange"],
  "Vegetable": ["Carrot", "Lettuce", "Tomato"],
  "Meat": ["Beef", "Chicken", "Pork"]
}
```

**How it works:** The CLI creates hidden named ranges and sets up `INDIRECT` formulas so Excel automatically filters the second dropdown based on the first.

---

## validation get

Queries validation rules on a range.

```bash
xlforge validation get <file.xlsx> <sheet!range>
xlforge validation get <file.xlsx> <sheet!range> --json
```

### JSON Output

```json
{
  "range": "Data!B2:B20",
  "validation": {
    "type": "number",
    "min": 1,
    "max": 10,
    "allow_blank": true,
    "title": "Performance Score",
    "error": "Score must be between 1 and 10",
    "alert_type": "stop"
  }
}
```

### List Validation Output

```json
{
  "range": "Data!Status",
  "validation": {
    "type": "list",
    "formula": "Active,Pending,Closed",
    "allow_blank": true,
    "is_dynamic": false
  }
}
```

---

## validation list

Lists all validation rules in a sheet.

```bash
xlforge validation list <file.xlsx>
xlforge validation list <file.xlsx> --sheet "Data"
xlforge validation list <file.xlsx> --json
```

### JSON Output

```json
{
  "sheet": "Data",
  "validations": [
    {
      "range": "A1:A10",
      "type": "list",
      "formula": "Yes,No,Maybe",
      "title": "Selection Required"
    },
    {
      "range": "B2:B20",
      "type": "number",
      "min": 1,
      "max": 10,
      "title": "Score",
      "error": "Must be 1-10"
    }
  ]
}
```

---

## validation clear

Removes validation from a range.

```bash
xlforge validation clear <file.xlsx> <sheet!range>
```

**Example:**
```bash
# Remove validation before bulk import
xlforge validation clear report.xlsx "Data!A1:Z100"

# Import new data
xlforge import csv report.xlsx new_data.csv --sheet Data --cell A1

# Restore validation (re-create)
xlforge validation create report.xlsx "Data!Status" \
    --type list --formula "Active,Pending,Closed"
```

---

## Strict Mode (CLI Validation)

**Critical:** Excel's COM interface bypasses validation. `xlforge cell set A1 "InvalidValue"` will succeed even if the cell has a list validation.

### Enabling Strict Mode

**Command-level:**
```bash
xlforge cell set report.xlsx "Data!B2" "Invalid" --strict
# Error 110: Value "Invalid" violates validation (list: Yes,No,Maybe)
```

**Global (environment variable):**
```bash
export XLFORGE_STRICT=true
xlforge cell set report.xlsx "Data!B2" "Invalid"
# Error 110: Value violates validation
```

**Session-level:**
```bash
xlforge validation strict on
xlforge cell set report.xlsx "Data!B2" "Invalid"  # Now checks
xlforge validation strict off
```

### Strict Mode Behavior

When `--strict` is enabled, the CLI:
1. Reads the validation rule on the target cell
2. Evaluates whether the value satisfies the rule
3. Returns error `110` if validation fails
4. Does not write the value

**Supported validations in strict mode:**
- `list` - Value must be in the list
- `number` / `decimal` - Value must be within min/max
- `date` / `time` - Value must be within range
- `text-length` - Character count must be within range

---

## Performance

Validation uses `Range.Validation.Add` on the entire range object (not cell-by-cell):

```bash
# Efficient: single COM call
xlforge validation create report.xlsx "Data!A1:A1000" \
    --type list --formula "Option1,Option2,Option3"

# NOT: looping through A1, then A2, then A3...
```

**Range performance:** Adding validation to 1000 cells takes ~50ms via COM.

---

## Complete Examples

### Professional data entry form
```bash
# Create validation with full UX
xlforge validation create report.xlsx "Entry!Customer" \
    --type list \
    --from-table "Customers[CustomerName]" \
    --title "Select Customer" \
    --error "Please select a customer from the list" \
    --alert-type stop

xlforge validation create report.xlsx "Entry!Amount" \
    --type decimal \
    --min 0.01 \
    --max 999999.99 \
    --title "Invoice Amount" \
    --error "Amount must be between 0.01 and 999,999.99" \
    --alert-type stop

xlforge validation create report.xlsx "Entry!Date" \
    --type date \
    --min 2026-01-01 \
    --title "Invoice Date" \
    --error "Date must be in 2026" \
    --alert-type warning
```

### Dynamic product selection
```bash
# Main category dropdown
xlforge validation create report.xlsx "Order!Category" \
    --type list \
    --from-table "Products[Category]" \
    --title "Product Category"

# Dependent product dropdown
xlforge validation create report.xlsx "Order!Product" \
    --type dependent \
    --parent "Order!Category" \
    --map product_map.json \
    --title "Select Product"

# Dependent quantity
xlforge validation create report.xlsx "Order!Qty" \
    --type number \
    --min 1 \
    --max 1000 \
    --title "Quantity"
```

### Strict mode workflow
```bash
# Enable strict validation checking
xlforge validation strict on

# These will fail validation
xlforge cell set report.xlsx "Score" "150"    # Error 110: exceeds max
xlforge cell set report.xlsx "Status" "Maybe" # Error 110: not in list

# These will succeed
xlforge cell set report.xlsx "Score" "85"
xlforge cell set report.xlsx "Status" "Active"

xlforge validation strict off
```

---

## Error Codes

| Code | Meaning |
|------|---------|
| `110` | Value violates validation (strict mode) |
| `111` | Validation type not supported |
| `112` | Invalid formula syntax |
| `113` | Dependent validation map not found |
| `114` | Parent cell validation not found |
| `115` | Circular dependency in dependent validation |
| `116` | Validation range is too large |
