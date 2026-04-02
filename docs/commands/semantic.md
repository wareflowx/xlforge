# Semantic Commands (AI Context)

Transforms Excel into an AI-queryable knowledge base. Designed for LLM agents that need to understand spreadsheet structure without reading every cell.

---

## index create

Creates a vector embedding index of the spreadsheet for semantic search.

```bash
xlforge index create <file.xlsx>
xlforge index create <file.xlsx> --force   # Rebuild index
xlforge index create <file.xlsx> --engine local   # Local embedding model
xlforge index create <file.xlsx> --provider ollama # Ollama endpoint
```

### Engine Options

| Engine | Use Case | Privacy |
|--------|----------|---------|
| `cloud` (default) | OpenAI / Anthropic embeddings | Data sent to external API |
| `local` | Ollama / Llama.cpp | All data stays on-premise |

```bash
# Local-first for corporate security
xlforge index create report.xlsx --engine local --provider ollama
```

### Privacy Check

Scans for sensitive patterns before indexing. Prevents PII from leaving the network.

```bash
xlforge index create report.xlsx --privacy-check
```

**Detects:**
- SSNs (`\d{3}-\d{2}-\d{4}`)
- Credit Cards (`\d{4}[\s-]?\d{4}[\s-]?\d{4}[\s-]?\d{4}`)
- Email addresses (when not in allowed domains)
- API keys / secrets (common patterns)

**Output:**
```json
{
  "privacy_check": "passed",
  "issues": [],
  "cells_scanned": 50000,
  "sensitive_patterns_found": 0
}
```

**On failure:**
```json
{
  "privacy_check": "failed",
  "issues": [
    {"cell": "EmployeeData!D15", "type": "ssn", "redacted": "XXX-XX-XXXX"}
  ],
  "recommendation": "Remove or redact sensitive data before indexing"
}
```

### What Gets Indexed

- Cell text values (with sampling for large ranges)
- Column headers
- Sheet names
- Data patterns (dates, currencies, percentages)
- Named ranges
- Table names

**Output:** `.xlforge/index/<file>.json`

---

## index watch

Background daemon that updates the vector index in real-time as the user edits.

```bash
xlforge index watch <file.xlsx>
xlforge index watch <file.xlsx> --debounce 2s   # Update after 2s of inactivity
```

**Why:** Ensures `query` is never more than a few seconds out of date. The agent stays "aware" of what the human is doing.

**Behavior:**
- Runs as a lightweight background process
- Monitors `SheetChange` events via COM
- Batches changes to avoid excessive rebuilds
- Auto-stops when Excel closes

---

## query

Semantic search - find data by meaning, not coordinates.

```bash
xlforge query <file.xlsx> "<question>"
xlforge query <file.xlsx> "<question>" --json
xlforge query <file.xlsx> "<question>" --min-confidence 0.85
```

### Coordinate-Only Output

For agent pipelines, query can return just the coordinate:

```bash
xlforge query report.xlsx "Net Profit" --coordinate
# Output: Summary!B7
```

**Agent workflow:**
```bash
COORD=$(xlforge query report.xlsx "Total Revenue 2025" --coordinate)
xlforge cell set $COORD "50000"
```

### Full JSON Output

```json
{
  "question": "Where is the total revenue for North America?",
  "answer": "Summary!B7",
  "formula": "=SUM(Data!D:D)",
  "value": 1500000,
  "confidence": 0.92,
  "context": "Found in the Summary sheet, row 7"
}
```

### Query Shorthand

Use `query:` prefix in any cell command for inline semantic lookup:

```bash
xlforge cell get "query:Net Profit 2025"
# Internally runs: xlforge query report.xlsx "Net Profit 2025"
# Then fetches the cell at the returned coordinate
```

---

## describe

Returns an LLM-optimized summary of a range using statistical sampling.

```bash
xlforge describe <file.xlsx> <sheet!range>
xlforge describe <file.xlsx> <sheet!range> --json
xlforge describe <file.xlsx> <sheet!range> --sample 20
```

### Sampling Strategy

For large ranges (10,000+ rows), `describe` uses **Statistical Sampling** to stay within token limits:

1. **Headers** - First row (always)
2. **First 5 rows** - Initial data pattern
3. **Last 5 rows** - Recent data / totals
4. **5 random rows** - Mid-range diversity
5. **Basic statistics** - Min, Max, Mean, Null-count per column

**Example for a 100,000-row dataset:**
```
Sampling 21 rows from 100,000 (0.021%)
Headers: [Date, Region, Revenue, Expenses, Profit]
Stats computed: Min, Max, Mean, Nulls per column
```

### Object Metadata

`describe` also reports on Excel objects within the range:

```bash
xlforge describe report.xlsx "Summary!A1:Z50" --include-objects
```

**Additional output:**
```markdown
## Objects in Range

### Pivot Tables
- "SalesPivot" (A10:H25) - Regional breakdown by quarter

### Named Ranges
- "TotalRevenue" → Summary!B7
- "GrowthRate" → Summary!E7

### Charts
- "RevenueChart" (I10:U35) - Bar chart, linked to Data!A1:E50

### Data Validation
- C2:C100 - Dropdown list (North, South, East, West)
```

### Schema-Only Output

For agents that need machine-readable schema without prose:

```bash
xlforge describe report.xlsx "Data!A1:Z1000" --schema-only --json
```

**JSON Schema output:**
```json
{
  "range": "Data!A1:Z1000",
  "columns": [
    {"name": "Date", "type": "date", "nulls": 0, "sample": ["2026-01-01", "2026-01-02"]},
    {"name": "Region", "type": "categorical", "values": ["North", "South", "East", "West"], "nulls": 5},
    {"name": "Revenue", "type": "currency", "min": 1000, "max": 500000, "mean": 25000, "nulls": 0}
  ],
  "row_count": 1000,
  "tables": ["SalesData"],
  "named_ranges": ["TotalRevenue"],
  "charts": 3
}
```

### Full Markdown Output

```markdown
## Summary!A1:F50

**Purpose:** Quarterly sales report with regional breakdown

**Columns:**
- A: Region (categorical: North, South, East, West)
- B: Revenue (currency, €)
- C: Expenses (currency, €)
- D: Profit (currency, €)
- E: Growth (percentage)
- F: Quarter (Q1, Q2, Q3, Q4)

**Statistics (sampled):**
| Column | Min | Max | Mean | Nulls |
|--------|-----|-----|------|-------|
| Revenue | 10,000 | 500,000 | 125,000 | 0 |
| Expenses | 5,000 | 300,000 | 75,000 | 2 |

**Patterns:**
- Row 1: Headers (bold)
- Rows 2-5: Q1-Q4 data
- Row 6: Totals row (bold, shaded)
- Missing values in E3 (no growth data for Q1)

**Objects:**
- Pivot Table "SalesPivot" at A10:H25
- Named Range "TotalRevenue" → B7
```

---

## semantic-check

AI-powered linter that validates spreadsheet against business rules.

```bash
xlforge semantic-check <file.xlsx> --rules "Ensure no negative numbers in Revenue"
xlforge semantic-check <file.xlsx> --rules-file business_rules.txt --json
```

**Example:**
```bash
xlforge semantic-check report.xlsx --rules "Revenue must be greater than 0" --rules "Dates must be in 2026"
```

**Output:**
```json
{
  "file": "report.xlsx",
  "rules_checked": 2,
  "violations": [
    {
      "rule": "Revenue must be greater than 0",
      "cell": "Summary!B12",
      "value": -500,
      "severity": "error",
      "message": "Negative revenue violates business rule"
    }
  ],
  "warnings": [],
  "passed": false
}
```

**Rule syntax:**
- `column <name> must be <condition>` (e.g., `Revenue must be > 0`)
- `column <name> must be <type>` (e.g., `Date must be date`)
- `no duplicates in <range>`
- `sum of <range> must equal <cell>`

---

## Macro Recorder

Records manual Excel actions as `.xlf` scripts. Transforms user expertise into automation.

### record start

Starts recording your manual Excel actions as a `.xlf` script.

```bash
xlforge record start <file.xlsx>
xlforge record start <file.xlsx> --output script.xlf
xlforge record start <file.xlsx> --clean   # Normalize commands on stop
```

**Implementation:** Hooks into Excel's `AppEvents` (SheetChange, WindowSelectionChange). The recorder monitors:
- Cell selections
- Value changes
- Formatting operations
- Sheet navigation

**Watchdog Timer:** If the CLI loses connection to Excel for > 5 seconds, the recorder auto-detaches hooks to keep the Excel session stable.

```bash
xlforge record start report.xlsx --watchdog 10s
# If connection lost for 10s, auto-stop recording
```

---

### record --interactive (Teacher Mode)

As you record, the CLI prints the command it *would* have generated:

```
[Recording] Human bolds cell A1
[Generated] xlforge format cell "Summary!A1" --bold

[Recording] Human sets B2 = "100"
[Generated] xlforge cell set "Summary!B2" "100"

[Recording] Human applies currency format
[Generated] xlforge format cell "Summary!B2" --number-format currency
```

**Use case:** Teaches users CLI syntax while they work.

---

### record stop

Stops recording and generates the `.xlf` script.

```bash
xlforge record stop
xlforge record stop --output daily_report.xlf
```

**Command Normalization:** The stop command cleans up noisy recordings:
- Removes redundant navigation clicks (B2 → C2 → B2 becomes just B2 operation)
- Merges consecutive format calls on same cell
- Eliminates dead-end operations
- Outputs idempotent, clean commands

**Output:**
```bash
xlforge record stop --clean
# Saved to script.xlf (42 commands, normalized from 156 recorded actions)
```

---

### record status

Shows if recording is active.

```bash
xlforge record status
```

**Output:**
```json
{
  "recording": true,
  "file": "report.xlsx",
  "duration": "00:05:23",
  "commands_captured": 42,
  "watchdog_active": true
}
```

---

### record pause / resume

Temporarily pause recording without stopping.

```bash
xlforge record pause
# ... human does non-relevant work ...
xlforge record resume
```

---

## Complete Examples

### Agent discovers and updates a value
```bash
# Agent queries semantic index
COORD=$(xlforge query report.xlsx "Total Revenue 2025" --coordinate)

# Agent reads current value
xlforge cell get $COORD --json

# Agent updates the value
xlforge cell set $COORD "5000000"

# Verify
xlforge cell get $COORD --json
```

### LLM generates a report with full context
```bash
# Get schema for system prompt
xlforge describe report.xlsx "Data!A1:Z100000" --schema-only --json > schema.json

# Get markdown description for analysis
xlforge describe report.xlsx "Summary!A1:F50" --include-objects --json > summary.json

# Run report generation
xlforge run generate_report.xlf
```

### Corporate secure indexing
```bash
# Create local index (no cloud)
xlforge index create report.xlsx --engine local --provider ollama --privacy-check

# Watch for real-time updates
xlforge index watch report.xlsx &

# Query
xlforge query report.xlsx "Where is the net profit?" --min-confidence 0.85 --json
```

### Learning mode for new users
```bash
# Start interactive recording (teacher mode)
xlforge record start report.xlsx --interactive

# User does work in Excel...
# CLI prints generated commands

# User reviews cleaned script
xlforge record stop --output learned_commands.xlf
cat learned_commands.xlf
```

---

## Error Codes

| Code | Meaning |
|------|---------|
| `80` | Index not found (run `index create` first) |
| `81` | LLM provider error (quota, network) |
| `82` | Privacy check failed (sensitive data detected) |
| `83` | Confidence below threshold |
| `84` | Recording already active |
| `85` | No active recording to stop |
| `86` | Watchdog timeout (Excel connection lost) |
| `87` | Invalid rule syntax for semantic-check |
