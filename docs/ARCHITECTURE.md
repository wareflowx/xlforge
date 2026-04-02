# Architecture

xlforge follows a layered architecture: **API → SDK → CLI**.

## Layers Overview

```
┌─────────────────────────────────────────────────────────────┐
│  CLI (Typer Commands)                                      │
│  - Parse command-line arguments                             │
│  - Format output (JSON, text, errors)                        │
│  - User experience (help, completion, prompts)                │
└─────────────────────────────┬───────────────────────────────┘
                              │
┌─────────────────────────────▼───────────────────────────────┐
│  SDK (Core)                                                 │
│  - Business logic and validation                             │
│  - Orchestration of operations                               │
│  - Context management (active file, sheet)                   │
│  - Type coercion and error handling                          │
└─────────────────────────────┬───────────────────────────────┘
                              │
┌─────────────────────────────▼───────────────────────────────┐
│  API (Engines)                                               │
│  - Pure Excel interaction                                   │
│  - xlwings implementation (Windows, full features)           │
│  - openpyxl implementation (headless, cross-platform)        │
│  - DuckDB engine (SQL queries)                               │
└─────────────────────────────────────────────────────────────┘
```

## Layer Responsibilities

### CLI Layer (`xlforge/commands/`)

Typer command functions. Thin, no business logic.

```python
# xlforge/commands/cell.py
@app.command()
def set(file: str, cell: str, value: str, type: Optional[str] = None):
    """Set a cell value."""
    result = sdk.cell_set(file, cell, value, type_hint=type)
    if json_output:
        echo(json.dumps(result))
    else:
        echo(result)
```

**Responsibilities:**
- Argument parsing
- Output formatting
- Help text and CLI UX
- Error formatting

**Never:**
- Direct engine calls
- Business logic
- File path manipulation

### SDK Layer (`xlforge/core.py`)

Business logic and orchestration. Reusable by anyone.

```python
# xlforge/core.py
def cell_set(file: str, cell: str, value: str, type_hint: Optional[str] = None) -> dict:
    """Set a cell value in an Excel file."""
    # Validate inputs
    file_path = resolve_path(file)
    sheet, coord = parse_cell_reference(cell)  # "Data!A1" → "Data", "A1"

    # Select engine
    engine = get_engine(file_path)

    # Type coercion
    typed_value = coerce_value(value, type_hint)

    # Execute via engine
    result = engine.set_cell(sheet, coord, typed_value)

    return result
```

**Responsibilities:**
- Input validation
- Path resolution
- Cell reference parsing (`"Data!A1"` → components)
- Type coercion
- Engine selection
- Error mapping (code → message)
- Context management

**Never:**
- Direct xlwings/openpyxl calls
- COM manipulation

### API Layer (`xlforge/engines/`)

Pure Excel interaction implementations.

```python
# xlforge/engines/base.py
class Engine(ABC):
    @abstractmethod
    def set_cell(self, sheet: str, coord: str, value: Any) -> dict:
        pass

    @abstractmethod
    def get_cell(self, sheet: str, coord: str) -> dict:
        pass

# xlforge/engines/xlwings.py
class XlwingsEngine(Engine):
    def set_cell(self, sheet: str, coord: str, value: Any) -> dict:
        # Direct xlwings calls only
        # No business logic
        pass

# xlforge/engines/openpyxl.py
class OpenpyxlEngine(Engine):
    def set_cell(self, sheet: str, coord: str, value: Any) -> dict:
        # Direct openpyxl calls only
        # No business logic
        pass
```

**Responsibilities:**
- xlwings/openpyxl API calls
- COM process management
- File handle management
- Raw cell read/write

**Never:**
- Validation
- Error messages to user
- Type coercion

## Directory Structure

```
xlforge/                                  # Package root
│
├── __init__.py                           # CLI app entry point
│   typer.Typer() instance named `app`  # Exports: app
│
├── __main__.py                          # Entry point for `python -m xlforge`
│   from xlforge import app
│   if __name__ == "__main__":
│       app()
│
├── errors.py                            # Error codes and exceptions
│   ├── class ErrorCode(IntEnum)         # All 127 error codes
│   ├── class XlforgeError(Exception)     # Base exception with code + message
│   └── ERROR_MESSAGES: dict[int, str]   # Code → human message mapping
│
├── context.py                           # CLI context management
│   ├── class Context                    # Active file/sheet state
│   ├── DEFAULT_CONTEXT: Context         # Global default context
│   ├── get_context()                    # Get current context
│   └── set_context(file, sheet)         # Set active context
│
├── core.py                              # SDK: business logic layer
│   ├── cell_get(), cell_set()           # Cell operations
│   ├── sheet_list(), sheet_create()     # Sheet operations
│   ├── file_open(), file_save()         # File operations
│   ├── parse_cell_ref()                  # "Data!A1" → (sheet, coord)
│   ├── resolve_path()                    # Relative → absolute path
│   ├── coerce_value()                    # String → typed value
│   └── get_engine()                     # Engine selection
│
├── core/
│   └── types/
│       ├── __init__.py                   # Re-exports result types
│       └── result.py                    # Result[T, E] and Maybe[T] types
│           ├── Ok[T], Err[E]            # Result variants
│           ├── Some[T], Nothing[T]       # Maybe variants
│           └── is_ok(), is_err()        # Type guards
│
├── engines/                             # API: Excel interaction layer
    ├── __init__.py
    ├── base.py                          # Engine interface
│   │   ├── class Engine(ABC)             # Abstract base class
│   │   ├── class CellValue               # NamedTuple: value, type, formula
│   │   └── METHOD_NOT_SUPPORTED          # Error code 9 helper
│   │
│   ├── xlwings_.py                      # xlwings implementation
│   │   ├── class XlwingsEngine(Engine)
│   │   ├── WORKBOOK_CACHE: dict          # PID → workbook instance
│   │   ├── _get_or_create_workbook()     # COM session management
│   │   ├── _ensure_sheet()               # Sheet access with auto-create
│   │   └── _com_error_handler()          # COM → ErrorCode mapping
│   │
│   ├── openpyxl_.py                     # openpyxl implementation
│   │   ├── class OpenpyxlEngine(Engine)
│   │   ├── _load_workbook()              # File → Workbook with lock retry
│   │   ├── _save_workbook()              # Save with backup
│   │   └── _validate_sheet()             # Sheet existence check
│   │
│   └── duckdb.py                         # DuckDB SQL engine
│       ├── class DuckDBEngine(Engine)
│       ├── _execute_query()              # Run SQL, return results
│       └── _register_excel()             # Register .xlsx as DuckDB table
│
├── commands/                             # CLI: Typer command groups
│   ├── __init__.py                       # Command group imports
│   │
│   ├── file.py                           # file command group
│   │   ├── open(file, engine, visible)   # Open/create workbook
│   │   ├── save(file, output, dry_run)   # Save workbook
│   │   ├── close(file, force)           # Close workbook
│   │   ├── info(file, json)             # Show metadata + PID
│   │   ├── kill(file_or_pid, force)     # Kill Excel process
│   │   ├── recover(file, force)         # Kill + reopen
│   │   ├── check(file, repair)          # Health check
│   │   ├── monitor(file, timeout)        # Watch for changes
│   │   └── template(action, name, file) # Template management
│   │
│   ├── cell.py                           # cell command group
│   │   ├── get(file, cell, json, formula, calculate)  # Read cell
│   │   ├── set(file, cell, value, type)               # Write cell
│   │   ├── formula(file, cell, formula)               # Set formula
│   │   ├── clear(file, cell, format_only, value_only) # Clear
│   │   ├── copy(file, src, dst)                       # Copy cell
│   │   ├── bulk(file, pattern, filter, format, set)   # Bulk ops
│   │   ├── search(file, query, sheet, json)           # Find cell
│   │   └── fill(file, range, direction, stop)        # Auto-fill
│   │
│   ├── sheet.py                          # sheet command group
│   │   ├── list(file, json)              # List all sheets
│   │   ├── create(file, name, copy_from) # Create sheet
│   │   ├── delete(file, name)            # Delete sheet
│   │   ├── rename(file, old, new)        # Rename sheet
│   │   ├── copy(file, src, dst)         # Copy sheet
│   │   └── use(file, sheet)              # Set active sheet
│   │
│   ├── format.py                         # format command group
│   │   ├── cell(file, cell, bold, size, color, ...)
│   │   ├── range(file, range, pattern, border, ...)
│   │   └── apply(file, range, style_name)
│   │
│   ├── data.py                           # data command group
│   │   ├── import_csv(file, csv, sheet, cell, has_headers, ...)
│   │   └── export_csv(file, range, csv, headers)
│   │
│   ├── table.py                          # table command group
│   │   ├── create(file, range, csv, style, freeze_header)
│   │   ├── link(file, range, db_url, table, mode, key_col)
│   │   ├── sync_schema(file, range, strict, prune)
│   │   └── refresh(file, range)
│   │
│   ├── chart.py                          # chart command group
│   │   └── create(file, range, csv, type, x, y, title, ...)
│   │
│   ├── validation.py                     # validation command group
│   │   └── create(file, range, type, formula1, formula2, ...)
│   │
│   ├── protection.py                     # protection command group
│   │   ├── freeze(file, cell)           # Freeze panes
│   │   ├── protect(file, sheet, password)
│   │   └── unprotect(file, sheet, password)
│   │
│   ├── app.py                            # app command group
│   │   ├── visible(file, visible)       # Show/hide Excel
│   │   ├── calculate(file, mode)         # Force recalc
│   │   ├── focus(file, sheet)           # Activate window
│   │   ├── alert(file, message, buttons)# Show dialog
│   │   ├── wait_idle(file, timeout)     # Wait for Excel idle
│   │   └── screen_update(file, enable) # Enable/disable screen
│   │
│   ├── checkpoint.py                      # checkpoint command group
│   │   ├── create(file, message)         # Create checkpoint
│   │   ├── list(file, json)             # List checkpoints
│   │   ├── restore(file, checkpoint_id) # Restore checkpoint
│   │   └── delete(file, checkpoint_id)  # Delete checkpoint
│   │
│   ├── branch.py                          # branch command group
│   │   ├── create(file, name)            # Create branch
│   │   ├── list(file, json)             # List branches
│   │   ├── checkout(file, name)          # Switch branch
│   │   ├── merge(file, name)             # Merge branch
│   │   └── delete(file, name)           # Delete branch
│   │
│   ├── watch.py                          # watch command group
│   │   ├── start(file, commands)         # Start watching
│   │   └── stop(file)                   # Stop watching
│   │
│   ├── sql.py                            # sql command group
│   │   ├── query(query, to, db)          # Execute SQL query
│   │   ├── push(query, db, to, format)  # Push to Excel
│   │   ├── pull(file, range, db, table) # Pull from Excel
│   │   └── connect(name, url)            # Register DB connection
│   │
│   └── semantic.py                       # semantic command group
│       ├── create_index(file, engine, privacy_check)
│       ├── query(file, query, coordinate)
│       └── describe(file, range, schema_only, json)
│
└── utils/                                # Shared utilities
    ├── __init__.py
    ├── path.py                           # Path utilities
    │   ├── resolve_path(path)             # → absolute path
    │   ├── normalize_path(path)          # → / separators
    │   └── is_absolute_path(path)        # → bool
    │
    ├── cell.py                           # Cell reference utilities
    │   ├── parse_cell_ref(ref)           # "Data!A1" → CellRef
    │   ├── cell_ref_to_row_col(ref)      # "A1" → (0, 0)
    │   ├── row_col_to_cell_ref(row, col) # → "B2"
    │   ├── range_to_coords(range_str)    # "A1:C3" → coords
    │   └── expand_range(range_str)       # "A:*" → ["A1", "A2", ...]
    │
    └── types.py                          # Type coercion utilities
        ├── infer_type(value)             # → "string" | "number" | "date" | "bool"
        ├── coerce_to_type(value, type_)  # → typed value
        ├── parse_date(value)             # → datetime | None
        ├── parse_number(value)           # → float | None
        └── EXCEL_DATE_FORMAT             # ISO → Excel serial
```

## Module Dependencies

```
__init__.py
    └── app (typer.Typer)

__main__.py
    └── imports __init__.app

xlforge/core/types/result.py           # No dependencies (types only)
    └── Used by: core, engines, commands

errors.py                              # No dependencies
    └── Used by: core, engines, commands

context.py
    └── xlforge/core/types/result.py (Maybe types)

core.py
    ├── result.py
    ├── errors.py
    ├── context.py
    ├── utils/path.py
    ├── utils/cell.py
    ├── utils/types.py
    └── engines/

engines/
    ├── base.py
    │   └── errors.py
    ├── xlwings.py
    │   ├── base.py
    │   └── errors.py
    ├── openpyxl.py
    │   ├── base.py
    │   └── errors.py
    └── duckdb.py
        └── base.py

commands/                    # Each command module
    ├── core.py
    ├── errors.py
    ├── result.py
    ├── context.py
    └── engines/

utils/
    └── (no dependencies, pure functions)
```

## Testing Strategy

Each layer is tested independently:

```
tests/
├── unit/                   # SDK and API unit tests
│   ├── test_core.py        # SDK business logic
│   ├── test_engines.py     # Engine implementations
│   └── test_utils.py       # Utilities
├── integration/            # CLI + SDK integration
│   ├── test_commands.py    # CLI commands with mocked SDK
│   └── test_file_commands.py
└── e2e/                    # Full stack with real Excel
    └── test_xlwings.py
```

**Key principle:** CLI tests mock the SDK. SDK tests mock the engines.

## Data Flow

### Example: `xlforge cell set report.xlsx "Data!A1" "Hello" --type string`

```
1. CLI Layer (commands/cell.py)
   └─> runner.invoke(app, ["cell", "set", "report.xlsx", "Data!A1", "Hello", "--type", "string"])

2. SDK Layer (core.py)
   └─> core.cell_set("report.xlsx", "Data!A1", "Hello", type_hint="string")
       ├─ resolve_path("report.xlsx") → "C:/Users/name/report.xlsx"
       ├─ parse_cell_reference("Data!A1") → sheet="Data", coord="A1"
       ├─ coerce_value("Hello", "string") → typed_value="Hello"
       ├─ engine = get_engine("C:/Users/name/report.xlsx") → OpenpyxlEngine (no Excel)
       └─ engine.set_cell("Data", "A1", "Hello") → {success: true}

3. API Layer (engines/openpyxl.py)
   └─> OpenpyxlEngine.set_cell("Data", "A1", "Hello")
       ├─ openpyxl.load_workbook("C:/Users/name/report.xlsx")
       ├─ ws = wb["Data"]
       ├─ ws["A1"] = "Hello"
       └─ wb.save()
```

## Engine Selection

Engines are selected automatically based on environment:

```python
# xlforge/core.py
def get_engine(file_path: str) -> Engine:
    if is_excel_installed():
        return XlwingsEngine()
    else:
        return OpenpyxlEngine()
```

| Engine | When Used | Capabilities |
|--------|-----------|--------------|
| **xlwings** | Excel installed | Full: macros, charts, formatting, live interaction |
| **openpyxl** | Headless/Linux | Read/write cells, sheet ops, basic formatting |

Force with `--engine <name>` global flag.

## Error Handling

Errors flow upward through layers:

```
API Layer: xlwings raises COMError
    ↓ mapped to ErrorCode.COM_ERROR (7)
SDK Layer: catches, wraps in XlforgeError
    ↓
CLI Layer: catches, formats as JSON or user message
```

Each layer adds context without losing the original error code.

## Functional Patterns

xlforge uses `Result[T, E]` and `Maybe[T]` types for explicit error handling instead of exceptions.

### Result[T, E]

Represents either a success (`Ok`) or a failure (`Err`).

```python
from xlforge.result import Result, Ok, Err, is_ok, is_err

def cell_get(file: str, cell: str) -> Result[CellValue, ErrorCode]:
    if not file_exists(file):
        return Err(ErrorCode.FILE_NOT_FOUND)
    if not sheet_exists(file, cell.sheet):
        return Err(ErrorCode.SHEET_NOT_FOUND)

    value = engine.get_cell(sheet, cell.coord)
    return Ok(CellValue(value=value, type=infer_type(value)))
```

**Usage:**

```python
result = cell_get("report.xlsx", "Data!A1")

if is_ok(result):
    print(result.value)  # CellValue

if is_err(result):
    print(result.error)  # ErrorCode.FILE_NOT_FOUND

# Chaining
result.map(lambda v: v.upper()).unwrap_or("N/A")

# Error propagation
def process(file: str) -> Result[ProcessedData, ErrorCode]:
    return cell_get(file, "A1").and_then(validate_value)
```

### Maybe[T]

Represents an optional value (`Some`) or absence (`Nothing`).

```python
from xlforge.result import Maybe, Some, Nothing, is_some, is_nothing

def find_config(key: str) -> Maybe[str]:
    for k, v in CONFIG.items():
        if k == key:
            return Some(v)
    return Nothing()
```

**Usage:**

```python
config = find_config("timeout")

if is_some(config):
    print(config.value)

# With default
timeout = find_config("timeout").unwrap_or(30)

# Transform
find_config("host").map(lambda h: f"{h}:{port}")
```

### Type Benefits

| Benefit | Description |
|---------|-------------|
| **Explicit** | Every error path is visible in the type signature |
| **Composable** | `and_then`, `map` for chaining operations |
| **Testable** | Easy to assert on `Ok` vs `Err` branches |
| **No exceptions** | No `try/except` clutter in business logic |

### Error Codes

All errors use the documented error codes (0-127). See [Error Codes](../reference.md#error-codes).

```python
class ErrorCode(IntEnum):
    SUCCESS = 0
    GENERAL_ERROR = 1
    FILE_NOT_FOUND = 2
    SHEET_NOT_FOUND = 3
    CELL_NOT_FOUND = 4
    # ... (all 127 codes)
```

## Why This Architecture?

1. **Separation of concerns** - CLI knows nothing about Excel APIs
2. **Testability** - Mock any layer independently
3. **Reusability** - SDK can be imported as a library
4. **Extensibility** - Add new engines (e.g., xlsb, Google Sheets) without changing CLI
5. **Maintainability** - Changes in one layer rarely affect others
