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
xlforge/
├── __init__.py              # CLI entry point, app export
├── __main__.py              # python -m xlforge entry
├── core.py                  # SDK: business logic
├── errors.py                # Error codes and exceptions
├── result.py                # Result[T, E] and Maybe[T] types
├── context.py               # Context management (active file/sheet)
├── engines/
│   ├── __init__.py
│   ├── base.py             # Engine abstract class
│   ├── xlwings.py         # xlwings implementation
│   ├── openpyxl.py        # openpyxl implementation
│   └── duckdb.py          # DuckDB SQL engine
├── commands/
│   ├── __init__.py
│   ├── file.py             # file open, save, close, info, kill
│   ├── cell.py             # cell get, set, formula, clear, copy, bulk
│   ├── sheet.py            # sheet list, create, delete, rename
│   ├── format.py           # format cell, range
│   ├── data.py             # import csv, export csv
│   ├── table.py            # table create, link, sync-schema
│   ├── chart.py            # chart create
│   ├── validation.py       # validation create
│   ├── protection.py       # freeze, protect, unprotect
│   ├── app.py              # app visible, calculate, focus, alert
│   ├── checkpoint.py       # checkpoint create, restore
│   ├── branch.py           # branch operations
│   ├── watch.py            # watch start, stop
│   ├── sql.py              # sql query, push, pull
│   └── semantic.py         # index, query, describe
└── utils/
    ├── path.py             # Path resolution utilities
    ├── cell.py             # Cell reference parsing
    └── types.py            # Type coercion utilities
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
