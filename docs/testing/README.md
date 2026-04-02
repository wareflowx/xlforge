# Testing Strategy

## Overview

xlforge uses a multi-layer testing approach to ensure reliability while keeping CI fast and Excel-independent.

## Test Layers

### 1. Unit Tests (`tests/unit/`)

Pure functions with no I/O or external dependencies.

```bash
tests/unit/
├── test_result.py            # Result[T, E] and Maybe[T] types
├── test_cell_parsing.py      # Parse "Sheet!A1" notation
├── test_type_coercion.py      # String → number, date, bool, formula
├── test_engine_selection.py    # Auto-detect xlwings vs openpyxl
├── test_error_codes.py         # Verify all 127 error codes
└── test_path_resolution.py     # Relative → absolute path handling
```

**Run:** `uv run pytest tests/unit -v`

### 2. Integration Tests (`tests/integration/`)

CLI commands with mocked file I/O and engines.

```bash
tests/integration/
├── test_file_commands.py       # file open, save, close, info, kill
├── test_cell_commands.py       # cell get, set, formula, clear, copy, bulk
├── test_sheet_commands.py      # sheet list, create, delete, rename
├── test_format_commands.py     # format cell, range, apply
├── test_data_commands.py       # import csv, export csv
├── test_global_flags.py        # --json, --json-errors, --dry-run, --engine
└── fixtures/                   # Shared test fixtures
    ├── sample.xlsx             # Simple workbook with known data
    ├── empty.xlsx              # Blank workbook for write tests
    └── corrupted.xlsx          # Invalid file for error testing
```

**Run:** `uv run pytest tests/integration -v`

### 3. E2E Tests (`tests/e2e/`)

Real Excel/xlwings execution. Only runs on Windows with Excel installed.

```bash
tests/e2e/
├── test_xlwings_commands.py    # Full xlwings integration
└── fixtures/                   # Test workbooks
```

**Run:** `uv run pytest tests/e2e -v` (requires Excel)

## Running Tests

```bash
# All tests except E2E (default for CI)
uv run pytest tests/ -v

# Only unit tests
uv run pytest tests/unit -v

# Only integration tests
uv run pytest tests/integration -v

# Only E2E tests
uv run pytest tests/e2e -v

# With coverage
uv run pytest tests/ --cov=xlforge --cov-report=html
```

## Fixtures

### Built-in Fixtures

| Fixture | Scope | Description |
|---------|-------|-------------|
| `runner` | session | Pre-configured `CliRunner` instance |
| `tmp_path` | function | Temporary directory for file operations |
| `sample_xlsx` | function | Path to `tests/integration/fixtures/sample.xlsx` |

### Creating Fixtures

```python
# tests/conftest.py
import pytest
from typer.testing import CliRunner

from xlforge import app

@pytest.fixture(scope="session")
def runner():
    return CliRunner()

@pytest.fixture
def sample_xlsx(tmp_path):
    # Create a test workbook
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Header"
    ws["B1"] = 42
    path = tmp_path / "sample.xlsx"
    wb.save(path)
    return path
```

## Testing Commands

### Pattern: Happy Path

```python
def test_cell_get_json_output(runner, sample_xlsx):
    result = runner.invoke(app, ["cell", "get", str(sample_xlsx), "Data!A1", "--json"])
    assert result.exit_code == 0
    data = json.loads(result.output)
    assert data["cell"] == "Data!A1"
    assert data["value"] == "Header"
```

### Pattern: Error Cases

```python
def test_cell_get_file_not_found(runner):
    result = runner.invoke(app, ["cell", "get", "nonexistent.xlsx", "Sheet1!A1"])
    assert result.exit_code == 2  # Error code 2: File not found

def test_cell_get_sheet_not_found(runner, sample_xlsx):
    result = runner.invoke(app, ["cell", "get", str(sample_xlsx), "NonExistent!A1"])
    assert result.exit_code == 3  # Error code 3: Sheet not found
```

### Pattern: Global Flags

```python
def test_cell_get_with_json_flag(runner, sample_xlsx):
    result = runner.invoke(app, ["cell", "get", str(sample_xlsx), "Data!A1", "--json"])
    assert result.exit_code == 0
    # Verify JSON output structure
    data = json.loads(result.output)
    assert "cell" in data
    assert "value" in data

def test_cell_get_with_json_errors(runner):
    result = runner.invoke(app, ["cell", "get", "nonexistent.xlsx", "A1", "--json-errors"])
    assert result.exit_code == 2
    error = json.loads(result.output)
    assert error["success"] is False
    assert error["code"] == 2
```

### Pattern: Dry Run

```python
def test_file_save_dry_run(runner, sample_xlsx):
    result = runner.invoke(app, ["file", "save", str(sample_xlsx), "--dry-run"])
    assert result.exit_code == 0
    # File should not be modified
```

## Mocking Engines

### Unit Test Mocking

```python
# tests/unit/test_engine_selection.py
from unittest.mock import patch

def test_auto_select_xlwings_when_excel_installed():
    with patch("xlforge.core.find_excel_executable", return_value="C:/Program Files..."):
        engine = select_engine()
        assert engine == "xlwings"

def test_auto_select_openpyxl_in_docker():
    with patch("xlforge.core.find_excel_executable", return_value=None):
        engine = select_engine()
        assert engine == "openpyxl"
```

### Integration Test Mocking

```python
# tests/integration/test_file_commands.py
from unittest.mock import patch, MagicMock

def test_file_open_with_xlwings_engine(runner, tmp_path):
    mock_wb = MagicMock()
    with patch("xlforge.engines.xlwings.Workbook", return_value=mock_wb):
        xlsx = tmp_path / "test.xlsx"
        result = runner.invoke(app, ["file", "open", str(xlsx), "--engine", "xlwings"])
        assert result.exit_code == 0
        mock_wb.open.assert_called_once()
```

## Test Data

### Sample Workbooks

Create test workbooks programmatically:

```python
@pytest.fixture
def workbook_with_data(tmp_path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"

    # Headers
    ws["A1"] = "Name"
    ws["B1"] = "Value"
    ws["C1"] = "Date"

    # Data rows
    ws["A2"] = "Alice"
    ws["B2"] = 100
    ws["C2"] = "2026-01-15"

    ws["A3"] = "Bob"
    ws["B3"] = 200
    ws["C3"] = "2026-01-16"

    path = tmp_path / "data.xlsx"
    wb.save(path)
    return path
```

## CI Configuration

```yaml
# .github/workflows/ci.yml
name: Tests

on: [push, pull_request]

jobs:
  test:
    runs-on: windows-latest  # Windows required for xlwings
    steps:
      - uses: actions/checkout@v4
      - uses: astral-sh/setup-uv@v4
      - run: uv sync
      - run: uv run pytest tests/unit tests/integration -v
      - run: uv run pytest tests/e2e -v --require-excel  # Only if Excel available
```

## Coverage Goals

| Layer | Target |
|-------|--------|
| Unit | 90%+ |
| Integration | 80%+ |
| E2E | 70%+ |
| **Overall** | **85%+** |

## Error Code Testing

All 127 error codes should be tested. Use parametrized tests:

```python
@pytest.mark.parametrize("code,description", [
    (0, "Success"),
    (2, "File not found"),
    (3, "Sheet not found"),
    (4, "Cell not found"),
    # ... all 127 codes
])
def test_error_code_exists(code, description):
    from xlforge.errors import ErrorCode
    assert ErrorCode(code).name is not None
    assert ErrorCode(code).value == code
```

## Testing Result and Maybe Types

See [Functional Patterns](../ARCHITECTURE.md#functional-patterns) for type definitions.

### Testing Result[T, E]

```python
# tests/unit/test_result.py
import pytest
from xlforge.result import Ok, Err, is_ok, is_err

def test_ok_unwrap():
    result: Result[int, str] = Ok(42)
    assert result.unwrap() == 42

def test_err_unwrap_raises():
    result = Err("oops")
    with pytest.raises(ValueError):
        result.unwrap()

def test_map_transforms_ok():
    result = Ok(42)
    mapped = result.map(lambda x: x * 2)
    assert mapped == Ok(84)

def test_map_preserves_err():
    result = Err("oops")
    mapped = result.map(lambda x: x * 2)
    assert mapped == Err("oops")

def test_and_then_chains_ok():
    def parse(x: int) -> Result[str, str]:
        return Ok(f"value:{x}")

    result = Ok(42).and_then(parse)
    assert result == Ok("value:42")

def test_and_then_preserves_err():
    def parse(x: int) -> Result[str, str]:
        return Ok(f"value:{x}")

    result = Err("error").and_then(parse)
    assert result == Err("error")
```

### Testing Maybe[T]

```python
# tests/unit/test_maybe.py
import pytest
from xlforge.result import Some, Nothing, is_some, is_nothing

def test_some_unwrap():
    maybe: Maybe[int] = Some(42)
    assert maybe.unwrap() == 42

def test_nothing_unwrap_raises():
    maybe = Nothing()
    with pytest.raises(ValueError):
        maybe.unwrap()

def test_map_transforms_some():
    maybe = Some(42)
    mapped = maybe.map(lambda x: x * 2)
    assert mapped == Some(84)

def test_map_preserves_nothing():
    maybe = Nothing()
    mapped = maybe.map(lambda x: x * 2)
    assert mapped == Nothing()

def test_filter_returns_some_when_match():
    maybe = Some(42)
    filtered = maybe.filter(lambda x: x > 10)
    assert filtered == Some(42)

def test_filter_returns_nothing_when_no_match():
    maybe = Some(5)
    filtered = maybe.filter(lambda x: x > 10)
    assert filtered == Nothing()

def test_and_then_chains_some():
    def get_positive(x: int) -> Maybe[int]:
        if x > 0:
            return Some(x * 2)
        return Nothing()

    result = Some(21).and_then(get_positive)
    assert result == Some(42)

def test_and_then_preserves_nothing():
    def get_positive(x: int) -> Maybe[int]:
        if x > 0:
            return Some(x * 2)
        return Nothing()

    result = Nothing().and_then(get_positive)
    assert result == Nothing()
```

### Testing SDK Functions Returning Result

```python
# tests/unit/test_core.py
from xlforge.result import Ok, Err

def test_cell_get_returns_ok_on_success(sample_xlsx):
    result = core.cell_get(str(sample_xlsx), "Data!A1")
    assert is_ok(result)
    assert result.value.cell == "Data!A1"
    assert result.value.value == "Header"

def test_cell_get_returns_err_on_file_not_found():
    result = core.cell_get("nonexistent.xlsx", "Sheet1!A1")
    assert is_err(result)
    assert result.error == ErrorCode.FILE_NOT_FOUND

def test_cell_get_returns_err_on_sheet_not_found(sample_xlsx):
    result = core.cell_get(str(sample_xlsx), "NonExistent!A1")
    assert is_err(result)
    assert result.error == ErrorCode.SHEET_NOT_FOUND
```

## Performance Testing

For bulk operations, ensure acceptable performance:

```python
def test_cell_bulk_performance(tmp_path):
    import time
    xlsx = create_large_workbook(tmp_path, rows=10000)

    start = time.time()
    result = runner.invoke(app, ["cell", "bulk", str(xlsx), "Data!*", "--filter", "empty"])
    elapsed = time.time() - start

    assert result.exit_code == 0
    assert elapsed < 5.0  # Should complete in under 5 seconds
```
