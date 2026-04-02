# OOP Architecture for xlforge

xlforge follows object-oriented design principles using Python's advanced features: data classes, protocols, decorators, and structural typing.

## Core Principles

1. **Value Objects** - Immutable, equality by value, no side effects
2. **Entities** - Objects with identity, mutable state, side effects
3. **Services** - Stateless operations, dependency injection
4. **Strategy Pattern** - Pluggable engine implementations
5. **Composition over inheritance** - Favor object composition

---

## Value Objects

Immutable objects that represent single values. Used for typed IDs, references, and domain concepts.

### CellRef

Represents a cell reference in `Sheet!A1` notation.

```python
@dataclass(frozen=True, slots=True)
class CellRef:
    """Immutable cell reference like 'Data!A1' or 'Sheet1!B2:C10'."""

    sheet: str
    coord: str

    @cached_property
    def row(self) -> int:
        """Zero-based row index."""
        return cell_ref_to_row(self.coord)

    @cached_property
    def col(self) -> int:
        """Zero-based column index."""
        return cell_ref_to_col(self.coord)

    @cached_property
    def is_range(self) -> bool:
        """Check if this is a range reference."""
        return ":" in self.coord

    def to_a1_notation(self) -> str:
        """Convert to 'A1' notation."""
        return self.coord

    def __post_init__(self) -> None:
        validate_cell_ref(self.coord)

    def __str__(self) -> str:
        if self.sheet:
            return f"{self.sheet}!{self.coord}"
        return self.coord
```

**Usage:**
```python
ref = CellRef(sheet="Data", coord="A1")
assert ref.row == 0
assert ref.col == 0
assert str(ref) == "Data!A1"

range_ref = CellRef(sheet="Data", coord="A1:C3")
assert range_ref.is_range
```

### CellValue

Represents a typed cell value with coercion helpers.

```python
@dataclass(frozen=True, slots=True)
class CellValue:
    """Immutable cell value with type information."""

    raw: Any
    type: ValueType

    def as_string(self) -> str:
        return str(self.raw)

    def as_number(self) -> float:
        if self.type == ValueType.NUMBER:
            return float(self.raw)
        raise TypeError(f"Cannot convert {self.type} to number")

    def as_bool(self) -> bool:
        if self.type == ValueType.BOOL:
            return bool(self.raw)
        raise TypeError(f"Cannot convert {self.type} to bool")

    def as_date(self) -> datetime:
        if self.type == ValueType.DATE:
            return self.raw
        raise TypeError(f"Cannot convert {self.type} to date")

    @classmethod
    def from_python(cls, value: Any) -> CellValue:
        """Create CellValue from a Python value."""
        type_ = infer_value_type(value)
        return cls(raw=value, type=type_)

    @classmethod
    def from_string(cls, value: str, type_hint: ValueType | None = None) -> CellValue:
        """Create CellValue from string with optional type hint."""
        if type_hint:
            coerced = coerce_to_type(value, type_hint)
            return cls(raw=coerced, type=type_hint)
        type_ = infer_type_from_string(value)
        coerced = coerce_to_type(value, type_)
        return cls(raw=coerced, type=type_)
```

### ValueType (Enum)

```python
class ValueType(Enum):
    STRING = "string"
    NUMBER = "number"
    BOOL = "bool"
    DATE = "date"
    FORMULA = "formula"
    EMPTY = "empty"
    ERROR = "error"
```

---

## Entities

Objects with identity and mutable state. These are the core domain objects.

### Workbook

Represents an Excel workbook with engine abstraction.

```python
class Workbook:
    """Excel workbook entity with engine abstraction."""

    def __init__(
        self,
        path: Path,
        engine: Engine | None = None,
        *,
        read_only: bool = False,
        data_only: bool = True,
    ):
        self._path = path
        self._engine = engine or EngineSelector.for_path(path)
        self._read_only = read_only
        self._data_only = data_only
        self._is_open = False
        self._sheets: list[Sheet] = []

    @property
    def path(self) -> Path:
        return self._path

    @property
    def engine(self) -> Engine:
        return self._engine

    @property
    def is_open(self) -> bool:
        return self._is_open

    def open(self) -> Self:
        """Open the workbook."""
        if self._is_open:
            return self
        self._engine.open(self._path, read_only=self._read_only, data_only=self._data_only)
        self._is_open = True
        self._load_sheets()
        return self

    def close(self) -> None:
        """Close the workbook."""
        if not self._is_open:
            return
        self._engine.close(self._path)
        self._is_open = False
        self._sheets.clear()

    def __enter__(self) -> Self:
        return self.open()

    def __exit__(self, *args: Any) -> None:
        self.close()

    @property
    def sheets(self) -> list[Sheet]:
        """Get all sheets."""
        if not self._is_open:
            raise RuntimeError("Workbook not open. Call open() first.")
        return self._sheets.copy()

    def sheet(self, name: str) -> Sheet:
        """Get a sheet by name."""
        if not self._is_open:
            raise RuntimeError("Workbook not open. Call open() first.")
        for s in self._sheets:
            if s.name == name:
                return s
        raise SheetNotFoundError(name)

    def _load_sheets(self) -> None:
        sheet_names = self._engine.list_sheets(self._path)
        self._sheets = [Sheet(name=n, workbook=self) for n in sheet_names]


class SheetNotFoundError(XlforgeError):
    def __init__(self, sheet_name: str):
        super().__init__(
            code=ErrorCode.SHEET_NOT_FOUND,
            message=f"Sheet '{sheet_name}' not found",
            details={"sheet": sheet_name},
        )
```

**Usage:**
```python
wb = Workbook(Path("report.xlsx"))

# Context manager pattern
with Workbook(Path("report.xlsx")) as wb:
    sheet = wb.sheet("Data")
    cell = sheet.cell(CellRef(sheet="Data", coord="A1"))
    print(cell.as_string())

# Or explicit open/close
wb = Workbook(Path("report.xlsx"))
wb.open()
try:
    data = wb.sheet("Data").range("A1:C10").values()
finally:
    wb.close()
```

### Sheet

```python
@dataclass
class Sheet:
    """Sheet entity with cell access."""

    name: str
    workbook: Workbook

    def cell(self, ref: CellRef | str) -> CellValue:
        """Get cell value."""
        if isinstance(ref, str):
            ref = CellRef(sheet=self.name, coord=ref)
        return self.workbook.engine.get_cell(self.name, ref.coord)

    def set_cell(self, ref: CellRef | str, value: CellValue | Any) -> None:
        """Set cell value."""
        if isinstance(ref, str):
            ref = CellRef(sheet=self.name, coord=ref)
        if not isinstance(value, CellValue):
            value = CellValue.from_python(value)
        self.workbook.engine.set_cell(self.name, ref.coord, value)

    def range(self, coord: str) -> Range:
        """Get a range for bulk operations."""
        return Range(sheet=self, coord=coord)

    @property
    def is_protected(self) -> bool:
        return self.workbook.engine.is_sheet_protected(self.name)

    @property
    def used_range(self) -> Range:
        """Get the used range of this sheet."""
        dimensions = self.workbook.engine.get_sheet_dimensions(self.name)
        return Range(sheet=self, coord=dimensions)
```

### Range

```python
@dataclass
class Range:
    """Represents a cell range with bulk operations."""

    sheet: Sheet
    coord: str

    @cached_property
    def cell_ref(self) -> CellRef:
        return CellRef(sheet=self.sheet.name, coord=self.coord)

    @cached_property
    def values(self) -> list[list[CellValue]]:
        """Get all values in range as 2D array."""
        return self.sheet.workbook.engine.get_range(self.sheet.name, self.coord)

    def clear(self) -> None:
        """Clear all cells in range."""
        self.sheet.workbook.engine.clear_range(self.sheet.name, self.coord)

    def copy_to(self, dest: CellRef) -> None:
        """Copy range to destination."""
        self.sheet.workbook.engine.copy_range(
            self.sheet.name, self.coord,
            dest.sheet, dest.coord,
        )
```

---

## Engine (Strategy Pattern)

Abstract engine interface with multiple implementations.

### Engine Protocol

```python
class Engine(Protocol):
    """Protocol for workbook engines."""

    def open(self, path: Path, *, read_only: bool = False, data_only: bool = True) -> None:
        """Open a workbook."""
        ...

    def close(self, path: Path) -> None:
        """Close a workbook."""
        ...

    def list_sheets(self, path: Path) -> list[str]:
        """List all sheet names."""
        ...

    def get_cell(self, sheet: str, coord: str) -> CellValue:
        """Get cell value."""
        ...

    def set_cell(self, sheet: str, coord: str, value: CellValue) -> None:
        """Set cell value."""
        ...

    def get_range(self, sheet: str, coord: str) -> list[list[CellValue]]:
        """Get range values as 2D array."""
        ...

    def set_range(self, sheet: str, coord: str, values: list[list[Any]]) -> None:
        """Set range values from 2D array."""
        ...
```

### Engine Implementations

```python
class OpenpyxlEngine:
    """Openpyxl-based engine for headless environments."""

    def __init__(self) -> None:
        self._workbooks: dict[Path, openpyxl.Workbook] = {}

    def open(self, path: Path, *, read_only: bool = False, data_only: bool = True) -> None:
        wb = openpyxl.load_workbook(path, read_only=read_only, data_only=data_only)
        self._workbooks[path] = wb

    def close(self, path: Path) -> None:
        if path in self._workbooks:
            self._workbooks[path].close()
            del self._workbooks[path]

    def get_cell(self, sheet: str, coord: str) -> CellValue:
        wb = self._workbooks[path]
        ws = wb[sheet]
        cell = ws[coord]
        return CellValue(raw=cell.value, type=infer_value_type(cell.value))

    # ... other methods


class XlwingsEngine:
    """xlwings-based engine for full Excel integration."""

    def __init__(self) -> None:
        self._instances: dict[Path, xlwings.Workbook] = {}

    def open(self, path: Path, *, read_only: bool = False, data_only: bool = True) -> None:
        app = xlwings.App()
        wb = app.books.open(path, read_only=read_only)
        self._instances[path] = wb

    def close(self, path: Path) -> None:
        if path in self._instances:
            self._instances[path].close()
            del self._instances[path]

    # ... other methods with COM management
```

### Engine Selector

```python
class EngineSelector:
    """Select appropriate engine based on environment."""

    @classmethod
    def for_path(cls, path: Path) -> Engine:
        """Select engine for a file path."""
        if cls._is_excel_available():
            return XlwingsEngine()
        return OpenpyxlEngine()

    @classmethod
    def for_engine_name(cls, name: str) -> Engine:
        """Select engine by name."""
        engines = {
            "xlwings": XlwingsEngine,
            "openpyxl": OpenpyxlEngine,
        }
        if name not in engines:
            raise ValueError(f"Unknown engine: {name}")
        return engines[name]()

    @staticmethod
    def _is_excel_available() -> bool:
        """Check if Excel is available."""
        try:
            import xlwings
            xlwings.App()
            return True
        except Exception:
            return False
```

---

## Domain Services

Stateless services that encapsulate business logic.

### CellService

```python
class CellService:
    """Service for cell operations."""

    def __init__(self, engine: Engine):
        self._engine = engine

    def get(self, path: Path, cell_ref: CellRef) -> CellValue:
        """Get cell value with full validation."""
        self._validate_file_exists(path)
        self._validate_sheet_exists(path, cell_ref.sheet)
        return self._engine.get_cell(cell_ref.sheet, cell_ref.coord)

    def set(
        self,
        path: Path,
        cell_ref: CellRef,
        value: Any,
        type_hint: ValueType | None = None,
    ) -> None:
        """Set cell value with type coercion."""
        self._validate_file_exists(path)
        self._validate_sheet_exists(path, cell_ref.sheet)

        cell_value = CellValue.from_string(str(value), type_hint)
        self._engine.set_cell(cell_ref.sheet, cell_ref.coord, cell_value)

    def bulk_set(
        self,
        path: Path,
        pattern: str,
        value: Any,
        *,
        filter_fn: Callable[[CellValue], bool] | None = None,
    ) -> int:
        """Bulk set cells matching pattern. Returns count of modified cells."""
        sheet, coord_pattern = parse_cell_ref(pattern)
        cells = self._expand_pattern(sheet, coord_pattern, filter_fn)

        count = 0
        for cell in cells:
            self.set(path, cell, value)
            count += 1
        return count
```

### FileService

```python
class FileService:
    """Service for file operations."""

    def __init__(self, engine: Engine):
        self._engine = engine

    def open(self, path: Path, **kwargs: Any) -> Workbook:
        """Open workbook with selected engine."""
        self._validate_file_exists(path)
        return Workbook(path=path, engine=self._engine, **kwargs)

    def create(self, path: Path) -> Workbook:
        """Create new workbook."""
        wb = Workbook(path=path, engine=self._engine)
        wb.open()
        return wb

    def info(self, path: Path) -> WorkbookInfo:
        """Get workbook information."""
        self._validate_file_exists(path)
        wb = self._engine.open(path, read_only=True)
        try:
            sheets = wb.sheet_names
            return WorkbookInfo(
                path=path,
                sheets=sheets,
                engine=type(self._engine).__name__,
            )
        finally:
            wb.close()
```

---

## Factory Pattern

```python
class WorkbookFactory:
    """Factory for creating workbook instances."""

    @staticmethod
    def open(
        path: str | Path,
        engine: str | Engine | None = None,
        **kwargs: Any,
    ) -> Workbook:
        """Open existing workbook."""
        path = Path(path)

        if engine is None:
            eng = EngineSelector.for_path(path)
        elif isinstance(engine, str):
            eng = EngineSelector.for_engine_name(engine)
        else:
            eng = engine

        return Workbook(path=path, engine=eng, **kwargs).open()

    @staticmethod
    def create(
        path: str | Path,
        engine: str | Engine | None = None,
    ) -> Workbook:
        """Create new workbook."""
        path = Path(path)

        if engine is None:
            eng = EngineSelector.for_path(path)
        elif isinstance(engine, str):
            eng = EngineSelector.for_engine_name(engine)
        else:
            eng = engine

        wb = Workbook(path=path, engine=eng)
        wb.open()
        return wb
```

---

## Error Handling with Result Type

```python
class CellOperationResult(Result[CellValue, XlforgeError]):
    """Result type for cell operations."""

    @classmethod
    def success(cls, value: CellValue) -> CellOperationResult:
        return Ok(value)

    @classmethod
    def failure(cls, code: ErrorCode, details: dict | None = None) -> CellOperationResult:
        return Err(XlforgeError(code=code, details=details))


class CellServiceWithResult:
    """Cell service using Result for error handling."""

    def get(self, path: Path, cell_ref: CellRef) -> CellOperationResult:
        try:
            self._validate_file_exists(path)
            value = self._engine.get_cell(cell_ref.sheet, cell_ref.coord)
            return CellOperationResult.success(value)
        except FileNotFoundError:
            return CellOperationResult.failure(
                ErrorCode.FILE_NOT_FOUND,
                {"path": str(path)},
            )
        except XlforgeError as e:
            return CellOperationResult.failure(e.code, e.details)
```

---

## CLI Integration

```python
# commands/cell.py
class CellCommands:
    """Cell command group with OOP service injection."""

    def __init__(self, service: CellService | None = None):
        self._service = service or CellService(EngineSelector.for_path(Path.cwd()))

    def get(self, file: str, cell: str, json: bool = False) -> None:
        """Get cell value."""
        path = Path(file)
        ref = CellRef(sheet="", coord=cell)

        result = self._service.get(path, ref)

        if is_err(result):
            if json:
                echo(json.dumps(result.error.to_dict()))
            else:
                echo(f"Error: {result.error}", err=True)
            raise SystemExit(result.error.code)
        else:
            if json:
                echo(json.dumps({"value": result.value.raw, "type": result.value.type}))
            else:
                echo(result.value.raw)


# Entry point wiring
def create_app() -> Typer:
    app = Typer()

    # Dependency injection
    engine = EngineSelector.for_path(Path.cwd())
    cell_service = CellService(engine)
    cell_commands = CellCommands(cell_service)

    @app.command()
    def get(file: str, cell: str, json: bool = False):
        cell_commands.get(file, cell, json)

    return app
```

---

## Project Structure (OOP)

```
xlforge/
├── core/
│   ├── __init__.py
│   ├── types/
│   │   ├── __init__.py
│   │   ├── result.py          # Result/Ok/Err
│   │   └── error.py          # ErrorCode, XlforgeError
│   │
│   ├── value_objects/         # Immutable domain values
│   │   ├── __init__.py
│   │   ├── cell_ref.py       # CellRef value object
│   │   ├── cell_value.py     # CellValue with type
│   │   └── value_type.py     # ValueType enum
│   │
│   ├── entities/              # Domain entities
│   │   ├── __init__.py
│   │   ├── workbook.py       # Workbook entity
│   │   ├── sheet.py          # Sheet entity
│   │   └── range.py          # Range entity
│   │
│   ├── services/              # Domain services
│   │   ├── __init__.py
│   │   ├── cell_service.py   # Cell operations
│   │   ├── file_service.py   # File operations
│   │   └── sheet_service.py  # Sheet operations
│   │
│   ├── engines/              # Engine implementations
│   │   ├── __init__.py
│   │   ├── base.py           # Engine protocol/interface
│   │   ├── openpyxl_.py      # openpyxl implementation
│   │   ├── xlwings_.py      # xlwings implementation
│   │   └── selector.py       # EngineSelector
│   │
│   └── factories/            # Factory classes
│       ├── __init__.py
│       └── workbook_factory.py
│
├── commands/                 # CLI commands (thin, delegate to services)
│   ├── __init__.py
│   ├── cell.py
│   ├── file.py
│   └── sheet.py
│
└── __init__.py              # App entry point
```

---

## Summary

| Pattern | Usage |
|---------|-------|
| **Value Object** | `CellRef`, `CellValue`, `ValueType` |
| **Entity** | `Workbook`, `Sheet`, `Range` |
| **Service** | `CellService`, `FileService`, `SheetService` |
| **Factory** | `WorkbookFactory` |
| **Strategy** | `Engine` protocol with `OpenpyxlEngine`, `XlwingsEngine` |
| **Result Type** | Error handling with `Result[Ok, Err]` |
| **Protocol** | `Engine` structural typing |

This OOP design provides:
- **Testability**: Mock engines, services, entities independently
- **Extensibility**: Add new engines without changing domain logic
- **Type Safety**: Value objects prevent invalid states
- **Clean Architecture**: CLI depends on services, services depend on engines
