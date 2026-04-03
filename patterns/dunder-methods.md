# Maximize Dunder Methods

**Rule:** Use Python dunder methods (`__str__`, `__repr__`, `__eq__`, `__len__`, etc.) wherever they make sense.

**Rationale:**
- Provides idiomatic Python interfaces for objects
- Makes objects behave like built-in types
- Improves readability and usability in Python code
- Enables integration with Python's data model (collections, iteration, etc.)

**Required dunders by type:**

### Value Objects (frozen dataclasses)
```python
@dataclass(frozen=True)
class CellRef:
    sheet: str
    coord: str

    def __str__(self) -> str:
        """Human-readable representation."""
        if self.sheet:
            return f"{self.sheet}!{self.coord}"
        return self.coord

    def __repr__(self) -> str:
        """Unambiguous representation for debugging."""
        return f"CellRef(sheet={self.sheet!r}, coord={self.coord!r})"

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, CellRef):
            return NotImplemented
        return self.sheet == other.sheet and self.coord == other.coord

    def __hash__(self) -> int:
        return hash((self.sheet, self.coord))
```

### Entities (mutable)
```python
class Workbook:
    def __init__(self, path: Path):
        self._path = path

    def __str__(self) -> str:
        return str(self._path)

    def __repr__(self) -> str:
        return f"Workbook(path={self._path!r})"

    def __bool__(self) -> bool:
        return self._is_open

    def __enter__(self) -> Workbook:
        return self.open()

    def __exit__(self, *args: Any) -> None:
        self.close()
```

### Collection-like objects
```python
class Sheet:
    def __len__(self) -> int:
        """Number of rows in used range."""
        return self.workbook.engine.get_sheet_dimensions(self.name).rows

    def __iter__(self) -> Iterator[CellValue]:
        """Iterate over rows."""
        for row in self.values:
            yield row

    def __contains__(self, item: str) -> bool:
        """Check if cell exists."""
        return self.workbook.engine.cell_exists(item)
```

**Benefits:**
- `str(obj)` → `__str__`
- `repr(obj)` → `__repr__`
- `obj == other` → `__eq__`
- `hash(obj)` → `__hash__` (for frozen objects in sets/dicts)
- `len(obj)` → `__len__`
- `for x in obj` → `__iter__`
- `with obj:` → `__enter__`/`__exit__`
- `x in obj` → `__contains__`
- `obj[key]` → `__getitem__`
- `obj[key] = value` → `__setitem__`
