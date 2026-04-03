# Code Quality Patterns

This directory stores the project's code quality patterns and guidelines.

Patterns defined here should be followed consistently across the codebase to maintain high code quality, readability, and maintainability.

---

## Pattern: Entity-Oriented Design

**Rule:** Think in entities, not services. Business logic lives on objects.

**Rationale:**
- Entities represent domain objects with identity and state
- Business logic should live directly on entities (Fat Models pattern)
- Services introduce unnecessary indirection and hide state
- Entity methods have access to everything they need via composition
- Easier to reason about: "What can this object do?" vs "What can this service do?"

**Entity-oriented reasoning:**

| Instead of asking... | Ask... |
|-----------------------|--------|
| "What does CellService do?" | "What does Sheet do?" |
| "Where does this logic live?" | "Which entity owns this?" |
| "How do I get data from service?" | "How do I ask this object?" |

**Anti-pattern (do not use):**

```python
# DO NOT USE - Service thinking
class CellService:
    def __init__(self, engine: Engine):
        self._engine = engine

    def get_cell(self, path, sheet, coord):
        return self._engine.get_cell(path, sheet, coord)

    def set_cell(self, path, sheet, coord, value):
        return self._engine.set_cell(path, sheet, coord, value)

# Usage: cell_service.get_cell(path, sheet, coord)
```

**Preferred pattern:**

```python
# USE - Entity thinking
class Sheet:
    def __init__(self, name: str, workbook: Workbook):
        self._name = name
        self._workbook = workbook

    @property
    def engine(self) -> Engine:
        return self._workbook.engine

    def cell(self, coord: str) -> CellValue:
        """Get cell value."""
        return self.engine.get_cell(self._name, coord)

    def set_cell(self, coord: str, value: CellValue) -> None:
        """Set cell value."""
        self.engine.set_cell(self._name, coord, value)

# Usage: sheet.cell("A1")
```

**Key principles:**
1. **Entities own their state** - Sheet knows its name, Workbook knows its path
2. **Entities delegate to engines** - `self.engine.get_cell()` not `self._service.get()`
3. **Methods return domain types** - `CellValue` not raw `Any`
4. **Fluent interfaces when appropriate** - `sheet.cell("A1").as_string()`
