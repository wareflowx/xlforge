"""Cell value value object."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any

from xlforge.core.types.value_type import ValueType


@dataclass(frozen=True, slots=True)
class CellValue:
    """Immutable cell value with type information.

    Use factory methods to create instances:
    - CellValue.from_python(value) - infer type from Python value
    - CellValue.from_string(value, type_hint) - parse from string with optional type
    """

    raw: Any
    type: ValueType

    def as_string(self) -> str:
        """Convert to string."""
        if self.type == ValueType.EMPTY:
            return ""
        if isinstance(self.raw, (int, float)):
            return str(self.raw)
        return str(self.raw)

    def as_number(self) -> float:
        """Convert to number."""
        if self.type == ValueType.NUMBER:
            return float(self.raw)
        if self.type == ValueType.BOOL:
            return 1.0 if self.raw else 0.0
        raise TypeError(f"Cannot convert {self.type} to number")

    def as_bool(self) -> bool:
        """Convert to boolean."""
        if self.type == ValueType.BOOL:
            return bool(self.raw)
        raise TypeError(f"Cannot convert {self.type} to bool")

    def as_date(self) -> datetime:
        """Convert to datetime."""
        if self.type == ValueType.DATE:
            return self.raw
        raise TypeError(f"Cannot convert {self.type} to date")

    def is_empty(self) -> bool:
        """Check if cell is empty."""
        return self.type == ValueType.EMPTY

    def is_error(self) -> bool:
        """Check if cell contains an error."""
        return self.type == ValueType.ERROR

    @classmethod
    def from_python(cls, value: Any) -> CellValue:
        """Create CellValue from a Python value with inferred type."""
        if value is None:
            return cls(raw=None, type=ValueType.EMPTY)

        if isinstance(value, bool):
            return cls(raw=value, type=ValueType.BOOL)

        if isinstance(value, (int, float)):
            return cls(raw=value, type=ValueType.NUMBER)

        if isinstance(value, datetime):
            return cls(raw=value, type=ValueType.DATE)

        if isinstance(value, str):
            if value.startswith("="):
                return cls(raw=value, type=ValueType.FORMULA)
            return cls(raw=value, type=ValueType.STRING)

        # Fallback to string
        return cls(raw=str(value), type=ValueType.STRING)

    @classmethod
    def from_string(cls, value: str, type_hint: ValueType | None = None) -> CellValue:
        """Create CellValue from string with optional type hint.

        Args:
            value: String value to parse
            type_hint: Optional type to force (e.g., ValueType.STRING to preserve leading zeros)
        """
        if not value or value == "":
            return cls(raw=None, type=ValueType.EMPTY)

        if type_hint:
            coerced = _coerce_string_to_type(value, type_hint)
            return cls(raw=coerced, type=type_hint)

        # Infer type from string content
        inferred = _infer_type_from_string(value)
        coerced = _coerce_string_to_type(value, inferred)
        return cls(raw=coerced, type=inferred)

    def __repr__(self) -> str:
        return f"CellValue({self.raw!r}, {self.type!r})"


def _infer_type_from_string(value: str) -> ValueType:
    """Infer ValueType from string content."""
    if value.startswith("="):
        return ValueType.FORMULA

    # Try boolean
    if value.upper() in ("TRUE", "FALSE"):
        return ValueType.BOOL

    # Try number
    try:
        float(value)
        return ValueType.NUMBER
    except ValueError:
        pass

    # Try date (ISO format)
    try:
        datetime.fromisoformat(value.replace("Z", "+00:00"))
        return ValueType.DATE
    except ValueError:
        pass

    return ValueType.STRING


def _coerce_string_to_type(value: str, type_hint: ValueType) -> Any:
    """Coerce string value to the specified type."""
    if type_hint == ValueType.STRING:
        return value

    if type_hint == ValueType.NUMBER:
        # Handle comma decimal separator
        normalized = value.replace(",", ".")
        return float(normalized)

    if type_hint == ValueType.BOOL:
        return value.upper() in ("TRUE", "1", "YES")

    if type_hint == ValueType.DATE:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))

    if type_hint == ValueType.FORMULA:
        return value if value.startswith("=") else f"={value}"

    if type_hint == ValueType.EMPTY:
        return None

    return value
