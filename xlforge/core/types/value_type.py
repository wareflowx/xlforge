"""Value type enumeration for cell values."""

from __future__ import annotations

from enum import Enum


class ValueType(Enum):
    """Excel value types."""

    STRING = "string"
    NUMBER = "number"
    BOOL = "bool"
    DATE = "date"
    FORMULA = "formula"
    EMPTY = "empty"
    ERROR = "error"

    def __repr__(self) -> str:
        return f"ValueType.{self.name}"
