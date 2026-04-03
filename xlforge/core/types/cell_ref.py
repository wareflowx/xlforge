"""Cell reference value object."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    pass


# Regex for cell reference like "A1", "B2:C10"
CELL_REF_PATTERN = re.compile(r"^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$")


@dataclass(frozen=True, slots=True)
class CellRef:
    """Immutable cell reference.

    Represents a cell reference like 'A1' or a range like 'A1:C3'.
    May include sheet prefix like 'Data!A1'.
    """

    sheet: str
    coord: str

    @property
    def row(self) -> int:
        """Zero-based row index."""
        match = CELL_REF_PATTERN.match(self.coord)
        if not match:
            raise ValueError(f"Invalid cell reference: {self.coord}")
        return int(match.group(2)) - 1

    @property
    def col(self) -> int:
        """Zero-based column index (A=0)."""
        match = CELL_REF_PATTERN.match(self.coord)
        if not match:
            raise ValueError(f"Invalid cell reference: {self.coord}")
        col_letters = match.group(1).upper()
        return col_to_index(col_letters)

    @property
    def is_range(self) -> bool:
        """Check if this is a range reference."""
        return ":" in self.coord

    @property
    def end_row(self) -> int | None:
        """Zero-based end row for ranges, None for single cells."""
        if not self.is_range:
            return None
        match = CELL_REF_PATTERN.match(self.coord)
        if match and match.group(4):
            return int(match.group(4)) - 1
        return None

    @property
    def end_col(self) -> int | None:
        """Zero-based end column for ranges, None for single cells."""
        if not self.is_range:
            return None
        match = CELL_REF_PATTERN.match(self.coord)
        if match and match.group(3):
            return col_to_index(match.group(3).upper())
        return None

    def to_a1_notation(self) -> str:
        """Convert to 'A1' notation without sheet."""
        return self.coord

    def __str__(self) -> str:
        if self.sheet:
            return f"{self.sheet}!{self.coord}"
        return self.coord

    def __post_init__(self) -> None:
        if not self.coord:
            raise ValueError("Cell reference cannot be empty")


def col_to_index(col_letters: str) -> int:
    """Convert column letters (e.g., 'A', 'B', 'AA') to zero-based index."""
    result = 0
    for char in col_letters.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result - 1


def index_to_col(index: int) -> str:
    """Convert zero-based column index to letters."""
    result = ""
    index += 1
    while index > 0:
        index -= 1
        result = chr(ord("A") + index % 26) + result
        index //= 26
    return result


def cell_ref_to_row_col(coord: str) -> tuple[int, int]:
    """Convert cell reference to (row, col) zero-based tuple."""
    match = CELL_REF_PATTERN.match(coord)
    if not match:
        raise ValueError(f"Invalid cell reference: {coord}")
    row = int(match.group(2)) - 1
    col = col_to_index(match.group(1))
    return row, col


def row_col_to_cell_ref(row: int, col: int) -> str:
    """Convert (row, col) zero-based tuple to cell reference."""
    return f"{index_to_col(col)}{row + 1}"
