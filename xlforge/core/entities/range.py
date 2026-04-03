"""Range entity."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from xlforge.core.types.cell_ref import CellRef
from xlforge.core.types.cell_value import CellValue

if TYPE_CHECKING:
    from xlforge.core.entities.sheet import Sheet


class Range:
    """Represents a cell range with bulk operations."""

    def __init__(self, sheet: Sheet, coord: str) -> None:
        self._sheet = sheet
        self._coord = coord

    @property
    def sheet(self) -> Sheet:
        """Parent sheet."""
        return self._sheet

    @property
    def coord(self) -> str:
        """Range coordinate."""
        return self._coord

    @property
    def cell_ref(self) -> CellRef:
        """CellRef for this range's top-left cell."""
        return CellRef(sheet=self._sheet.name, coord=self._coord.split(":")[0])

    @property
    def values(self) -> list[list[CellValue]]:
        """Get all values in range as 2D array."""
        return self._sheet.workbook.engine.get_range(
            self._sheet.workbook.path, self._sheet.name, self._coord
        )

    def set_values(self, values: list[list[Any]]) -> None:
        """Set all values in range from 2D array.

        Args:
            values: 2D list of values to set.
        """
        cell_values = [[CellValue.from_python(v) for v in row] for row in values]
        self._sheet.workbook.engine.set_range(
            self._sheet.workbook.path, self._sheet.name, self._coord, cell_values
        )

    def clear(self) -> None:
        """Clear all cells in range."""
        from xlforge.core.types.cell_value import CellValue
        from xlforge.core.types.value_type import ValueType

        empty = CellValue(raw=None, type=ValueType.EMPTY)
        self._sheet.workbook.engine.set_range(
            self._sheet.workbook.path, self._sheet.name, self._coord, [[empty]]
        )

    def copy_to(self, dest_coord: str) -> None:
        """Copy range to destination.

        Args:
            dest_coord: Destination coordinate (top-left cell).
        """
        values = self.values
        dest_range = self._sheet.range(dest_coord)
        dest_range.set_values([[v.raw for v in row] for row in values])

    def __str__(self) -> str:
        return self._coord

    def __repr__(self) -> str:
        return f"Range({self._sheet.name!r}, {self._coord!r})"
