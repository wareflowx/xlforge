"""Sheet entity."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from xlforge.core.types.cell_ref import CellRef
from xlforge.core.types.cell_value import CellValue

if TYPE_CHECKING:
    from xlforge.core.entities.workbook import Workbook
    from xlforge.core.entities.range import Range


class Sheet:
    """Sheet entity with cell access."""

    def __init__(self, name: str, workbook: Workbook) -> None:
        self._name = name
        self._workbook = workbook

    @property
    def name(self) -> str:
        """Sheet name."""
        return self._name

    @property
    def workbook(self) -> Workbook:
        """Parent workbook."""
        return self._workbook

    def cell(self, coord: str | CellRef) -> CellValue:
        """Get cell value.

        Args:
            coord: Cell coordinate (e.g., 'A1') or CellRef.

        Returns:
            CellValue with the cell's value and type.
        """
        if isinstance(coord, CellRef):
            coord = coord.coord
        return self._workbook.engine.get_cell(self._workbook.path, self._name, coord)

    def set_cell(self, coord: str | CellRef, value: Any) -> None:
        """Set cell value.

        Args:
            coord: Cell coordinate (e.g., 'A1') or CellRef.
            value: Value to set (will be converted to CellValue).
        """
        if isinstance(coord, CellRef):
            coord = coord.coord
        if not isinstance(value, CellValue):
            value = CellValue.from_python(value)
        self._workbook.engine.set_cell(self._workbook.path, self._name, coord, value)

    def range(self, coord: str) -> Range:
        """Get a range for bulk operations.

        Args:
            coord: Range coordinate (e.g., 'A1:C3').

        Returns:
            Range entity.
        """
        return Range(sheet=self, coord=coord)

    def clear(self, coord: str | None = None) -> None:
        """Clear a cell or range.

        Args:
            coord: Cell or range to clear. If None, clears entire sheet content.
        """
        from xlforge.core.types.cell_value import CellValue
        from xlforge.core.types.value_type import ValueType

        if coord is None:
            coord = "A1:ZZZ1048576"
        if isinstance(coord, CellRef):
            coord = coord.coord
        empty = CellValue(raw=None, type=ValueType.EMPTY)
        self._workbook.engine.set_range(
            self._workbook.path, self._name, coord, [[empty]]
        )

    @property
    def is_protected(self) -> bool:
        """Check if sheet is protected."""
        # TODO: Implement with openpyxl/xlwings
        return False

    def __repr__(self) -> str:
        return f"Sheet({self._name!r})"
