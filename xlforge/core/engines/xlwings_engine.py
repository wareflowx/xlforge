"""Xlwings-based engine for xlforge using full Excel application."""

from __future__ import annotations

from datetime import datetime
from importlib.util import find_spec
from pathlib import Path
from typing import Any

from xlforge.core.engines.base import Engine
from xlforge.core.types.cell_value import CellValue
from xlforge.core.types.value_type import ValueType
from xlforge.core.types.cell_ref import cell_ref_to_row_col


class XlwingsEngine(Engine):
    """Xlwings-based engine using the full Excel application.

    This engine provides full Excel compatibility including:
    - Working formulas
    - Excel-specific features
    - Live Excel connection

    Requires Excel to be installed on Windows or macOS.
    """

    def __init__(self) -> None:
        self._workbooks: dict[Path, Any] = {}  # xlwings.Book objects
        self._xlwings: Any = None  # Lazily loaded

    @property
    def _xw(self) -> Any:
        """Lazy load xlwings module."""
        if self._xlwings is None:
            if not self._is_xlwings_available():
                raise ImportError(
                    "xlwings is not available. Install xlwings: pip install xlwings"
                )
            import xlwings
            self._xlwings = xlwings
        return self._xlwings

    @staticmethod
    def _is_xlwings_available() -> bool:
        """Check if xlwings module is available."""
        return find_spec("xlwings") is not None

    def _get_workbook(self, path: Path) -> Any:
        """Get an open workbook or raise FileNotFoundError."""
        wb = self._workbooks.get(path)
        if wb is None:
            raise FileNotFoundError(f"Workbook not open: {path}")
        return wb

    def _cell_to_value(self, cell: Any) -> CellValue:
        """Convert xlwings cell to CellValue."""
        value = cell.value

        if value is None:
            return CellValue(raw=None, type=ValueType.EMPTY)

        if isinstance(value, bool):
            return CellValue(raw=value, type=ValueType.BOOL)

        if isinstance(value, (int, float)):
            return CellValue(raw=value, type=ValueType.NUMBER)

        if isinstance(value, str):
            if value.startswith("="):
                return CellValue(raw=value, type=ValueType.FORMULA)
            return CellValue(raw=value, type=ValueType.STRING)

        # xlwings returns datetime objects for date cells
        if isinstance(value, datetime):
            return CellValue(raw=value, type=ValueType.DATE)

        # Handle pandas.Timestamp and other datetime-like objects
        if hasattr(value, "year") and hasattr(value, "month"):
            return CellValue(raw=value, type=ValueType.DATE)

        return CellValue(raw=str(value), type=ValueType.STRING)

    def _value_to_cell(self, value: CellValue) -> Any:
        """Convert CellValue to xlwings-compatible value."""
        if value.type == ValueType.EMPTY:
            return None
        return value.raw

    def open(
        self, path: Path, *, read_only: bool = False, data_only: bool = True
    ) -> None:
        """Open a workbook with xlwings."""
        if path in self._workbooks:
            return

        xw = self._xw
        wb = xw.Book(path)

        if read_only:
            # xlwings doesn't have direct read_only mode like openpyxl
            # but we can open it normally and let Excel handle it
            pass

        if data_only:
            # xlwings automatically gets cached values by default
            # This is the default behavior
            pass

        self._workbooks[path] = wb

    def close(self, path: Path) -> None:
        """Close a workbook."""
        if path in self._workbooks:
            self._workbooks[path].close()
            del self._workbooks[path]

    def list_sheets(self, path: Path) -> list[str]:
        """List all sheet names."""
        wb = self._get_workbook(path)
        return [sheet.name for sheet in wb.sheets]

    def get_cell(self, path: Path, sheet: str, coord: str) -> CellValue:
        """Get cell value using xlwings."""
        wb = self._get_workbook(path)
        ws = wb.sheets[sheet]
        cell = ws.range(coord)
        return self._cell_to_value(cell)

    def set_cell(self, path: Path, sheet: str, coord: str, value: CellValue) -> None:
        """Set cell value using xlwings."""
        wb = self._get_workbook(path)
        ws = wb.sheets[sheet]
        cell = ws.range(coord)
        cell.value = self._value_to_cell(value)

    def get_range(self, path: Path, sheet: str, coord: str) -> list[list[CellValue]]:
        """Get range values as 2D array."""
        wb = self._get_workbook(path)
        ws = wb.sheets[sheet]
        rng = ws.range(coord)

        # Get the values as a 2D list
        values = rng.value

        # Handle single cell case (returns scalar, not 2D list)
        if not isinstance(values, list):
            return [[self._cell_to_value(rng)]]

        # Ensure we have a 2D list (xlwings might return 1D for single row)
        if values and not isinstance(values[0], list):
            values = [values]

        result: list[list[CellValue]] = []
        for row in values:
            row_values: list[CellValue] = []
            if isinstance(row, list):
                for cell_val in row:
                    # Create a mock cell-like object for _cell_to_value
                    row_values.append(CellValue.from_python(cell_val))
            else:
                row_values.append(CellValue.from_python(row))
            result.append(row_values)

        return result

    def set_range(
        self, path: Path, sheet: str, coord: str, values: list[list[CellValue]]
    ) -> None:
        """Set range values from 2D array."""
        wb = self._get_workbook(path)
        ws = wb.sheets[sheet]
        rng = ws.range(coord)

        # Convert CellValue 2D array to plain Python values
        plain_values: list[list[Any]] = []
        for row in values:
            plain_row: list[Any] = []
            for value in row:
                plain_row.append(self._value_to_cell(value))
            plain_values.append(plain_row)

        rng.value = plain_values

    def create_sheet(self, path: Path, sheet: str) -> None:
        """Create a new sheet."""
        wb = self._get_workbook(path)
        wb.sheets.add(name=sheet)

    def delete_sheet(self, path: Path, sheet: str) -> None:
        """Delete a sheet."""
        wb = self._get_workbook(path)
        if sheet in [s.name for s in wb.sheets]:
            wb.sheets[sheet].delete()

    def rename_sheet(self, path: Path, old_name: str, new_name: str) -> None:
        """Rename a sheet."""
        wb = self._get_workbook(path)
        wb.sheets[old_name].name = new_name

    def save(self, path: Path) -> None:
        """Save a workbook."""
        wb = self._get_workbook(path)
        wb.save(path)

    def sheet_exists(self, path: Path, sheet: str) -> bool:
        """Check if a sheet exists."""
        wb = self._workbooks.get(path)
        if wb is None:
            return False
        return sheet in [s.name for s in wb.sheets]

    def get_sheet_dimensions(self, path: Path, sheet: str) -> str:
        """Get the used range dimensions for a sheet."""
        wb = self._get_workbook(path)
        ws = wb.sheets[sheet]
        # xlwings uses 'used_range' property
        try:
            dimensions = ws.used_range.address
            return dimensions
        except Exception:
            # If we can't get dimensions, return a sensible default
            return "A1"

    def cell_exists(self, path: Path, sheet: str, coord: str) -> bool:
        """Check if a cell exists within the used range."""
        wb = self._get_workbook(path)
        ws = wb.sheets[sheet]

        try:
            # Get the cell's used range
            used_range = ws.used_range
            if used_range is None:
                return False

            # Get the top-left and bottom-right of used range
            start_cell = used_range(1, 1)
            end_row = used_range.rows.count
            end_col = used_range.columns.count

            # Parse the target cell coordinates using cell_ref_to_row_col
            cell_row, cell_col = cell_ref_to_row_col(coord)

            # Check if coord is within used range (1-based for xlwings)
            start_row = start_cell.row
            start_col = start_cell.column

            return (start_row <= cell_row + 1 <= start_row + end_row - 1 and
                    start_col <= cell_col + 1 <= start_col + end_col - 1)
        except Exception:
            return False

    def copy_sheet(self, path: Path, source_sheet: str, new_sheet: str) -> None:
        """Copy a sheet to a new sheet."""
        wb = self._get_workbook(path)
        source_ws = wb.sheets[source_sheet]
        new_ws = source_ws.copy(after=source_ws)
        new_ws.name = new_sheet

    def set_active_sheet(self, path: Path, sheet: str) -> None:
        """Set the active/selected sheet."""
        wb = self._get_workbook(path)
        wb.sheets[sheet].activate()
