"""Engine interface for xlforge."""

from __future__ import annotations

from abc import ABC
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlforge.core.types.cell_value import CellValue


class Engine(ABC):
    """Abstract base class for workbook engines.

    Engines provide the low-level Excel interaction for reading and writing cells.
    Implementations include OpenpyxlEngine (headless) and XlwingsEngine (full Excel).
    """

    def open(
        self, path: Path, *, read_only: bool = False, data_only: bool = True
    ) -> None:
        """Open a workbook.

        Args:
            path: Path to the workbook file.
            read_only: Open in read-only mode.
            data_only: Read cached values instead of formulas.
        """
        raise NotImplementedError

    def close(self, path: Path) -> None:
        """Close a workbook."""
        raise NotImplementedError

    def list_sheets(self, path: Path) -> list[str]:
        """List all sheet names in a workbook.

        Args:
            path: Path to the workbook file.

        Returns:
            List of sheet names.
        """
        raise NotImplementedError

    def get_cell(self, path: Path, sheet: str, coord: str) -> CellValue:
        """Get cell value.

        Args:
            path: Path to the workbook file.
            sheet: Sheet name.
            coord: Cell coordinate (e.g., 'A1').

        Returns:
            CellValue with the cell's value and type.
        """
        raise NotImplementedError

    def set_cell(self, path: Path, sheet: str, coord: str, value: CellValue) -> None:
        """Set cell value.

        Args:
            path: Path to the workbook file.
            sheet: Sheet name.
            coord: Cell coordinate (e.g., 'A1').
            value: CellValue to set.
        """
        raise NotImplementedError

    def get_range(self, path: Path, sheet: str, coord: str) -> list[list[CellValue]]:
        """Get range values as 2D array.

        Args:
            path: Path to the workbook file.
            sheet: Sheet name.
            coord: Range coordinate (e.g., 'A1:C3').

        Returns:
            2D list of CellValues.
        """
        raise NotImplementedError

    def set_range(
        self, path: Path, sheet: str, coord: str, values: list[list[CellValue]]
    ) -> None:
        """Set range values from 2D array.

        Args:
            path: Path to the workbook file.
            sheet: Sheet name.
            coord: Range coordinate (e.g., 'A1:C3').
            values: 2D list of CellValues.
        """
        raise NotImplementedError

    def create_sheet(self, path: Path, sheet: str) -> None:
        """Create a new sheet.

        Args:
            path: Path to the workbook file.
            sheet: Name for the new sheet.
        """
        raise NotImplementedError

    def delete_sheet(self, path: Path, sheet: str) -> None:
        """Delete a sheet.

        Args:
            path: Path to the workbook file.
            sheet: Sheet name to delete.
        """
        raise NotImplementedError

    def rename_sheet(self, path: Path, old_name: str, new_name: str) -> None:
        """Rename a sheet.

        Args:
            path: Path to the workbook file.
            old_name: Current sheet name.
            new_name: New sheet name.
        """
        raise NotImplementedError

    def save(self, path: Path) -> None:
        """Save a workbook.

        Args:
            path: Path to the workbook file.
        """
        raise NotImplementedError

    def sheet_exists(self, path: Path, sheet: str) -> bool:
        """Check if a sheet exists.

        Args:
            path: Path to the workbook file.
            sheet: Sheet name.

        Returns:
            True if sheet exists.
        """
        raise NotImplementedError

    def get_sheet_dimensions(self, path: Path, sheet: str) -> str:
        """Get the used range dimensions for a sheet.

        Args:
            path: Path to the workbook file.
            sheet: Sheet name.

        Returns:
            Range string (e.g., 'A1:C10') representing the used range.
        """
        raise NotImplementedError

    def cell_exists(self, path: Path, sheet: str, coord: str) -> bool:
        """Check if a cell exists within the used range.

        Args:
            path: Path to the workbook file.
            sheet: Sheet name.
            coord: Cell coordinate (e.g., 'A1').

        Returns:
            True if the cell exists in the used range.
        """
        raise NotImplementedError

    def copy_sheet(self, path: Path, source_sheet: str, new_sheet: str) -> None:
        """Copy a sheet to a new sheet.

        Args:
            path: Path to the workbook file.
            source_sheet: Name of the sheet to copy.
            new_sheet: Name for the new sheet.
        """
        raise NotImplementedError

    def set_active_sheet(self, path: Path, sheet: str) -> None:
        """Set the active/selected sheet.

        Args:
            path: Path to the workbook file.
            sheet: Sheet name to make active.
        """
        raise NotImplementedError
