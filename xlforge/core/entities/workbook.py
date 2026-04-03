"""Workbook entity."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Self

from xlforge.core.entities.sheet import Sheet

if TYPE_CHECKING:
    from xlforge.core.engines.base import Engine


class Workbook:
    """Excel workbook entity.

    Manages workbook lifecycle with context manager pattern.
    """

    def __init__(
        self,
        path: Path,
        engine: Engine,
        *,
        read_only: bool = False,
        data_only: bool = True,
    ) -> None:
        self._path = path
        self._engine = engine
        self._read_only = read_only
        self._data_only = data_only
        self._is_open = False
        self._sheet_cache: dict[str, Sheet] = {}

    @property
    def path(self) -> Path:
        """Path to the workbook file."""
        return self._path

    @property
    def engine(self) -> Engine:
        """Engine used by this workbook."""
        return self._engine

    @property
    def is_open(self) -> bool:
        """Whether the workbook is currently open."""
        return self._is_open

    def open(self) -> Self:
        """Open the workbook."""
        if self._is_open:
            return self
        self._engine.open(
            self._path, read_only=self._read_only, data_only=self._data_only
        )
        self._is_open = True
        return self

    def close(self) -> None:
        """Close the workbook."""
        if not self._is_open:
            return
        self._engine.close(self._path)
        self._is_open = False
        self._sheet_cache.clear()

    def __enter__(self) -> Self:
        return self.open()

    def __str__(self) -> str:
        return str(self._path)

    def __repr__(self) -> str:
        return f"Workbook(path={self._path!r})"

    def __bool__(self) -> bool:
        return self._is_open

    def __exit__(self, *args: object) -> None:
        self.close()

    def sheet(self, name: str) -> Sheet:
        """Get a sheet by name.

        Args:
            name: Sheet name.

        Returns:
            Sheet entity.

        Raises:
            RuntimeError: If workbook is not open.
        """
        if not self._is_open:
            raise RuntimeError("Workbook not open. Call open() first.")
        if name not in self._sheet_cache:
            self._sheet_cache[name] = Sheet(name=name, workbook=self)
        return self._sheet_cache[name]

    def sheets(self) -> list[Sheet]:
        """Get all sheets.

        Returns:
            List of Sheet entities.
        """
        if not self._is_open:
            raise RuntimeError("Workbook not open. Call open() first.")
        sheet_names = self._engine.list_sheets(self._path)
        return [self.sheet(name) for name in sheet_names]

    def create_sheet(self, name: str) -> Sheet:
        """Create a new sheet.

        Args:
            name: Name for the new sheet.

        Returns:
            New Sheet entity.
        """
        if not self._is_open:
            raise RuntimeError("Workbook not open. Call open() first.")
        self._engine.create_sheet(self._path, name)
        return self.sheet(name)

    def delete_sheet(self, name: str) -> None:
        """Delete a sheet.

        Args:
            name: Name of sheet to delete.
        """
        if not self._is_open:
            raise RuntimeError("Workbook not open. Call open() first.")
        self._engine.delete_sheet(self._path, name)
        if name in self._sheet_cache:
            del self._sheet_cache[name]

    def rename_sheet(self, old_name: str, new_name: str) -> None:
        """Rename a sheet.

        Args:
            old_name: Current sheet name.
            new_name: New sheet name.
        """
        if not self._is_open:
            raise RuntimeError("Workbook not open. Call open() first.")
        self._engine.rename_sheet(self._path, old_name, new_name)
        if old_name in self._sheet_cache:
            sheet = self._sheet_cache.pop(old_name)
            self._sheet_cache[new_name] = sheet

    def save(self) -> None:
        """Save the workbook."""
        if not self._is_open:
            raise RuntimeError("Workbook not open. Call open() first.")
        self._engine.save(self._path)

    def copy_sheet(self, source_sheet: str, new_sheet: str) -> Sheet:
        """Copy a sheet to a new sheet.

        Args:
            source_sheet: Name of the sheet to copy.
            new_sheet: Name for the new sheet.

        Returns:
            New Sheet entity.
        """
        if not self._is_open:
            raise RuntimeError("Workbook not open. Call open() first.")
        self._engine.copy_sheet(self._path, source_sheet, new_sheet)
        return self.sheet(new_sheet)

    def set_active_sheet(self, sheet: str) -> None:
        """Set the active/selected sheet.

        Args:
            sheet: Sheet name to make active.
        """
        if not self._is_open:
            raise RuntimeError("Workbook not open. Call open() first.")
        self._engine.set_active_sheet(self._path, sheet)
