"""Cell service for xlforge."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Any

from xlforge.core.types.cell_ref import CellRef
from xlforge.core.types.cell_value import CellValue
from xlforge.core.types.value_type import ValueType
from xlforge.core.types.result import Result, Ok, Err
from xlforge.core.errors import ErrorCode

if TYPE_CHECKING:
    from xlforge.core.engines.base import Engine


class CellService:
    """Service for cell operations with Result-based error handling."""

    def __init__(self, engine: Engine) -> None:
        self._engine = engine

    def get(self, path: Path, cell_ref: str | CellRef) -> Result[CellValue, ErrorCode]:
        """Get cell value.

        Args:
            path: Path to the workbook.
            cell_ref: Cell reference (e.g., 'Data!A1' or 'A1').

        Returns:
            Result containing CellValue on success, ErrorCode on failure.
        """
        from xlforge.core.errors import ErrorCode, XlforgeError

        try:
            path = Path(path)
            if isinstance(cell_ref, str):
                cell_ref = self._parse_cell_ref(cell_ref)
            else:
                cell_ref = cell_ref

            if not path.exists():
                return Err(ErrorCode.FILE_NOT_FOUND)

            if not self._engine.sheet_exists(path, cell_ref.sheet):
                return Err(ErrorCode.SHEET_NOT_FOUND)

            value = self._engine.get_cell(path, cell_ref.sheet, cell_ref.coord)
            return Ok(value)

        except XlforgeError as e:
            return Err(e.code)
        except Exception:
            return Err(ErrorCode.GENERAL_ERROR)

    def set(
        self,
        path: Path,
        cell_ref: str | CellRef,
        value: Any,
        type_hint: ValueType | None = None,
    ) -> Result[None, ErrorCode]:
        """Set cell value.

        Args:
            path: Path to the workbook.
            cell_ref: Cell reference (e.g., 'Data!A1' or 'A1').
            value: Value to set.
            type_hint: Optional type hint for string coercion.

        Returns:
            Result containing None on success, ErrorCode on failure.
        """
        from xlforge.core.errors import ErrorCode, XlforgeError

        try:
            path = Path(path)
            if isinstance(cell_ref, str):
                cell_ref = self._parse_cell_ref(cell_ref)
            if not path.exists():
                return Err(ErrorCode.FILE_NOT_FOUND)
            if not self._engine.sheet_exists(path, cell_ref.sheet):
                return Err(ErrorCode.SHEET_NOT_FOUND)

            if not isinstance(value, CellValue):
                if isinstance(value, str):
                    value = CellValue.from_string(value, type_hint)
                else:
                    value = CellValue.from_python(value)

            self._engine.set_cell(path, cell_ref.sheet, cell_ref.coord, value)
            return Ok(None)

        except XlforgeError as e:
            return Err(e.code)
        except Exception:
            return Err(ErrorCode.GENERAL_ERROR)

    def _parse_cell_ref(self, ref: str) -> CellRef:
        """Parse cell reference string.

        Args:
            ref: Cell reference like 'A1' or 'Data!A1'.

        Returns:
            CellRef object.
        """
        if "!" in ref:
            sheet, coord = ref.split("!", 1)
            return CellRef(sheet=sheet, coord=coord)
        return CellRef(sheet="", coord=ref)
