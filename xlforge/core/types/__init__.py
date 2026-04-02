"""Type definitions for xlforge SDK."""

from xlforge.core.types.result import (
    Ok,
    Err,
    Result,
    Some,
    Nothing,
    Maybe,
    is_ok,
    is_err,
    is_some,
    is_nothing,
)
from xlforge.core.types.value_type import ValueType
from xlforge.core.types.cell_ref import (
    CellRef,
    col_to_index,
    index_to_col,
    cell_ref_to_row_col,
    row_col_to_cell_ref,
)
from xlforge.core.types.cell_value import CellValue

__all__ = [
    # Result types
    "Ok",
    "Err",
    "Result",
    "Some",
    "Nothing",
    "Maybe",
    "is_ok",
    "is_err",
    "is_some",
    "is_nothing",
    # Value types
    "ValueType",
    # Value objects
    "CellRef",
    "CellValue",
    # Utilities
    "col_to_index",
    "index_to_col",
    "cell_ref_to_row_col",
    "row_col_to_cell_ref",
]
