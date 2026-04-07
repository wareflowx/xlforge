"""Cell operations CLI commands."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Annotated, Optional

import typer

from xlforge.core.entities.workbook import Workbook
from xlforge.core.engines.selector import EngineSelector
from xlforge.core.errors import ErrorCode, XlforgeError
from xlforge.core.types.cell_value import CellValue
from xlforge.core.types.value_type import ValueType

cell_app = typer.Typer(help="Cell operations for Excel workbooks.")


@cell_app.command()
def read(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[str, typer.Argument(help="Sheet name.")],
    coord: Annotated[str, typer.Argument(help="Cell coordinate (e.g., A1).")],
    json_output: Annotated[
        bool, typer.Option("--json", "-j", help="Output as JSON.")
    ] = False,
) -> None:
    """Read a cell value from a sheet."""
    # Check if file exists
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        engine = EngineSelector.for_path(path)
        workbook = Workbook(path=path, engine=engine, read_only=True)

        with workbook:
            # Check if sheet exists
            if not engine.sheet_exists(path, sheet):
                typer.secho(
                    f"Error: Sheet not found: {sheet}",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            sheet_obj = workbook.sheet(name=sheet)
            cell_value = sheet_obj.cell(coord)

            if json_output:
                data = {
                    "value": cell_value.raw,
                    "type": cell_value.type.value,
                    "coord": coord,
                    "sheet": sheet,
                }
                typer.echo(json.dumps(data, indent=2))
            else:
                typer.echo(f"Value: {cell_value.raw}")
                typer.echo(f"Type: {cell_value.type.value}")

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@cell_app.command()
def write(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[str, typer.Argument(help="Sheet name.")],
    coord: Annotated[str, typer.Argument(help="Cell coordinate (e.g., A1).")],
    value: Annotated[str, typer.Argument(help="Value to write.")],
    value_type: Annotated[
        Optional[str],
        typer.Option("--type", "-t", help="Value type: string, number, bool, date."),
    ] = None,
) -> None:
    """Write a value to a cell."""
    # Check if file exists
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        engine = EngineSelector.for_path(path)
        workbook = Workbook(path=path, engine=engine, read_only=False)

        with workbook:
            # Check if sheet exists
            if not engine.sheet_exists(path, sheet):
                typer.secho(
                    f"Error: Sheet not found: {sheet}",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            # Parse value type
            type_hint = None
            if value_type is not None:
                type_str = value_type.lower()
                if type_str == "string":
                    type_hint = ValueType.STRING
                elif type_str == "number":
                    type_hint = ValueType.NUMBER
                elif type_str == "bool":
                    type_hint = ValueType.BOOL
                elif type_str == "date":
                    type_hint = ValueType.DATE
                else:
                    typer.secho(
                        f"Error: Invalid type: {value_type}. Use string, number, bool, or date.",
                        fg=typer.colors.RED,
                        err=True,
                    )
                    raise typer.Exit(code=int(ErrorCode.TYPE_COERCION_FAILED))

            # Create CellValue with type coercion
            cell_value = CellValue.from_string(value, type_hint)

            # Write to cell
            sheet_obj = workbook.sheet(name=sheet)
            sheet_obj.set_cell(coord, cell_value)

            # Save the workbook
            workbook.save()

            typer.echo(
                f"Written: {cell_value.raw} ({cell_value.type.value}) to {coord}"
            )

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@cell_app.command()
def formula(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[str, typer.Argument(help="Sheet name.")],
    coord: Annotated[str, typer.Argument(help="Cell coordinate (e.g., A1).")],
    formula: Annotated[str, typer.Argument(help="Formula to set (e.g., =SUM(A:A)).")],
) -> None:
    """Set a formula in a cell."""
    # Check if file exists
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    # Ensure formula starts with =
    if not formula.startswith("="):
        formula = f"={formula}"

    try:
        engine = EngineSelector.for_path(path)
        workbook = Workbook(path=path, engine=engine, read_only=False)

        with workbook:
            # Check if sheet exists
            if not engine.sheet_exists(path, sheet):
                typer.secho(
                    f"Error: Sheet not found: {sheet}",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            # Create CellValue with FORMULA type
            cell_value = CellValue(raw=formula, type=ValueType.FORMULA)

            # Write to cell
            sheet_obj = workbook.sheet(name=sheet)
            sheet_obj.set_cell(coord, cell_value)

            # Save the workbook
            workbook.save()

            typer.echo(f"Formula set: {formula} at {coord}")

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@cell_app.command()
def clear(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[str, typer.Argument(help="Sheet name.")],
    coord: Annotated[str, typer.Argument(help="Cell coordinate (e.g., A1).")],
) -> None:
    """Clear a cell value."""
    # Check if file exists
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        engine = EngineSelector.for_path(path)
        workbook = Workbook(path=path, engine=engine, read_only=False)

        with workbook:
            # Check if sheet exists
            if not engine.sheet_exists(path, sheet):
                typer.secho(
                    f"Error: Sheet not found: {sheet}",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            # Create empty CellValue
            cell_value = CellValue(raw=None, type=ValueType.EMPTY)

            # Clear the cell
            sheet_obj = workbook.sheet(name=sheet)
            sheet_obj.set_cell(coord, cell_value)

            # Save the workbook
            workbook.save()

            typer.echo(f"Cleared cell {coord}")

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@cell_app.command()
def copy(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    src_sheet: Annotated[str, typer.Argument(help="Source sheet name.")],
    src_cell: Annotated[str, typer.Argument(help="Source cell coordinate (e.g., A1).")],
    dst_sheet: Annotated[str, typer.Argument(help="Destination sheet name.")],
    dst_cell: Annotated[
        str, typer.Argument(help="Destination cell coordinate (e.g., B1).")
    ],
) -> None:
    """Copy a cell value to another location."""
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        engine = EngineSelector.for_path(path)
        workbook = Workbook(path=path, engine=engine, read_only=False)

        with workbook:
            # Check if source sheet exists
            if not engine.sheet_exists(path, src_sheet):
                typer.secho(
                    f"Error: Source sheet not found: {src_sheet}",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            # Check if destination sheet exists
            if not engine.sheet_exists(path, dst_sheet):
                typer.secho(
                    f"Error: Destination sheet not found: {dst_sheet}",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            # Read source cell value
            src_sheet_obj = workbook.sheet(name=src_sheet)
            cell_value = src_sheet_obj.cell(src_cell)

            # Write to destination cell
            dst_sheet_obj = workbook.sheet(name=dst_sheet)
            dst_sheet_obj.set_cell(dst_cell, cell_value)

            # Save the workbook
            workbook.save()

            typer.echo(
                f"Copied {src_sheet}!{src_cell} ({cell_value.raw}) to {dst_sheet}!{dst_cell}"
            )

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@cell_app.command()
def search(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    query: Annotated[str, typer.Argument(help="Search query string.")],
    sheet_name: Annotated[
        Optional[str], typer.Option("--sheet", "-s", help="Sheet name to search in.")
    ] = None,
    json_output: Annotated[
        bool, typer.Option("--json", "-j", help="Output as JSON.")
    ] = False,
) -> None:
    """Search for a cell containing the query string."""
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        engine = EngineSelector.for_path(path)
        workbook = Workbook(path=path, engine=engine, read_only=False)

        with workbook:
            # Get sheets to search
            if sheet_name is not None:
                if not engine.sheet_exists(path, sheet_name):
                    typer.secho(
                        f"Error: Sheet not found: {sheet_name}",
                        fg=typer.colors.RED,
                        err=True,
                    )
                    raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))
                sheets_to_search = [sheet_name]
            else:
                sheets_to_search = engine.list_sheets(path)

            # Search each sheet
            for sname in sheets_to_search:
                dimensions = engine.get_sheet_dimensions(path, sname)

                if not dimensions:
                    continue

                # Parse dimensions to get range
                try:
                    from xlforge.core.types.cell_ref import (
                        cell_ref_to_row_col,
                        index_to_col,
                    )

                    start_coord = dimensions.split(":")[0]
                    start_row, start_col = cell_ref_to_row_col(start_coord)

                    range_values = engine.get_range(path, sname, dimensions)
                    for row_idx, row in enumerate(range_values):
                        for col_idx, cell_val in enumerate(row):
                            if (
                                cell_val.type == ValueType.STRING
                                and cell_val.raw is not None
                            ):
                                if query.lower() in str(cell_val.raw).lower():
                                    # Found a match - calculate cell coord (0-based to 1-based)
                                    found_col = index_to_col(start_col + col_idx)
                                    found_row = start_row + row_idx + 1
                                    found_coord = f"{found_col}{found_row}"

                                    if json_output:
                                        data = {
                                            "value": cell_val.raw,
                                            "type": cell_val.type.value,
                                            "coord": found_coord,
                                            "sheet": sname,
                                        }
                                        typer.echo(json.dumps(data, indent=2))
                                    else:
                                        typer.echo(
                                            f"Found in {sname}!{found_coord}: {cell_val.raw}"
                                        )
                                    return

                    # Also check if query matches in string conversion of other types
                    for row_idx, row in enumerate(range_values):
                        for col_idx, cell_val in enumerate(row):
                            cell_str = (
                                str(cell_val.raw) if cell_val.raw is not None else ""
                            )
                            if query.lower() in cell_str.lower():
                                found_col = index_to_col(start_col + col_idx)
                                found_row = start_row + row_idx + 1
                                found_coord = f"{found_col}{found_row}"

                                if json_output:
                                    data = {
                                        "value": cell_val.raw,
                                        "type": cell_val.type.value,
                                        "coord": found_coord,
                                        "sheet": sname,
                                    }
                                    typer.echo(json.dumps(data, indent=2))
                                else:
                                    typer.echo(
                                        f"Found in {sname}!{found_coord}: {cell_val.raw}"
                                    )
                                return
                except Exception:
                    continue

            typer.secho(
                f"Error: No cell found containing '{query}'",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.CELL_NOT_FOUND))

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@cell_app.command()
def bulk(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[str, typer.Argument(help="Sheet name.")],
    coord: Annotated[str, typer.Argument(help="Range coordinate (e.g., A1:C3).")],
    set_value: Annotated[
        Optional[str],
        typer.Option("--set", help="Value to set for all cells in range."),
    ] = None,
    clear_cells: Annotated[
        bool,
        typer.Option("--clear", help="Clear all cells in range."),
    ] = False,
) -> None:
    """Bulk operations on a range of cells."""
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    if set_value is None and not clear_cells:
        typer.secho(
            "Error: Must specify either --set <value> or --clear",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

    try:
        engine = EngineSelector.for_path(path)
        workbook = Workbook(path=path, engine=engine, read_only=False)

        with workbook:
            if not engine.sheet_exists(path, sheet):
                typer.secho(
                    f"Error: Sheet not found: {sheet}",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            from xlforge.core.types.value_type import ValueType
            from xlforge.core.types.cell_ref import cell_ref_to_row_col

            if clear_cells:
                empty = CellValue(raw=None, type=ValueType.EMPTY)
                # Use set_cell for single cells, set_range for ranges
                if ":" not in coord:
                    engine.set_cell(path, sheet, coord, empty)
                else:
                    engine.set_range(path, sheet, coord, [[empty]])
                workbook.save()
                typer.echo(f"Cleared range {coord}")
            else:
                # Parse the range to get dimensions
                if ":" not in coord:
                    # Single cell - use set_cell
                    cell_value = CellValue.from_string(set_value)  # type: ignore[arg-type]
                    engine.set_cell(path, sheet, coord, cell_value)
                    workbook.save()
                    typer.echo(f"Set {coord} = {set_value}")
                else:
                    start_cell, end_cell = coord.split(":")
                    start_row, start_col = cell_ref_to_row_col(start_cell)
                    end_row, end_col = cell_ref_to_row_col(end_cell)

                    rows = end_row - start_row + 1
                    cols = end_col - start_col + 1

                    # Create 2D array with the set_value
                    cell_value = CellValue.from_string(set_value)  # type: ignore[arg-type]
                    values = [[cell_value for _ in range(cols)] for _ in range(rows)]
                    engine.set_range(path, sheet, coord, values)
                    workbook.save()
                    typer.echo(f"Set {coord} = {set_value}")

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@cell_app.command()
def fill(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[str, typer.Argument(help="Sheet name.")],
    coord: Annotated[str, typer.Argument(help="Range coordinate (e.g., A1:C3).")],
) -> None:
    """Auto-fill a range by copying the first cell value."""
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        engine = EngineSelector.for_path(path)
        workbook = Workbook(path=path, engine=engine, read_only=False)

        with workbook:
            if not engine.sheet_exists(path, sheet):
                typer.secho(
                    f"Error: Sheet not found: {sheet}",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            # Get first cell value
            range_values = engine.get_range(path, sheet, coord)
            if not range_values or not range_values[0]:
                typer.secho(
                    f"Error: Range {coord} is empty",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=1)

            first_value = range_values[0][0]

            # Check if first cell is empty
            if first_value.type == ValueType.EMPTY:
                typer.secho(
                    f"Error: Range {coord} is empty",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=1)

            # Parse the range to get dimensions
            from xlforge.core.types.cell_ref import cell_ref_to_row_col

            start_cell, end_cell = coord.split(":")
            start_row, start_col = cell_ref_to_row_col(start_cell)
            end_row, end_col = cell_ref_to_row_col(end_cell)

            rows = end_row - start_row + 1
            cols = end_col - start_col + 1

            # Create 2D array with the first value
            fill_values = [[first_value for _ in range(cols)] for _ in range(rows)]
            engine.set_range(path, sheet, coord, fill_values)
            workbook.save()
            typer.echo(f"Filled range {coord} with value: {first_value.raw}")

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)
