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

            typer.echo(f"Written: {cell_value.raw} ({cell_value.type.value}) to {coord}")

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
