"""Range operations CLI commands."""

from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Annotated, Optional

import typer

from xlforge.core.entities.workbook import Workbook
from xlforge.core.engines.selector import EngineSelector
from xlforge.core.errors import ErrorCode, XlforgeError
from xlforge.core.types.cell_value import CellValue

range_app = typer.Typer(help="Range operations for Excel workbooks.")


def _format_table(values: list[list]) -> str:
    """Format 2D array as a simple text table."""
    if not values:
        return ""

    # Calculate column widths
    col_widths = [max(len(str(row[i])) for row in values) for i in range(len(values[0]))]

    lines = []
    for row in values:
        formatted_cells = [str(cell).ljust(width) for cell, width in zip(row, col_widths)]
        lines.append(" | ".join(formatted_cells))

    return "\n".join(lines)


@range_app.command()
def read(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[str, typer.Argument(help="Sheet name.")],
    coord: Annotated[str, typer.Argument(help="Range coordinate (e.g., A1:C3).")],
    json_output: Annotated[
        bool, typer.Option("--json", "-j", help="Output as JSON.")
    ] = False,
) -> None:
    """Read a range of cells from a sheet."""
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

            # Get range values
            range_values = engine.get_range(path, sheet, coord)

            if json_output:
                # Output as JSON array of arrays
                data = [[cell.raw for cell in row] for row in range_values]
                typer.echo(json.dumps(data, indent=2))
            else:
                # Output as formatted table
                if not range_values:
                    typer.echo(f"Range {coord} is empty.")
                    return

                # Format as string values for table display
                table_values = [[str(cell.raw) for cell in row] for row in range_values]
                table = _format_table(table_values)
                typer.echo(f"Range: {coord}")
                typer.echo(f"Dimensions: {len(range_values)} rows x {len(range_values[0]) if range_values else 0} columns")
                typer.echo("")
                typer.echo(table)

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@range_app.command()
def write(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[str, typer.Argument(help="Sheet name.")],
    coord: Annotated[str, typer.Argument(help="Range coordinate (e.g., A1:C3).")],
    values_json: Annotated[
        Optional[str],
        typer.Argument(help="JSON array of values to write (e.g., '[[\"A\",1,true],[\"B\",2,false]]')."),
    ] = None,
    csv_file: Annotated[
        Optional[Path],
        typer.Option("--csv", "-c", help="Path to CSV file with values to write."),
    ] = None,
) -> None:
    """Write values to a range in a sheet."""
    # Check if file exists
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    # Must provide either values_json or csv_file
    if values_json is None and csv_file is None:
        typer.secho(
            "Error: Must provide either values as JSON or a CSV file with --csv.",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

    if values_json is not None and csv_file is not None:
        typer.secho(
            "Error: Cannot specify both JSON values and CSV file. Choose one.",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

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

            # Parse values based on input type
            if csv_file is not None:
                # Read from CSV file
                if not csv_file.exists():
                    typer.secho(
                        f"Error: CSV file does not exist: {csv_file}",
                        fg=typer.colors.RED,
                        err=True,
                    )
                    raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

                with open(csv_file, "r", newline="", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    values = [[cell for cell in row] for row in reader]
            else:
                # Parse JSON
                try:
                    values = json.loads(values_json)
                except json.JSONDecodeError as e:
                    typer.secho(
                        f"Error: Invalid JSON format: {e}",
                        fg=typer.colors.RED,
                        err=True,
                    )
                    raise typer.Exit(code=1)

            # Validate values is a 2D array
            if not isinstance(values, list) or not all(isinstance(row, list) for row in values):
                typer.secho(
                    "Error: Values must be a 2D array (list of lists).",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=1)

            if not values:
                typer.secho(
                    "Error: Values array cannot be empty.",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=1)

            # Convert values to CellValues and write to range
            cell_values = [[CellValue.from_python(v) for v in row] for row in values]
            engine.set_range(path, sheet, coord, cell_values)

            # Save the workbook
            workbook.save()

            rows = len(values)
            cols = len(values[0]) if values else 0
            typer.echo(f"Written {rows} row(s) x {cols} column(s) to range {coord} in sheet '{sheet}'.")

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)
