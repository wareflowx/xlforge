"""CSV import/export CLI commands."""

from __future__ import annotations

import csv
import sys
from pathlib import Path
from typing import Annotated, TextIO

import typer

from xlforge.core.entities.workbook import Workbook
from xlforge.core.engines.selector import EngineSelector
from xlforge.core.errors import ErrorCode, XlforgeError
from xlforge.core.types.cell_value import CellValue

csv_app = typer.Typer(help="CSV import/export operations for Excel workbooks.")


@csv_app.command(name="import")
def import_csv(
    csv_file: Annotated[Path, typer.Argument(help="Path to the CSV file to import.")],
    excel_file: Annotated[
        Path, typer.Argument(help="Path to the Excel workbook file.")
    ],
    sheet: Annotated[str, typer.Argument(help="Sheet name to import into.")],
    has_header: Annotated[
        bool,
        typer.Option("--has-header", "-H", help="Treat the first row as headers."),
    ] = False,
) -> None:
    """Import a CSV file into an Excel sheet."""
    # Check if CSV file exists
    if not csv_file.exists():
        typer.secho(
            f"Error: CSV file does not exist: {csv_file}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.CSV_NOT_FOUND))

    # Check if Excel file exists
    if not excel_file.exists():
        typer.secho(
            f"Error: Excel file does not exist: {excel_file}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        engine = EngineSelector.for_path(excel_file)
        workbook = Workbook(path=excel_file, engine=engine, read_only=False)

        with workbook:
            # Check if sheet exists
            if not engine.sheet_exists(excel_file, sheet):
                typer.secho(
                    f"Error: Sheet not found: {sheet}",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            # Read CSV data
            with open(csv_file, "r", newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)

            if not rows:
                typer.secho(
                    "Error: CSV file is empty.",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.INVALID_CSV_FORMAT))

            # Determine starting row
            start_row = 0
            if has_header and len(rows) > 1:
                start_row = 1

            # Convert CSV rows to CellValue 2D array with type coercion
            data_rows = rows[start_row:]
            cell_values: list[list[CellValue]] = []
            for row in data_rows:
                cell_row = [CellValue.from_string(cell) for cell in row]
                cell_values.append(cell_row)

            # Calculate range
            if cell_values:
                num_cols = max(len(row) for row in cell_values)
                num_rows = len(cell_values)
                start_coord = "A1"
                end_coord = f"{_column_letter(num_cols)}{num_rows}"
                coord_range = f"{start_coord}:{end_coord}"

                # Write to sheet using engine directly
                engine.set_range(excel_file, sheet, coord_range, cell_values)

                workbook.save()
                typer.echo(
                    f"Imported {num_rows} row(s) and {num_cols} column(s) from {csv_file} to sheet '{sheet}'"
                )
            else:
                typer.echo("No data to import.")

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except csv.Error as e:
        typer.secho(
            f"Error: Invalid CSV format: {e}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.INVALID_CSV_FORMAT))
    except UnicodeDecodeError as e:
        typer.secho(
            f"Error: Encoding error reading CSV: {e}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.ENCODING_ERROR))
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@csv_app.command()
def export(
    excel_file: Annotated[
        Path, typer.Argument(help="Path to the Excel workbook file.")
    ],
    sheet: Annotated[str, typer.Argument(help="Sheet name to export from.")],
    range_spec: Annotated[
        str | None,
        typer.Option(
            "--range",
            "-r",
            help="Range to export (e.g., A1:C10). If not specified, exports entire used range.",
        ),
    ] = None,
    output: Annotated[
        Path | None,
        typer.Option(
            "--output",
            "-o",
            help="Output CSV file. If not specified, outputs to stdout.",
        ),
    ] = None,
) -> None:
    """Export an Excel range/sheet to CSV."""
    # Check if Excel file exists
    if not excel_file.exists():
        typer.secho(
            f"Error: Excel file does not exist: {excel_file}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        engine = EngineSelector.for_path(excel_file)
        workbook = Workbook(path=excel_file, engine=engine, read_only=False)

        with workbook:
            # Check if sheet exists
            if not engine.sheet_exists(excel_file, sheet):
                typer.secho(
                    f"Error: Sheet not found: {sheet}",
                    fg=typer.colors.RED,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            sheet_obj = workbook.sheet(name=sheet)

            # Determine range
            if range_spec:
                coord_range = range_spec
            else:
                # Use used range
                dimensions = engine.get_sheet_dimensions(excel_file, sheet)
                coord_range = dimensions if dimensions else "A1"

            # Read data from Excel
            cell_values = engine.get_range(excel_file, sheet, coord_range)

            if not cell_values:
                typer.echo("No data to export.")
                return

            # Convert to CSV
            output_io: TextIO
            if output:
                f = open(output, "w", newline="", encoding="utf-8")
                output_io = f
            else:
                output_io = sys.stdout

            try:
                writer = csv.writer(output_io)
                for row in cell_values:
                    csv_row = [_cell_value_to_string(cell) for cell in row]
                    writer.writerow(csv_row)

                if output:
                    typer.echo(f"Exported to {output}")
            finally:
                if output:
                    f.close()

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except csv.Error as e:
        typer.secho(
            f"Error: Invalid CSV format: {e}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.INVALID_CSV_FORMAT))
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


def _column_letter(col_index: int) -> str:
    """Convert 1-based column index to Excel column letter (A, B, ..., Z, AA, AB, ...)."""
    result = ""
    while col_index > 0:
        col_index -= 1
        result = chr(ord("A") + col_index % 26) + result
        col_index //= 26
    return result


def _cell_value_to_string(cell: CellValue) -> str:
    """Convert CellValue to string for CSV output."""
    if cell.type.value == "empty":
        return ""
    if cell.type.value == "error":
        return str(cell.raw) if cell.raw else ""
    return cell.as_string()
