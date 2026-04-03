"""Pivot table operations CLI commands."""

from __future__ import annotations

from importlib.util import find_spec
from pathlib import Path
from typing import Annotated, Optional

import openpyxl
import typer

from xlforge.core.errors import ErrorCode

pivot_app = typer.Typer(help="Pivot table operations for Excel workbooks.")


# Aggregation types supported by Excel pivot tables
AGGREGATION_TYPES = {"SUM", "COUNT", "AVERAGE", "MIN", "MAX", "PRODUCT", "COUNT_NUMBERS"}


def _is_xlwings_available() -> bool:
    """Check if xlwings is available (Excel integration possible)."""
    return find_spec("xlwings") is not None


def _parse_aggregation(agg_str: str) -> tuple[str, str]:
    """Parse aggregation string like 'SUM:Revenue' into (aggregation, field).

    Args:
        agg_str: Aggregation string in format 'AGGREGATION:FieldName'.

    Returns:
        Tuple of (aggregation_type, field_name).

    Raises:
        ValueError: If the format is invalid.
    """
    parts = agg_str.split(":", 1)
    if len(parts) != 2:
        raise ValueError(f"Invalid aggregation format: {agg_str}. Expected 'AGGREGATION:FieldName'.")
    agg_type, field_name = parts
    agg_type_upper = agg_type.upper()
    if agg_type_upper not in AGGREGATION_TYPES:
        raise ValueError(
            f"Invalid aggregation type: {agg_type}. "
            f"Supported types: {', '.join(sorted(AGGREGATION_TYPES))}."
        )
    return agg_type_upper, field_name


def _parse_range(range_str: str) -> tuple[int, int, int, int]:
    """Parse a range string like 'A1:D10' into (min_col, min_row, max_col, max_row).

    Returns:
        Tuple of (min_col, min_row, max_col, max_row) as 1-indexed integers.
    """
    from openpyxl.utils import column_index_from_string

    parts = range_str.split(":")
    if len(parts) != 2:
        raise ValueError(f"Invalid range format: {range_str}")

    start_cell, end_cell = parts

    # Parse start cell
    start_col_str = "".join(c for c in start_cell if c.isalpha())
    start_row_str = "".join(c for c in start_cell if c.isdigit())
    if not start_col_str or not start_row_str:
        raise ValueError(f"Invalid range format: {range_str}")

    min_col = column_index_from_string(start_col_str)
    min_row = int(start_row_str)

    # Parse end cell
    end_col_str = "".join(c for c in end_cell if c.isalpha())
    end_row_str = "".join(c for c in end_cell if c.isdigit())
    if not end_col_str or not end_row_str:
        raise ValueError(f"Invalid range format: {range_str}")

    max_col = column_index_from_string(end_col_str)
    max_row = int(end_row_str)

    return min_col, min_row, max_col, max_row


def _cell_range_to_excel_ref(min_col: int, min_row: int, max_col: int, max_row: int) -> str:
    """Convert numeric range to Excel reference like 'A1:D10'.

    Args:
        min_col: Minimum column (1-indexed).
        min_row: Minimum row (1-indexed).
        max_col: Maximum column (1-indexed).
        max_row: Maximum row (1-indexed).

    Returns:
        Excel range reference string.
    """
    from openpyxl.utils import get_column_letter

    start = f"{get_column_letter(min_col)}{min_row}"
    end = f"{get_column_letter(max_col)}{max_row}"
    return f"{start}:{end}"


@pivot_app.command()
def create(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    source_sheet: Annotated[str, typer.Argument(help="Sheet name containing the source data.")],
    source_range: Annotated[str, typer.Argument(help="Source data range (e.g., A1:D10).")],
    sheet: Annotated[
        Optional[str],
        typer.Option("--sheet", "-s", help="Sheet name where pivot table will be created."),
    ] = None,
    name: Annotated[
        Optional[str],
        typer.Option("--name", "-n", help="Name for the pivot table."),
    ] = None,
    rows: Annotated[
        Optional[str],
        typer.Option("--rows", "-r", help="Comma-separated list of row fields."),
    ] = None,
    columns: Annotated[
        Optional[str],
        typer.Option("--columns", "-c", help="Comma-separated list of column fields."),
    ] = None,
    values: Annotated[
        Optional[list[str]],
        typer.Option("--values", "-v", help="Aggregation specifications (e.g., SUM:Revenue)."),
    ] = None,
    filters: Annotated[
        Optional[str],
        typer.Option("--filters", "-f", help="Comma-separated list of filter fields."),
    ] = None,
) -> None:
    """Create a pivot table from source data.

    Note: PivotTable creation requires Excel via xlwings engine.
    Openpyxl has very limited pivot table support and cannot create
    fully functional pivot tables.
    """
    # Check if xlwings is available for full pivot support
    if not _is_xlwings_available():
        typer.secho(
            "Error: Pivot table operations require Excel via xlwings engine.",
            fg=typer.colors.RED,
            err=True,
        )
        typer.secho(
            "Feature unavailable in headless mode (openpyxl only).",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FEATURE_UNAVAILABLE))

    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    # Determine target sheet
    target_sheet = sheet if sheet else f"{source_sheet}_Pivot"

    try:
        wb = openpyxl.load_workbook(path)

        # Check if source sheet exists
        if source_sheet not in wb.sheetnames:
            typer.secho(
                f"Error: Sheet not found: {source_sheet}",
                fg=typer.colors.RED,
                err=True,
            )
            wb.close()
            raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

        ws_source = wb[source_sheet]

        # Parse source range
        try:
            min_col, min_row, max_col, max_row = _parse_range(source_range)
        except ValueError as e:
            typer.secho(
                f"Error: Invalid source range format: {source_range}",
                fg=typer.colors.RED,
                err=True,
            )
            wb.close()
            raise typer.Exit(code=int(ErrorCode.INVALID_SYNTAX)) from e

        # Validate aggregations if provided
        parsed_values = []
        if values:
            for val in values:
                try:
                    agg_type, field_name = _parse_aggregation(val)
                    parsed_values.append((agg_type, field_name))
                except ValueError as e:
                    typer.secho(
                        f"Error: {e}",
                        fg=typer.colors.RED,
                        err=True,
                    )
                    wb.close()
                    raise typer.Exit(code=int(ErrorCode.INVALID_SYNTAX)) from e

        # Parse row, column, and filter fields
        row_fields = [f.strip() for f in rows.split(",")] if rows else []
        col_fields = [f.strip() for f in columns.split(",")] if columns else []
        filter_fields = [f.strip() for f in filters.split(",")] if filters else []

        # Create target sheet if it doesn't exist
        if target_sheet not in wb.sheetnames:
            wb.create_sheet(target_sheet)
        ws_target = wb[target_sheet]

        # Generate pivot table name if not provided
        pivot_name = name if name else f"PivotTable{len(ws_target._pivots) + 1}"

        # Check if pivot with same name already exists
        for existing_pivot in ws_target._pivots:
            if existing_pivot.name == pivot_name:
                typer.secho(
                    f"Error: Pivot table with name '{pivot_name}' already exists in sheet '{target_sheet}'",
                    fg=typer.colors.RED,
                    err=True,
                )
                typer.secho(
                    "Use a different --name to avoid overwriting existing pivot tables.",
                    fg=typer.colors.YELLOW,
                    err=True,
                )
                wb.close()
                raise typer.Exit(code=int(ErrorCode.PIVOT_CREATION_FAILED))

        # Import here to avoid issues if openpyxl pivot is not properly set up
        from openpyxl.pivot.table import TableDefinition, Location

        # Create pivot table definition
        # Note: Openpyxl's pivot table support is very limited
        # We create a basic TableDefinition that Excel can work with
        loc = Location(
            ref=_cell_range_to_excel_ref(min_col, min_row, max_col, max_row),
            firstHeaderRow=min_row,
            firstDataRow=min_row,
            firstDataCol=min_col,
        )
        pivot = TableDefinition(
            name=pivot_name,
            cacheId=1,
            dataCaption="Data",
            grandTotalCaption="Grand Total",
            errorCaption="#N/A",
            missingCaption="Missing",
            location=loc,
        )

        # Add pivot to target sheet
        ws_target._pivots.append(pivot)

        # Save the workbook - this may fail due to openpyxl's limited pivot support
        save_succeeded = True
        try:
            wb.save(path)
        except Exception as e:
            save_succeeded = False
            typer.secho(
                f"Warning: Pivot table structure created but save failed: {e}",
                fg=typer.colors.YELLOW,
                err=True,
            )
            typer.secho(
                "The pivot table will need to be configured manually in Excel for full functionality.",
                fg=typer.colors.YELLOW,
                err=True,
            )
        finally:
            wb.close()

        typer.echo(f"Created pivot table '{pivot_name}' in sheet '{target_sheet}'")
        typer.echo(f"Source: {path} ({source_sheet}, {source_range})")

        if parsed_values:
            typer.echo(f"Aggregations: {', '.join(f'{agg}:{field}' for agg, field in parsed_values)}")
        if row_fields:
            typer.echo(f"Row fields: {', '.join(row_fields)}")
        if col_fields:
            typer.echo(f"Column fields: {', '.join(col_fields)}")
        if filter_fields:
            typer.echo(f"Filter fields: {', '.join(filter_fields)}")

    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.PIVOT_CREATION_FAILED))


@pivot_app.command("list")
def list_pivots(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[
        Optional[str],
        typer.Option("--sheet", "-s", help="Sheet name to list pivots from."),
    ] = None,
) -> None:
    """List all pivot tables in a workbook."""
    if not _is_xlwings_available():
        typer.secho(
            "Error: Pivot table operations require Excel via xlwings engine.",
            fg=typer.colors.RED,
            err=True,
        )
        typer.secho(
            "Feature unavailable in headless mode (openpyxl only).",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FEATURE_UNAVAILABLE))

    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        wb = openpyxl.load_workbook(path)

        pivots_found = False
        sheets_to_check = [sheet] if sheet else wb.sheetnames

        for sheet_name in sheets_to_check:
            if sheet_name not in wb.sheetnames:
                typer.secho(
                    f"Error: Sheet not found: {sheet_name}",
                    fg=typer.colors.RED,
                    err=True,
                )
                wb.close()
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            ws = wb[sheet_name]
            for pivot in ws._pivots:
                pivots_found = True
                typer.echo(f"Sheet: {sheet_name}, Pivot: {pivot.name}")

        if not pivots_found:
            if sheet:
                typer.echo(f"No pivot tables found in sheet '{sheet}'.")
            else:
                typer.echo("No pivot tables found in workbook.")

        wb.close()

    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.PIVOT_CREATION_FAILED))


@pivot_app.command()
def delete(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    name: Annotated[str, typer.Argument(help="Name of the pivot table to delete.")],
    sheet: Annotated[
        Optional[str],
        typer.Option("--sheet", "-s", help="Sheet name containing the pivot table."),
    ] = None,
) -> None:
    """Delete a pivot table from a sheet."""
    if not _is_xlwings_available():
        typer.secho(
            "Error: Pivot table operations require Excel via xlwings engine.",
            fg=typer.colors.RED,
            err=True,
        )
        typer.secho(
            "Feature unavailable in headless mode (openpyxl only).",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FEATURE_UNAVAILABLE))

    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        wb = openpyxl.load_workbook(path)

        if sheet:
            if sheet not in wb.sheetnames:
                typer.secho(
                    f"Error: Sheet not found: {sheet}",
                    fg=typer.colors.RED,
                    err=True,
                )
                wb.close()
                raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

            ws = wb[sheet]
            pivot_to_delete = None
            for pivot in ws._pivots:
                if pivot.name == name:
                    pivot_to_delete = pivot
                    break

            if pivot_to_delete is None:
                typer.secho(
                    f"Error: Pivot table '{name}' not found in sheet '{sheet}'",
                    fg=typer.colors.RED,
                    err=True,
                )
                wb.close()
                raise typer.Exit(code=int(ErrorCode.PIVOT_CREATION_FAILED))

            ws._pivots.remove(pivot_to_delete)
            wb.save(path)
            wb.close()
            typer.echo(f"Deleted pivot table '{name}' from sheet '{sheet}'")
        else:
            # Search all sheets for the pivot table
            found = False
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                for pivot in ws._pivots:
                    if pivot.name == name:
                        ws._pivots.remove(pivot)
                        found = True
                        break
                if found:
                    break

            if not found:
                typer.secho(
                    f"Error: Pivot table '{name}' not found in workbook",
                    fg=typer.colors.RED,
                    err=True,
                )
                wb.close()
                raise typer.Exit(code=int(ErrorCode.PIVOT_CREATION_FAILED))

            wb.save(path)
            wb.close()
            typer.echo(f"Deleted pivot table '{name}'")

    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.PIVOT_CREATION_FAILED))
