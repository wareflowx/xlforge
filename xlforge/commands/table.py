"""Table operations CLI commands."""

# Note: This command uses openpyxl directly and bypasses the Engine abstraction.
# It works with OpenpyxlEngine but not with XlwingsEngine.

from __future__ import annotations

from pathlib import Path
from typing import Annotated

import openpyxl
import typer
from openpyxl.worksheet.table import Table

from xlforge.core.errors import ErrorCode

table_app = typer.Typer(help="Table operations for Excel workbooks.")


def _validate_table_name(name: str) -> bool:
    """Validate table name according to Excel rules."""
    # Table names must not be empty, exceed 255 characters,
    # or contain certain special characters
    if not name or len(name) > 255:
        return False
    # Table names cannot contain: : \ / ? * [ ]
    invalid_chars = {":", "\\", "/", "?", "*", "[", "]"}
    return not any(c in invalid_chars for c in name)


@table_app.command()
def create(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[str, typer.Argument(help="Sheet name containing the range.")],
    range_ref: Annotated[str, typer.Argument(help="Range reference (e.g., A1:C10).")],
    name: Annotated[str | None, typer.Option("--name", "-n", help="Name for the table.")] = None,
) -> None:
    """Create an Excel table from a range."""
    # Check if file exists
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    # Validate table name if provided
    if name and not _validate_table_name(name):
        typer.secho(
            f"Error: Invalid table name: {name}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.INVALID_TABLE_NAME))

    try:
        wb = openpyxl.load_workbook(path)

        # Check if sheet exists
        if sheet not in wb.sheetnames:
            typer.secho(
                f"Error: Sheet not found: {sheet}",
                fg=typer.colors.RED,
                err=True,
            )
            wb.close()
            raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

        ws = wb[sheet]

        # If no name provided, use a default name
        if not name:
            # Generate a unique name
            base_name = f"Table{len(ws.tables) + 1}"
            counter = 1
            while base_name in ws.tables:
                counter += 1
                base_name = f"Table{len(ws.tables) + counter}"
            name = base_name

        # Check if table already exists with this name
        if name in ws.tables:
            typer.secho(
                f"Error: Table '{name}' already exists in sheet '{sheet}'",
                fg=typer.colors.RED,
                err=True,
            )
            wb.close()
            raise typer.Exit(code=int(ErrorCode.TABLE_ALREADY_EXISTS))

        # Create the table
        table = Table(displayName=name, ref=range_ref)
        ws.add_table(table)

        wb.save(path)
        wb.close()

        typer.echo(f"Created table '{name}' at {range_ref} in sheet '{sheet}'")

    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))


@table_app.command("list")
def list_tables(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
) -> None:
    """List all tables in the workbook."""
    # Check if file exists
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        wb = openpyxl.load_workbook(path)

        tables_found = False
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for tbl in ws.tables.values():
                tables_found = True
                typer.echo(f"Sheet: {sheet_name}, Table: {tbl.name}, Range: {tbl.ref}")

        if not tables_found:
            typer.echo("No tables found in workbook.")

        wb.close()

    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))


@table_app.command()
def delete(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    name: Annotated[str, typer.Argument(help="Name of the table to delete.")],
    sheet: Annotated[str, typer.Option("--sheet", "-s", help="Sheet name containing the table.")] = None,
) -> None:
    """Delete a table from the workbook."""
    # Check if file exists
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    try:
        wb = openpyxl.load_workbook(path)

        # If sheet specified, only look in that sheet
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
            if name not in ws.tables:
                typer.secho(
                    f"Error: Table '{name}' not found in sheet '{sheet}'",
                    fg=typer.colors.RED,
                    err=True,
                )
                wb.close()
                raise typer.Exit(code=int(ErrorCode.TABLE_NOT_FOUND))

            del ws.tables[name]
            wb.save(path)
            wb.close()
            typer.echo(f"Deleted table '{name}' from sheet '{sheet}'")
        else:
            # Search all sheets for the table
            found = False
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                if name in ws.tables:
                    del ws.tables[name]
                    found = True
                    break

            if not found:
                typer.secho(
                    f"Error: Table '{name}' not found in workbook",
                    fg=typer.colors.RED,
                    err=True,
                )
                wb.close()
                raise typer.Exit(code=int(ErrorCode.TABLE_NOT_FOUND))

            wb.save(path)
            wb.close()
            typer.echo(f"Deleted table '{name}'")

    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))
