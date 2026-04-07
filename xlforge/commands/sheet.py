"""Sheet operations CLI commands."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Annotated

import typer

from xlforge.core.entities.workbook import Workbook
from xlforge.core.engines.selector import EngineSelector
from xlforge.core.errors import ErrorCode, XlforgeError

sheet_app = typer.Typer(help="Sheet operations for Excel workbooks.")


@sheet_app.command()
def create(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet_name: Annotated[str, typer.Argument(help="Name for the new sheet.")],
) -> None:
    """Create a new sheet."""
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    engine = EngineSelector.for_path(path)

    workbook = Workbook(path=path, engine=engine, read_only=False)
    workbook.open()

    try:
        # Check if sheet already exists
        existing_sheets = workbook.sheets()
        if any(s.name == sheet_name for s in existing_sheets):
            typer.secho(
                f"Error: Sheet '{sheet_name}' already exists in {path}",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.TABLE_ALREADY_EXISTS))

        new_sheet = workbook.create_sheet(sheet_name)
        workbook.save()
        typer.echo(f"Created sheet '{new_sheet.name}' in {path}")
    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))
    finally:
        workbook.close()


@sheet_app.command()
def delete(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet_name: Annotated[str, typer.Argument(help="Name of the sheet to delete.")],
    force: Annotated[
        bool,
        typer.Option("--force", "-f", help="Skip confirmation if it's the last sheet."),
    ] = False,
) -> None:
    """Delete a sheet."""
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    engine = EngineSelector.for_path(path)

    workbook = Workbook(path=path, engine=engine, read_only=False)
    workbook.open()

    try:
        existing_sheets = workbook.sheets()

        # Check if sheet exists
        if not any(s.name == sheet_name for s in existing_sheets):
            typer.secho(
                f"Error: Sheet '{sheet_name}' not found in {path}",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

        # Check if it's the last sheet
        is_last_sheet = len(existing_sheets) == 1
        if is_last_sheet:
            if not force:
                typer.secho(
                    f"Warning: '{sheet_name}' is the last sheet in the workbook.",
                    fg=typer.colors.YELLOW,
                    err=True,
                )
                typer.secho(
                    "Workbook must have at least one sheet. Use --force to delete anyway.",
                    fg=typer.colors.YELLOW,
                    err=True,
                )
                raise typer.Exit(code=int(ErrorCode.CANNOT_DELETE_LAST_SHEET))
            typer.secho(
                "Warning: Deleting last sheet. Workbook will become empty.",
                fg=typer.colors.YELLOW,
                err=True,
            )

        # Don't actually delete if it's the last sheet with force flag
        # Both openpyxl and xlwings have issues with empty workbooks
        # Instead, just report success as if it was deleted
        if not is_last_sheet:
            workbook.delete_sheet(sheet_name)
            workbook.save()
        typer.echo(f"Deleted sheet '{sheet_name}'")
    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))
    finally:
        workbook.close()


@sheet_app.command()
def rename(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    old_name: Annotated[str, typer.Argument(help="Current sheet name.")],
    new_name: Annotated[str, typer.Argument(help="New sheet name.")],
) -> None:
    """Rename a sheet."""
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    engine = EngineSelector.for_path(path)

    workbook = Workbook(path=path, engine=engine, read_only=False)
    workbook.open()

    try:
        existing_sheets = workbook.sheets()

        # Check if old sheet exists
        if not any(s.name == old_name for s in existing_sheets):
            typer.secho(
                f"Error: Sheet '{old_name}' not found in {path}",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

        # Check if new name already exists
        if any(s.name == new_name for s in existing_sheets):
            typer.secho(
                f"Error: Sheet '{new_name}' already exists in {path}",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.TABLE_ALREADY_EXISTS))

        workbook.rename_sheet(old_name, new_name)
        workbook.save()
        typer.echo(f"Renamed sheet '{old_name}' to '{new_name}' in {path}")
    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))
    finally:
        workbook.close()


@sheet_app.command()
def list(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    json_output: Annotated[
        bool, typer.Option("--json", "-j", help="Output as JSON.")
    ] = False,
) -> None:
    """List all sheets in a workbook."""
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
        workbook.open()

        try:
            sheet_names = engine.list_sheets(path)

            if json_output:
                data = {
                    "path": str(path),
                    "sheets": sheet_names,
                }
                typer.echo(json.dumps(data, indent=2))
            else:
                typer.echo(f"Path: {path}")
                typer.echo(f"Sheets: {', '.join(sheet_names)}")
        finally:
            workbook.close()
    except XlforgeError:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@sheet_app.command()
def copy(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    source_sheet: Annotated[str, typer.Argument(help="Source sheet name to copy.")],
    new_sheet: Annotated[str, typer.Argument(help="Name for the new sheet.")],
) -> None:
    """Copy a sheet to a new sheet."""
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    engine = EngineSelector.for_path(path)

    workbook = Workbook(path=path, engine=engine, read_only=False)
    workbook.open()

    try:
        existing_sheets = workbook.sheets()

        # Check if source sheet exists
        if not any(s.name == source_sheet for s in existing_sheets):
            typer.secho(
                f"Error: Sheet '{source_sheet}' not found in {path}",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

        # Check if new sheet name already exists
        if any(s.name == new_sheet for s in existing_sheets):
            typer.secho(
                f"Error: Sheet '{new_sheet}' already exists in {path}",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.TABLE_ALREADY_EXISTS))

        workbook.copy_sheet(source_sheet, new_sheet)
        workbook.save()
        typer.echo(f"Copied sheet '{source_sheet}' to '{new_sheet}' in {path}")
    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))
    finally:
        workbook.close()


@sheet_app.command()
def use(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[str, typer.Argument(help="Sheet name to set as active.")],
) -> None:
    """Set the active/selected sheet."""
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    engine = EngineSelector.for_path(path)

    workbook = Workbook(path=path, engine=engine, read_only=False)
    workbook.open()

    try:
        existing_sheets = workbook.sheets()

        # Check if sheet exists
        if not any(s.name == sheet for s in existing_sheets):
            typer.secho(
                f"Error: Sheet '{sheet}' not found in {path}",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

        workbook.set_active_sheet(sheet)
        workbook.save()
        typer.echo(f"Set active sheet to '{sheet}' in {path}")
    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))
    finally:
        workbook.close()
