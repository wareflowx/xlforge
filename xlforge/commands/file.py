"""File operations CLI commands."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Optional

import typer

from xlforge.core.entities.workbook import Workbook
from xlforge.core.engines.selector import EngineSelector

file_app = typer.Typer(help="File operations for Excel workbooks.")


@file_app.command()
def open(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    engine: Annotated[
        Optional[str],
        typer.Option("--engine", "-e", help="Engine to use: xlwings or openpyxl."),
    ] = None,
    read_only: Annotated[
        bool, typer.Option("--read-only", "-r", help="Open in read-only mode.")
    ] = False,
) -> None:
    """Open a workbook file."""
    if engine is not None:
        engine_obj = EngineSelector.for_engine_name(engine)
    else:
        engine_obj = EngineSelector.for_path(path)

    workbook = Workbook(path=path, engine=engine_obj, read_only=read_only)
    workbook.open()

    engine_name = engine_obj.__class__.__name__.replace("Engine", "").lower()
    typer.echo(f"Opened: {path}")
    typer.echo(f"Engine: {engine_name}")


@file_app.command()
def save(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    output: Annotated[
        Optional[Path],
        typer.Option("--output", "-o", help="Output path for saving."),
    ] = None,
) -> None:
    """Save a workbook file."""
    # Auto-detect engine based on file extension
    engine = EngineSelector.for_path(path)

    workbook = Workbook(path=path, engine=engine, read_only=False)
    workbook.open()

    if output is not None:
        # For now, just save to the original path
        # The output path would require more complex logic for copy/save-as
        typer.echo(f"Output path specified: {output}")
        typer.echo("Saving to original path...")

    workbook.save()
    workbook.close()
    typer.echo(f"Saved: {path}")


@file_app.command()
def info(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    json_output: Annotated[
        bool, typer.Option("--json", "-j", help="Output as JSON.")
    ] = False,
) -> None:
    """Show information about a workbook file."""
    engine = EngineSelector.for_path(path)

    workbook = Workbook(path=path, engine=engine, read_only=True)
    workbook.open()

    try:
        sheet_names = workbook.sheets()
        engine_name = engine.__class__.__name__.replace("Engine", "").lower()

        if json_output:
            import json

            data = {
                "path": str(path),
                "engine": engine_name,
                "sheets": [s.name for s in sheet_names],
            }
            typer.echo(json.dumps(data, indent=2))
        else:
            typer.echo(f"Path: {path}")
            typer.echo(f"Engine: {engine_name}")
            typer.echo(f"Sheets: {', '.join(s.name for s in sheet_names)}")
    finally:
        workbook.close()


@file_app.command()
def close(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
) -> None:
    """Close a workbook file."""
    engine = EngineSelector.for_path(path)

    workbook = Workbook(path=path, engine=engine, read_only=True)
    workbook.open()
    workbook.close()

    typer.echo(f"Closed: {path}")
