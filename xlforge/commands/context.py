"""Context management CLI commands for xlforge."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Annotated, Optional

import typer

CONTEXT_FILE = Path.home() / ".xlforge" / "context.json"

context_app = typer.Typer(help="Manage default context for xlforge commands.")


def get_context() -> dict:
    """Read the current context from the context file.

    Returns:
        Dictionary with 'file' and 'sheet' keys, or empty dict if no context set.
    """
    if CONTEXT_FILE.exists():
        try:
            return json.loads(CONTEXT_FILE.read_text())
        except (json.JSONDecodeError, IOError):
            return {}
    return {}


def set_context(file: Path, sheet: str | None = None) -> None:
    """Write the context to the context file.

    Args:
        file: Path to the default workbook file.
        sheet: Optional default sheet name.
    """
    CONTEXT_FILE.parent.mkdir(parents=True, exist_ok=True)
    context = {"file": str(file), "sheet": sheet}
    CONTEXT_FILE.write_text(json.dumps(context, indent=2))


def clear_context() -> None:
    """Clear the context by removing the context file."""
    if CONTEXT_FILE.exists():
        CONTEXT_FILE.unlink()


@context_app.command()
def set(
    file: Annotated[Path, typer.Argument(help="Path to the default workbook file.")],
    sheet: Annotated[
        Optional[str],
        typer.Option("--sheet", "-s", help="Default sheet name."),
    ] = None,
) -> None:
    """Set the default context (file and optional sheet)."""
    try:
        # Validate the file path is absolute or can be resolved
        resolved_path = file.resolve() if not file.is_absolute() else file

        set_context(resolved_path, sheet)
        if sheet:
            typer.echo(f"Context set: file={resolved_path}, sheet={sheet}")
        else:
            typer.echo(f"Context set: file={resolved_path}")

    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@context_app.command()
def show() -> None:
    """Show the current context."""
    context = get_context()

    if not context:
        typer.echo("No context is set.")
        return

    file_path = context.get("file", "N/A")
    sheet = context.get("sheet")

    typer.echo(f"File: {file_path}")
    if sheet:
        typer.echo(f"Sheet: {sheet}")
    else:
        typer.echo("Sheet: (not set)")


@context_app.command()
def clear() -> None:
    """Clear the current context."""
    if not CONTEXT_FILE.exists():
        typer.echo("No context to clear.")
        return

    try:
        clear_context()
        typer.echo("Context cleared.")
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)
