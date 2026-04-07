"""App-level Excel commands using xlwings for full Excel control."""

from __future__ import annotations

from importlib.util import find_spec
from pathlib import Path
from typing import Annotated, Any, Optional

import typer

from xlforge.core.errors import ErrorCode, XlforgeError

app_cmd = typer.Typer(help="App-level Excel commands (require xlwings).")


def _check_xlwings_available() -> bool:
    """Check if xlwings is available."""
    return find_spec("xlwings") is not None


def _get_xlwings_app(path: Path) -> Any:
    """Open workbook with xlwings and return the app object.

    Raises:
        XlforgeError: If xlwings is not available or file not found.
    """
    if not _check_xlwings_available():
        raise XlforgeError(
            code=ErrorCode.FEATURE_UNAVAILABLE,
            message="xlwings is not available. Install xlwings: pip install xlwings",
        )

    if not path.exists():
        raise XlforgeError(
            code=ErrorCode.FILE_NOT_FOUND,
            message=f"File not found: {path}",
        )

    import xlwings as xw

    # Open workbook with xlwings
    wb = xw.Book(path)
    return wb.app


@app_cmd.command()
def visible(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    value: Annotated[str, typer.Argument(help="True to show, False to hide.")],
) -> None:
    """Show or hide the Excel window."""
    if value.lower() not in ("true", "false"):
        typer.secho(
            f"Error: Invalid value '{value}'. Use 'true' or 'false'.",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.INVALID_SYNTAX))

    try:
        app = _get_xlwings_app(path)
        app.visible = value.lower() == "true"
        state = "shown" if app.visible else "hidden"
        typer.echo(f"Excel window is now {state}")
    except XlforgeError as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(e.code))
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))


@app_cmd.command()
def calculate(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
) -> None:
    """Force recalculation of all formulas in the workbook."""
    try:
        app = _get_xlwings_app(path)
        app.calculate()
        typer.echo("Recalculation complete")
    except XlforgeError as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(e.code))
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))


@app_cmd.command()
def focus(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
) -> None:
    """Bring Excel to the foreground and activate its window."""
    try:
        app = _get_xlwings_app(path)
        app.activate()
        typer.echo("Excel window activated")
    except XlforgeError as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(e.code))
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))


@app_cmd.command()
def alert(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    message: Annotated[
        str, typer.Argument(help="Message to display in the alert dialog.")
    ],
) -> None:
    """Show an alert dialog in Excel."""
    try:
        app = _get_xlwings_app(path)
        app.dialog("msg", message)
        typer.echo(f"Alert shown: {message}")
    except XlforgeError as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(e.code))
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))


@app_cmd.command()
def wait_idle(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    timeout: Annotated[
        Optional[int],
        typer.Option(
            "--timeout", "-t", help="Timeout in seconds. Use 0 for no timeout."
        ),
    ] = None,
) -> None:
    """Wait for Excel to finish all pending calculations."""
    try:
        app = _get_xlwings_app(path)
        if timeout is not None and timeout > 0:
            app.wait_idle(timeout=timeout)
        else:
            app.wait_idle()
        typer.echo("Excel is idle")
    except XlforgeError as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(e.code))
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=int(ErrorCode.GENERAL_ERROR))
