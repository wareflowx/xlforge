import typer

from xlforge.commands.cell import cell_app
from xlforge.commands.csv_cmd import csv_app
from xlforge.commands.file import file_app
from xlforge.commands.range import range_app
from xlforge.commands.sheet import sheet_app

app = typer.Typer()
app.add_typer(cell_app, name="cell")
app.add_typer(csv_app, name="csv")
app.add_typer(file_app, name="file")
app.add_typer(range_app, name="range")
app.add_typer(sheet_app, name="sheet")


@app.command()
def ping():
    """Check if xlforge is running."""
    typer.echo("pong")


@app.command()
def version():
    """Show version information."""
    typer.echo("xlforge 0.1.0")
