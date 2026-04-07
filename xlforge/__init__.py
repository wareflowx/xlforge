import typer

from xlforge.commands.app import app_cmd
from xlforge.commands.cell import cell_app
from xlforge.commands.chart import chart_app
from xlforge.commands.context import context_app
from xlforge.commands.csv_cmd import csv_app
from xlforge.commands.file import file_app
from xlforge.commands.format_condition import format_condition_app
from xlforge.commands.named_range import named_range_app
from xlforge.commands.pivot import pivot_app
from xlforge.commands.properties import properties_app
from xlforge.commands.protection import protection_app
from xlforge.commands.range import range_app
from xlforge.commands.rowcol import col_app, row_app
from xlforge.commands.sheet import sheet_app
from xlforge.commands.style import style_app
from xlforge.commands.table import table_app
from xlforge.commands.validation import validation_app

app = typer.Typer()
app.add_typer(app_cmd, name="app")
app.add_typer(cell_app, name="cell")
app.add_typer(chart_app, name="chart")
app.add_typer(col_app, name="column")
app.add_typer(context_app, name="context")
app.add_typer(csv_app, name="csv")
app.add_typer(file_app, name="file")
app.add_typer(format_condition_app, name="format-condition")
app.add_typer(named_range_app, name="named-range")
app.add_typer(pivot_app, name="pivot")
app.add_typer(properties_app, name="properties")
app.add_typer(protection_app, name="protection")
app.add_typer(range_app, name="range")
app.add_typer(row_app, name="row")
app.add_typer(sheet_app, name="sheet")
app.add_typer(style_app, name="style")
app.add_typer(table_app, name="table")
app.add_typer(validation_app, name="validation")


@app.command()
def ping():
    """Check if xlforge is running."""
    typer.echo("pong")


@app.command()
def version():
    """Show version information."""
    typer.echo("xlforge 0.1.0")
