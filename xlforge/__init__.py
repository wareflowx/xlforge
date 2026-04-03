import typer

from xlforge.commands.file import file_app

app = typer.Typer()
app.add_typer(file_app, name="file")


@app.command()
def ping():
    """Check if xlforge is running."""
    typer.echo("pong")


@app.command()
def version():
    """Show version information."""
    typer.echo("xlforge 0.1.0")
