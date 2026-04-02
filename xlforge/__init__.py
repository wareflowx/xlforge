import typer

app = typer.Typer()


@app.command()
def ping():
    """Check if xlforge is running."""
    typer.echo("pong")


@app.command()
def version():
    """Show version information."""
    typer.echo("xlforge 0.1.0")
