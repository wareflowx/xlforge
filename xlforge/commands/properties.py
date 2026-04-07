"""Workbook properties CLI commands."""

# Note: This command uses openpyxl directly and bypasses the Engine abstraction.
# It works with OpenpyxlEngine but not with XlwingsEngine.

from __future__ import annotations

import json
from pathlib import Path
from typing import Annotated, Optional

import openpyxl
import typer

from xlforge.core.errors import ErrorCode

properties_app = typer.Typer(help="Workbook properties operations.")


@properties_app.command()
def get(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    json_output: Annotated[
        bool, typer.Option("--json", "-j", help="Output as JSON.")
    ] = False,
) -> None:
    """Get all workbook properties."""
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
        props = wb.properties

        properties_data = {
            "title": props.title or "",
            "author": props.creator or "",
            "subject": props.subject or "",
            "keywords": props.keywords or "",
            "comments": props.description or "",
            "created": props.created.isoformat() if props.created else None,
            "modified": props.modified.isoformat() if props.modified else None,
        }

        wb.close()

        if json_output:
            typer.echo(json.dumps(properties_data, indent=2))
        else:
            typer.echo("Workbook Properties:")
            typer.echo("-" * 40)
            if properties_data["title"]:
                typer.echo(f"  Title: {properties_data['title']}")
            if properties_data["author"]:
                typer.echo(f"  Author: {properties_data['author']}")
            if properties_data["subject"]:
                typer.echo(f"  Subject: {properties_data['subject']}")
            if properties_data["keywords"]:
                typer.echo(f"  Keywords: {properties_data['keywords']}")
            if properties_data["comments"]:
                typer.echo(f"  Comments: {properties_data['comments']}")
            if properties_data["created"]:
                typer.echo(f"  Created: {properties_data['created']}")
            if properties_data["modified"]:
                typer.echo(f"  Modified: {properties_data['modified']}")

            # If all properties are empty
            if all(not v for v in properties_data.values()):
                typer.echo("  (No properties set)")

    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)


@properties_app.command()
def set(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    title: Annotated[
        Optional[str],
        typer.Option("--title", "-t", help="Document title."),
    ] = None,
    author: Annotated[
        Optional[str],
        typer.Option("--author", "-a", help="Document author."),
    ] = None,
    subject: Annotated[
        Optional[str],
        typer.Option("--subject", "-s", help="Document subject."),
    ] = None,
    keywords: Annotated[
        Optional[str],
        typer.Option("--keywords", "-k", help="Document keywords."),
    ] = None,
    comments: Annotated[
        Optional[str],
        typer.Option("--comments", "-c", help="Document comments."),
    ] = None,
) -> None:
    """Set workbook properties."""
    # Check if file exists
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    # Check if at least one property is provided
    if all(v is None for v in [title, author, subject, keywords, comments]):
        typer.secho(
            "Error: Must provide at least one property to set.",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

    try:
        wb = openpyxl.load_workbook(path)
        props = wb.properties

        changes = []

        if title is not None:
            props.title = title
            changes.append(f"title='{title}'")

        if author is not None:
            props.creator = author
            changes.append(f"author='{author}'")

        if subject is not None:
            props.subject = subject
            changes.append(f"subject='{subject}'")

        if keywords is not None:
            props.keywords = keywords
            changes.append(f"keywords='{keywords}'")

        if comments is not None:
            props.description = comments
            changes.append(f"comments='{comments}'")

        wb.save(path)
        wb.close()

        typer.echo(f"Updated properties: {', '.join(changes)}")

    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)
