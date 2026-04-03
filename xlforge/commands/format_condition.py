"""Conditional formatting CLI commands for Excel workbooks."""

# Note: This command uses openpyxl directly and bypasses the Engine abstraction.
# It works with OpenpyxlEngine but not with XlwingsEngine.

from __future__ import annotations

import re
from pathlib import Path
from typing import Annotated, Optional

import openpyxl
import openpyxl.styles
import typer
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, IconSetRule, Rule

from xlforge.core.errors import ErrorCode, XlforgeError

format_condition_app = typer.Typer(help="Conditional formatting operations for Excel workbooks.")

CONDITIONAL_TYPES = ["data-bar", "color-scale", "icon-set", "formula"]
RULE_TYPES = ["greater-than", "less-than", "between", "equal", "contains"]
ICON_SETS = [
    "3Arrows", "3ArrowsGray", "3Flags", "3Signs", "3Symbols", "3Symbols2",
    "3TrafficLights1", "3TrafficLights2", "4Arrows", "4ArrowsGray",
    "4Rating", "4RedToBlack", "4TrafficLights", "5Arrows", "5ArrowsGray",
    "5Quarters", "5Rating",
]


def _is_valid_hex_color(color: str) -> bool:
    """Check if a color string is a valid hex color (#RRGGBB or RRGGBB)."""
    if color.startswith("#"):
        color = color[1:]
    return bool(re.match(r"^[0-9A-Fa-f]{6}$", color))


def _parse_style_string(style: str) -> dict:
    """Parse a style string like 'bold text-#FF0000' into font properties."""
    font_kwargs = {}
    parts = style.lower().split()

    for part in parts:
        if part == "bold":
            font_kwargs["bold"] = True
        elif part == "italic":
            font_kwargs["italic"] = True
        elif part == "underline":
            font_kwargs["underline"] = "single"
        elif part.startswith("text-"):
            color = part[5:]
            if _is_valid_hex_color(color):
                font_kwargs["color"] = "FF" + color.lstrip("#")
        elif part.startswith("bg-"):
            # Background color - not directly applicable to font
            pass
        elif _is_valid_hex_color(part.lstrip("#")):
            font_kwargs["color"] = "FF" + part.lstrip("#")

    return font_kwargs


@format_condition_app.command()
def add(
    path: Annotated[Path, typer.Argument(help="Path to the workbook file.")],
    sheet: Annotated[str, typer.Argument(help="Sheet name.")],
    range: Annotated[str, typer.Argument(help="Cell range (e.g., A1:A10).")],
    type: Annotated[
        Optional[str],
        typer.Option("--type", "-t", help=f"Conditional formatting type: {', '.join(CONDITIONAL_TYPES)}."),
    ] = None,
    rule: Annotated[
        Optional[str],
        typer.Option("--rule", "-r", help=f"Rule type (for type=rule): {', '.join(RULE_TYPES)}."),
    ] = None,
    value: Annotated[
        Optional[str],
        typer.Option("--value", "-v", help="Value for rule-based conditional formatting."),
    ] = None,
    formula: Annotated[
        Optional[str],
        typer.Option("--formula", "-f", help="Formula for formula-based conditional formatting."),
    ] = None,
    style: Annotated[
        Optional[str],
        typer.Option("--style", "-s", help="Style string (e.g., 'bold text-#FF0000' or 'italic bg-#FFFF00')."),
    ] = None,
    min: Annotated[
        Optional[str],
        typer.Option("--min", help="Min color for color scale (e.g., #FF0000)."),
    ] = None,
    max: Annotated[
        Optional[str],
        typer.Option("--max", help="Max color for color scale (e.g., #00FF00)."),
    ] = None,
    mid: Annotated[
        Optional[str],
        typer.Option("--mid", help="Mid color for color scale (e.g., #FFFF00)."),
    ] = None,
    icons: Annotated[
        Optional[str],
        typer.Option("--icons", "-i", help=f"Icon set name for icon-set type: {', '.join(ICON_SETS[:10])}..."),
    ] = None,
) -> None:
    """Add conditional formatting to a range."""
    # Check if file exists
    if not path.exists():
        typer.secho(
            f"Error: File does not exist: {path}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.FILE_DOES_NOT_EXIST))

    # Determine the actual type to use
    actual_type = type
    if rule is not None:
        actual_type = "rule"

    # Validate based on what was provided
    if actual_type is None:
        # Neither --type nor --rule provided
        if value is not None:
            # User provided --value but forgot --rule
            typer.secho(
                "Error: --rule is required when using --value. Valid rules: greater-than, less-than, between, equal, contains.",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.CONDITIONAL_FORMAT_NOT_SUPPORTED))
        typer.secho(
            "Error: --type is required. Use --type data-bar, color-scale, icon-set, or formula.",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.CONDITIONAL_FORMAT_NOT_SUPPORTED))

    if actual_type not in CONDITIONAL_TYPES and actual_type != "rule":
        typer.secho(
            f"Error: Invalid type: {type}. Valid types: {', '.join(CONDITIONAL_TYPES)}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=int(ErrorCode.CONDITIONAL_FORMAT_NOT_SUPPORTED))

    # Validate type-specific options
    if actual_type == "data-bar":
        pass  # No additional options required
    elif actual_type == "color-scale":
        if min is None or max is None:
            typer.secho(
                "Error: --min and --max colors are required for color-scale.",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.CONDITIONAL_FORMAT_NOT_SUPPORTED))
        if not _is_valid_hex_color(min):
            typer.secho(
                f"Error: Invalid min color: {min}. Use #RRGGBB or RRGGBB.",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.INVALID_STYLE_STRING))
        if not _is_valid_hex_color(max):
            typer.secho(
                f"Error: Invalid max color: {max}. Use #RRGGBB or RRGGBB.",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.INVALID_STYLE_STRING))
        if mid is not None and not _is_valid_hex_color(mid):
            typer.secho(
                f"Error: Invalid mid color: {mid}. Use #RRGGBB or RRGGBB.",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.INVALID_STYLE_STRING))
    elif actual_type == "icon-set":
        if icons is None:
            typer.secho(
                "Error: --icons is required for icon-set type.",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.CONDITIONAL_FORMAT_NOT_SUPPORTED))
    elif actual_type == "formula":
        if formula is None:
            typer.secho(
                "Error: --formula is required for formula type.",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.CONDITIONAL_FORMAT_NOT_SUPPORTED))
    elif actual_type == "rule":
        if rule not in RULE_TYPES:
            typer.secho(
                f"Error: Invalid rule: {rule}. Valid rules: {', '.join(RULE_TYPES)}",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.CONDITIONAL_FORMAT_NOT_SUPPORTED))
        if value is None:
            typer.secho(
                "Error: --value is required for rule-based formatting.",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.INVALID_FORMULA_SYNTAX))

    try:
        wb = openpyxl.load_workbook(path)

        # Check if sheet exists
        if sheet not in wb.sheetnames:
            wb.close()
            typer.secho(
                f"Error: Sheet not found: {sheet}",
                fg=typer.colors.RED,
                err=True,
            )
            raise typer.Exit(code=int(ErrorCode.SHEET_NOT_FOUND))

        ws = wb[sheet]

        # Create the appropriate conditional formatting rule
        if actual_type == "data-bar":
            rule_obj = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color="638EC6")
            ws.conditional_formatting.add(range, rule_obj)
            typer.echo(f"Added data bar conditional formatting to range {range} on sheet '{sheet}'")

        elif actual_type == "color-scale":
            min_color = "FF" + min.lstrip("#")
            max_color = "FF" + max.lstrip("#")
            if mid is not None:
                mid_color = "FF" + mid.lstrip("#")
                rule_obj = ColorScaleRule(
                    start_type="num", start_value=None, start_color=min_color,
                    mid_type="num", mid_value=None, mid_color=mid_color,
                    end_type="num", end_value=None, end_color=max_color,
                )
            else:
                rule_obj = ColorScaleRule(
                    start_type="num", start_value=None, start_color=min_color,
                    end_type="num", end_value=None, end_color=max_color,
                )
            ws.conditional_formatting.add(range, rule_obj)
            typer.echo(f"Added color scale conditional formatting to range {range} on sheet '{sheet}'")

        elif actual_type == "icon-set":
            # IconSetRule needs values - provide default thresholds based on icon set type
            rule_obj = IconSetRule(
                icon_style=icons,
                type="num",
                values=["0", "33", "67", "100"],
            )
            ws.conditional_formatting.add(range, rule_obj)
            typer.echo(f"Added icon set conditional formatting ({icons}) to range {range} on sheet '{sheet}'")

        elif actual_type == "formula":
            font_kwargs = _parse_style_string(style) if style else {}
            rule_obj = Rule(type="expression", formula=[formula])
            if font_kwargs:
                rule_obj.font = openpyxl.styles.Font(**font_kwargs)
            ws.conditional_formatting.add(range, rule_obj)
            typer.echo(f"Added formula conditional formatting to range {range} on sheet '{sheet}'")
            typer.echo(f"  Formula: {formula}")
            if style:
                typer.echo(f"  Style: {style}")

        elif actual_type == "rule":
            font_kwargs = _parse_style_string(style) if style else {}
            operator_map = {
                "greater-than": "greaterThan",
                "less-than": "lessThan",
                "between": "between",
                "equal": "equal",
                "contains": "containsText",
            }
            operator = operator_map[rule]

            if rule == "between":
                # For between, value should be two values separated by comma
                if "," not in value:
                    typer.secho(
                        "Error: --value for 'between' must be two values separated by comma (e.g., '10,100').",
                        fg=typer.colors.RED,
                        err=True,
                    )
                    raise typer.Exit(code=int(ErrorCode.INVALID_FORMULA_SYNTAX))
                parts = value.split(",")
                formula1 = parts[0].strip()
                formula2 = parts[1].strip()
                rule_obj = Rule(
                    type="cellIs",
                    operator=operator,
                    formula=[formula1, formula2],
                )
            elif rule == "contains":
                formula1 = value
                formula2 = None
                rule_obj = Rule(
                    type="containsText",
                    operator="containsText",
                    formula=[f'"{formula1}"', 'LEFT(A1,FIND("{0}",A1)-1)'],
                )
                # Simpler approach for contains
                rule_obj = Rule(
                    type="containsText",
                    operator="containsText",
                    formula=[f'"{formula1}"'],
                    text=value,
                )
            else:
                formula1 = value
                formula2 = None
                rule_obj = Rule(
                    type="cellIs",
                    operator=operator,
                    formula=[formula1],
                )

            if font_kwargs:
                rule_obj.font = openpyxl.styles.Font(**font_kwargs)

            ws.conditional_formatting.add(range, rule_obj)
            typer.echo(f"Added {rule} rule to range {range} on sheet '{sheet}'")
            typer.echo(f"  Value: {value}")
            if style:
                typer.echo(f"  Style: {style}")

        wb.save(path)
        wb.close()

    except XlforgeError:
        raise
    except typer.Exit:
        raise
    except Exception as e:
        typer.secho(f"Error: {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)
