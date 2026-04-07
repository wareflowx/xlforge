"""Utility functions for xlforge."""

from __future__ import annotations

import os
from pathlib import Path


def _is_file_open_in_excel(path: Path) -> bool:
    """Check if a file is currently open in Excel.

    Uses multiple approaches to detect if Excel has the file open:
    1. Tries to open the file in exclusive mode (Windows)
    2. If win32com is available, checks if Excel has the workbook open

    Args:
        path: Path to the file to check.

    Returns:
        True if file is open in Excel, False otherwise.
    """
    if not path.exists():
        return False

    # Approach 1: Try to open the file exclusively on Windows
    # If the file is open in Excel, this will fail
    if os.name == "nt":  # Windows
        try:
            # Try to open file in exclusive mode - this will fail if Excel has it open
            with open(path, "a+b"):
                pass
            return False
        except (OSError, IOError):
            # File could not be opened exclusively - likely open in Excel
            return True

    # Approach 2: Use win32com to check if Excel has the file open
    try:
        import win32com.client

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # Check all open workbooks
        abs_path_str = str(path.absolute())
        abs_path_lower = abs_path_str.lower()

        for wb in excel.Workbooks:
            try:
                wb_full_path = str(Path(wb.FullName).absolute()).lower()
                # Check if the workbook's full path matches our file
                if wb_full_path == abs_path_lower:
                    return True
            except Exception:
                # Skip workbooks we can't access
                continue

        return False

    except ImportError:
        # win32com not available
        return False
    except Exception:
        # Any other error (Excel not installed, etc.) - assume file is not open
        return False


def _check_file_not_open_in_excel(path: Path) -> tuple[bool, str]:
    """Check if a file is open in Excel and return a user-friendly message.

    Args:
        path: Path to the file to check.

    Returns:
        Tuple of (is_blocked, message) where is_blocked is True if the file
        is open in Excel, and message is an error message to display.
    """
    if _is_file_open_in_excel(path):
        return (
            True,
            f"File is currently open in Excel: {path.name}\n"
            "Please close Excel before modifying this file."
        )
    return False, ""
