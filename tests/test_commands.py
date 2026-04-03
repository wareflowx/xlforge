import os
import tempfile

import openpyxl
from typer.testing import CliRunner

from xlforge import app
from xlforge.core.errors import ErrorCode

runner = CliRunner()


def test_ping():
    result = runner.invoke(app, ["ping"])
    assert result.exit_code == 0
    assert "pong" in result.output


def test_version():
    result = runner.invoke(app, ["version"])
    assert result.exit_code == 0
    assert "xlforge 0.1.0" in result.output


def test_no_args_shows_help():
    result = runner.invoke(app)
    # Without invoke_without_command, no args shows help/missing command error
    assert result.exit_code in [0, 2]


class TestSheetCommands:
    """Tests for sheet commands."""

    def test_sheet_create(self):
        """Test creating a new sheet."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create a workbook with default sheet
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["sheet", "create", path, "NewSheet"])

            assert result.exit_code == 0
            assert "Created sheet 'NewSheet'" in result.output

            # Verify sheet exists
            wb = openpyxl.load_workbook(path)
            assert "NewSheet" in wb.sheetnames
            wb.close()

    def test_sheet_create_already_exists(self):
        """Test creating a sheet that already exists."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create a workbook with default sheet
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["sheet", "create", path, "Sheet"])

            # Sheet already exists by default
            assert result.exit_code == ErrorCode.TABLE_ALREADY_EXISTS
            assert "already exists" in result.output

    def test_sheet_create_file_not_found(self):
        """Test creating a sheet in non-existent file."""
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "nonexistent.xlsx")
            result = runner.invoke(app, ["sheet", "create", path, "NewSheet"])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_sheet_delete(self):
        """Test deleting a sheet."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create a workbook with two sheets
            wb = openpyxl.Workbook()
            wb.create_sheet("SecondSheet")
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["sheet", "delete", path, "Sheet"])

            assert result.exit_code == 0
            assert "Deleted sheet 'Sheet'" in result.output

            # Verify sheet was deleted
            wb = openpyxl.load_workbook(path)
            assert "Sheet" not in wb.sheetnames
            assert "SecondSheet" in wb.sheetnames
            wb.close()

    def test_sheet_delete_not_found(self):
        """Test deleting a non-existent sheet."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["sheet", "delete", path, "NonExistent"])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "not found" in result.output.lower()

    def test_sheet_delete_last_sheet_warns(self):
        """Test deleting the last sheet shows warning."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["sheet", "delete", path, "Sheet"])

            assert result.exit_code == ErrorCode.CANNOT_DELETE_LAST_SHEET
            assert "last sheet" in result.output.lower() or "warning" in result.output.lower()

    def test_sheet_delete_last_sheet_force(self):
        """Test force deleting the last sheet."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["sheet", "delete", path, "Sheet", "--force"])

            # Note: --force allows deleting last sheet, but openpyxl cannot save
            # an empty workbook, so the file remains unchanged on disk
            assert result.exit_code == 0
            assert "Deleted sheet 'Sheet'" in result.output

    def test_sheet_rename(self):
        """Test renaming a sheet."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["sheet", "rename", path, "Sheet", "RenamedSheet"])

            assert result.exit_code == 0
            assert "Renamed sheet 'Sheet' to 'RenamedSheet'" in result.output

            # Verify sheet was renamed
            wb = openpyxl.load_workbook(path)
            assert "RenamedSheet" in wb.sheetnames
            assert "Sheet" not in wb.sheetnames
            wb.close()

    def test_sheet_rename_not_found(self):
        """Test renaming a non-existent sheet."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["sheet", "rename", path, "NonExistent", "NewName"])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "not found" in result.output.lower()

    def test_sheet_rename_new_name_exists(self):
        """Test renaming to an existing sheet name."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.create_sheet("ExistingSheet")
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["sheet", "rename", path, "Sheet", "ExistingSheet"])

            assert result.exit_code == ErrorCode.TABLE_ALREADY_EXISTS
            assert "already exists" in result.output.lower()


class TestCellRead:
    """Tests for xlforge cell read command."""

    def test_cell_read_file_not_found(self):
        """Test cell read with non-existent file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "nonexistent.xlsx")
            result = runner.invoke(app, ["cell", "read", path, "Sheet1", "A1"])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_cell_read_sheet_not_found(self):
        """Test cell read with non-existent sheet returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "read", path, "NonexistentSheet", "A1"])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "Sheet not found" in result.output

    def test_cell_read_string_value(self):
        """Test reading a string cell value."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A1"] = "Hello World"
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "read", path, "Sheet1", "A1"])

            assert result.exit_code == 0
            assert "Hello World" in result.output
            assert "string" in result.output

    def test_cell_read_number_value(self):
        """Test reading a number cell value."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["B2"] = 42.5
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "read", path, "Sheet1", "B2"])

            assert result.exit_code == 0
            assert "42.5" in result.output
            assert "number" in result.output

    def test_cell_read_boolean_value(self):
        """Test reading a boolean cell value."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["C3"] = True
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "read", path, "Sheet1", "C3"])

            assert result.exit_code == 0
            assert "True" in result.output
            assert "bool" in result.output

    def test_cell_read_json_output(self):
        """Test reading a cell with JSON output."""
        import json

        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A1"] = "Test"
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "read", path, "Sheet1", "A1", "--json"])

            assert result.exit_code == 0
            data = json.loads(result.output)
            assert data["value"] == "Test"
            assert data["type"] == "string"
            assert data["coord"] == "A1"
            assert data["sheet"] == "Sheet1"

    def test_cell_read_empty_cell(self):
        """Test reading an empty cell."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            # A1 is empty by default
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "read", path, "Sheet1", "A1"])

            assert result.exit_code == 0
            assert "None" in result.output or "empty" in result.output.lower()


class TestCellWrite:
    """Tests for xlforge cell write command."""

    def test_cell_write_file_not_found(self):
        """Test cell write with non-existent file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "nonexistent.xlsx")
            result = runner.invoke(app, ["cell", "write", path, "Sheet1", "A1", "test"])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_cell_write_sheet_not_found(self):
        """Test cell write with non-existent sheet returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "write", path, "NonexistentSheet", "A1", "test"])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "Sheet not found" in result.output

    def test_cell_write_string_value(self):
        """Test writing a string value to a cell."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "write", path, "Sheet", "A1", "Hello World"])

            assert result.exit_code == 0
            assert "Written:" in result.output

            # Verify the value was written
            wb = openpyxl.load_workbook(path)
            assert wb.active["A1"].value == "Hello World"
            wb.close()

    def test_cell_write_number_value(self):
        """Test writing a number value to a cell."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "write", path, "Sheet", "A1", "42.5", "--type", "number"])

            assert result.exit_code == 0

            # Verify the value was written
            wb = openpyxl.load_workbook(path)
            assert wb.active["A1"].value == 42.5
            wb.close()

    def test_cell_write_boolean_true_value(self):
        """Test writing a boolean TRUE value to a cell."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "write", path, "Sheet", "A1", "TRUE", "--type", "bool"])

            assert result.exit_code == 0

            # Verify the value was written
            wb = openpyxl.load_workbook(path)
            assert wb.active["A1"].value is True
            wb.close()

    def test_cell_write_boolean_false_value(self):
        """Test writing a boolean FALSE value to a cell."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "write", path, "Sheet", "A1", "FALSE", "--type", "bool"])

            assert result.exit_code == 0

            # Verify the value was written
            wb = openpyxl.load_workbook(path)
            assert wb.active["A1"].value is False
            wb.close()

    def test_cell_write_invalid_type(self):
        """Test writing with an invalid type returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "write", path, "Sheet", "A1", "test", "--type", "invalid"])

            assert result.exit_code == ErrorCode.TYPE_COERCION_FAILED
            assert "Invalid type" in result.output

    def test_cell_write_date_value(self):
        """Test writing a date value to a cell."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "write", path, "Sheet", "A1", "2024-01-15", "--type", "date"])

            assert result.exit_code == 0

            # Verify the value was written (as a datetime)
            wb = openpyxl.load_workbook(path)
            cell_value = wb.active["A1"].value
            assert cell_value is not None
            assert hasattr(cell_value, "year")  # datetime has year attribute
            wb.close()

    def test_cell_write_preserves_leading_zeros_with_string_type(self):
        """Test writing a string that looks like a number preserves leading zeros."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["cell", "write", path, "Sheet", "A1", "00123", "--type", "string"])

            assert result.exit_code == 0

            # Verify the value was written as string preserving leading zeros
            wb = openpyxl.load_workbook(path)
            assert wb.active["A1"].value == "00123"
            wb.close()

