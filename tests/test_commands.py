import csv
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


class TestRangeRead:
    """Tests for xlforge range read command."""

    def test_range_read_file_not_found(self):
        """Test range read with non-existent file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "nonexistent.xlsx")
            result = runner.invoke(app, ["range", "read", path, "Sheet1", "A1:C3"])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_range_read_sheet_not_found(self):
        """Test range read with non-existent sheet returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["range", "read", path, "NonexistentSheet", "A1:C3"])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "Sheet not found" in result.output

    def test_range_read_table_output(self):
        """Test reading a range with table output."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A1"] = "Name"
            ws["B1"] = "Age"
            ws["C1"] = "Active"
            ws["A2"] = "Alice"
            ws["B2"] = 30
            ws["C2"] = True
            ws["A3"] = "Bob"
            ws["B3"] = 25
            ws["C3"] = False
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["range", "read", path, "Sheet1", "A1:C3"])

            assert result.exit_code == 0
            assert "Range: A1:C3" in result.output
            assert "3 rows x 3 columns" in result.output
            assert "Name" in result.output
            assert "Alice" in result.output
            assert "30" in result.output

    def test_range_read_json_output(self):
        """Test reading a range with JSON output."""
        import json

        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A1"] = "Hello"
            ws["B1"] = 42
            ws["A2"] = "World"
            ws["B2"] = True
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["range", "read", path, "Sheet1", "A1:B2", "--json"])

            assert result.exit_code == 0
            data = json.loads(result.output)
            assert data == [["Hello", 42], ["World", True]]

    def test_range_read_single_cell(self):
        """Test reading a single cell as a range."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A1"] = "Single Cell"
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["range", "read", path, "Sheet1", "A1:A1"])

            assert result.exit_code == 0
            assert "Single Cell" in result.output
            assert "1 rows x 1 columns" in result.output

    def test_range_read_empty_range(self):
        """Test reading an empty range."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["range", "read", path, "Sheet1", "A1:B2"])

            assert result.exit_code == 0
            # Empty cells will show as empty range message
            assert "A1:B2 is empty" in result.output


class TestRangeWrite:
    """Tests for xlforge range write command."""

    def test_range_write_file_not_found(self):
        """Test range write with non-existent file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "nonexistent.xlsx")
            result = runner.invoke(app, ["range", "write", path, "Sheet1", "A1:C3", '[["a","b"],["c","d"]]'])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_range_write_sheet_not_found(self):
        """Test range write with non-existent sheet returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["range", "write", path, "NonexistentSheet", "A1:C3", '[["a","b"],["c","d"]]'])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "Sheet not found" in result.output

    def test_range_write_json_values(self):
        """Test writing values from JSON."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, [
                "range", "write", path, "Sheet", "A1:C3",
                '[["Name","Age","Active"],["Alice",30,true],["Bob",25,false]]'
            ])

            assert result.exit_code == 0
            assert "Written 3 row(s) x 3 column(s) to range A1:C3" in result.output

            # Verify the values were written
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            assert ws["A1"].value == "Name"
            assert ws["B1"].value == "Age"
            assert ws["C1"].value == "Active"
            assert ws["A2"].value == "Alice"
            assert ws["B2"].value == 30
            assert ws["C2"].value is True
            assert ws["A3"].value == "Bob"
            assert ws["B3"].value == 25
            assert ws["C3"].value is False
            wb.close()

    def test_range_write_csv_file(self):
        """Test writing values from CSV file."""
        import csv

        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            # Create CSV file
            csv_path = os.path.join(tmpdir, "values.csv")
            with open(csv_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Product", "Price", "Qty"])
                writer.writerow(["Apple", "1.50", "100"])
                writer.writerow(["Banana", "0.75", "200"])

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, [
                "range", "write", path, "Sheet", "A1:C3",
                "--csv", csv_path
            ])

            assert result.exit_code == 0
            assert "Written 3 row(s) x 3 column(s) to range A1:C3" in result.output

            # Verify the values were written
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            assert ws["A1"].value == "Product"
            assert ws["B1"].value == "Price"
            assert ws["C1"].value == "Qty"
            assert ws["A2"].value == "Apple"
            assert ws["B2"].value == "1.50"
            assert ws["C2"].value == "100"
            assert ws["A3"].value == "Banana"
            assert ws["B3"].value == "0.75"
            assert ws["C3"].value == "200"
            wb.close()

    def test_range_write_invalid_json(self):
        """Test writing with invalid JSON returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, [
                "range", "write", path, "Sheet", "A1:B2",
                "not valid json"
            ])

            assert result.exit_code == 1  # ErrorCode.INVALID_ARGUMENT
            assert "Invalid JSON" in result.output

    def test_range_write_csv_file_not_found(self):
        """Test writing with non-existent CSV file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            csv_path = os.path.join(tmpdir, "nonexistent.csv")
            result = runner.invoke(app, [
                "range", "write", path, "Sheet", "A1:B2",
                "--csv", csv_path
            ])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_range_write_missing_values(self):
        """Test writing without providing values returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, [
                "range", "write", path, "Sheet", "A1:B2"
            ])

            assert result.exit_code == 1  # ErrorCode.INVALID_ARGUMENT
            assert "Must provide either" in result.output

    def test_range_write_both_json_and_csv_error(self):
        """Test writing with both JSON and CSV returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            csv_path = os.path.join(tmpdir, "values.csv")
            with open(csv_path, "w", newline="", encoding="utf-8") as f:
                f.write("a,b\nc,d")

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, [
                "range", "write", path, "Sheet", "A1:B2",
                '[["a","b"],["c","d"]]', "--csv", csv_path
            ])

            assert result.exit_code == 1  # ErrorCode.INVALID_ARGUMENT
            assert "Cannot specify both" in result.output


class TestCsvImport:
    """Tests for xlforge csv import command."""

    def test_csv_import_file_not_found(self):
        """Test csv import with non-existent CSV file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            csv_path = os.path.join(tmpdir, "nonexistent.csv")
            xlsx_path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["csv", "import", csv_path, xlsx_path, "Sheet"])

            assert result.exit_code == ErrorCode.CSV_NOT_FOUND
            assert "does not exist" in result.output.lower()

    def test_csv_import_excel_file_not_found(self):
        """Test csv import with non-existent Excel file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = os.path.join(tmpdir, "test.csv")
            xlsx_path = os.path.join(tmpdir, "nonexistent.xlsx")

            # Create CSV file
            with open(csv_path, "w") as f:
                f.write("a,b,c\n")

            result = runner.invoke(app, ["csv", "import", csv_path, xlsx_path, "Sheet"])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_csv_import_sheet_not_found(self):
        """Test csv import with non-existent sheet returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            csv_path = os.path.join(tmpdir, "test.csv")
            xlsx_path = os.path.join(tmpdir, "test.xlsx")

            # Create CSV file
            with open(csv_path, "w") as f:
                f.write("a,b,c\n")

            result = runner.invoke(app, ["csv", "import", csv_path, xlsx_path, "NonExistent"])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "not found" in result.output.lower()

    def test_csv_import_basic(self):
        """Test basic CSV import."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            csv_path = os.path.join(tmpdir, "test.csv")
            xlsx_path = os.path.join(tmpdir, "test.xlsx")

            # Create CSV file
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Name", "Age", "City"])
                writer.writerow(["Alice", "30", "NYC"])
                writer.writerow(["Bob", "25", "LA"])

            result = runner.invoke(app, ["csv", "import", csv_path, xlsx_path, "Sheet"])

            assert result.exit_code == 0
            assert "Imported" in result.output

            # Verify data was imported (numbers are type-coerced)
            wb = openpyxl.load_workbook(xlsx_path)
            ws = wb.active
            assert ws["A1"].value == "Name"
            assert ws["B1"].value == "Age"
            assert ws["C1"].value == "City"
            assert ws["A2"].value == "Alice"
            assert ws["B2"].value == 30  # Number, not string
            assert ws["C2"].value == "NYC"
            assert ws["A3"].value == "Bob"
            assert ws["B3"].value == 25
            assert ws["C3"].value == "LA"
            wb.close()

    def test_csv_import_with_header(self):
        """Test CSV import with --has-header option."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            csv_path = os.path.join(tmpdir, "test.csv")
            xlsx_path = os.path.join(tmpdir, "test.xlsx")

            # Create CSV file
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Name", "Age", "City"])
                writer.writerow(["Alice", "30", "NYC"])
                writer.writerow(["Bob", "25", "LA"])

            result = runner.invoke(app, ["csv", "import", csv_path, xlsx_path, "Sheet", "--has-header"])

            assert result.exit_code == 0

            # Verify header row was skipped and data starts at A1 (numbers are type-coerced)
            wb = openpyxl.load_workbook(xlsx_path)
            ws = wb.active
            assert ws["A1"].value == "Alice"
            assert ws["B1"].value == 30  # Number, not string
            assert ws["A2"].value == "Bob"
            wb.close()

    def test_csv_import_empty_file(self):
        """Test CSV import with empty file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            csv_path = os.path.join(tmpdir, "empty.csv")
            xlsx_path = os.path.join(tmpdir, "test.xlsx")

            # Create empty CSV file
            with open(csv_path, "w") as f:
                pass

            result = runner.invoke(app, ["csv", "import", csv_path, xlsx_path, "Sheet"])

            assert result.exit_code == ErrorCode.INVALID_CSV_FORMAT
            assert "empty" in result.output.lower()


class TestCsvExport:
    """Tests for xlforge csv export command."""

    def test_csv_export_file_not_found(self):
        """Test csv export with non-existent Excel file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "nonexistent.xlsx")
            result = runner.invoke(app, ["csv", "export", path, "Sheet"])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_csv_export_sheet_not_found(self):
        """Test csv export with non-existent sheet returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["csv", "export", path, "NonExistent"])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "not found" in result.output.lower()

    def test_csv_export_basic(self):
        """Test basic CSV export to stdout."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A1"] = "Name"
            ws["B1"] = "Age"
            ws["C1"] = "City"
            ws["A2"] = "Alice"
            ws["B2"] = "30"
            ws["C2"] = "NYC"
            ws["A3"] = "Bob"
            ws["B3"] = "25"
            ws["C3"] = "LA"
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["csv", "export", path, "Sheet1"])

            assert result.exit_code == 0
            assert "Name" in result.output
            assert "Alice" in result.output
            assert "30" in result.output
            assert "Bob" in result.output

    def test_csv_export_to_file(self):
        """Test CSV export to output file."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A1"] = "Name"
            ws["B1"] = "Age"
            ws["A2"] = "Alice"
            ws["B2"] = "30"
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            xlsx_path = os.path.join(tmpdir, "test.xlsx")
            csv_path = os.path.join(tmpdir, "output.csv")
            result = runner.invoke(app, ["csv", "export", xlsx_path, "Sheet1", "--output", csv_path])

            assert result.exit_code == 0
            assert "Exported" in result.output

            # Verify CSV content
            with open(csv_path, "r") as f:
                content = f.read()
            assert "Name" in content
            assert "Alice" in content

    def test_csv_export_with_range(self):
        """Test CSV export with specified range."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A1"] = "Name"
            ws["B1"] = "Age"
            ws["C1"] = "City"
            ws["A2"] = "Alice"
            ws["B2"] = "30"
            ws["C2"] = "NYC"
            ws["A3"] = "Bob"
            ws["B3"] = "25"
            ws["C3"] = "LA"
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            xlsx_path = os.path.join(tmpdir, "test.xlsx")
            csv_path = os.path.join(tmpdir, "output.csv")
            result = runner.invoke(app, ["csv", "export", xlsx_path, "Sheet1", "--range", "A1:B2", "--output", csv_path])

            assert result.exit_code == 0

            # Verify CSV content (only A1:B2)
            with open(csv_path, "r") as f:
                reader = csv.reader(f)
                rows = list(reader)
            assert rows[0][0] == "Name"
            assert rows[0][1] == "Age"
            assert rows[1][0] == "Alice"
            assert rows[1][1] == "30"
            assert len(rows) == 2

    def test_csv_export_number_coercion(self):
        """Test CSV export properly handles number types."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A1"] = "Name"
            ws["B1"] = "Score"
            ws["A2"] = "Alice"
            ws["B2"] = 42.5
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            xlsx_path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["csv", "export", xlsx_path, "Sheet1"])

            assert result.exit_code == 0
            assert "42.5" in result.output


class TestRowHide:
    """Tests for xlforge row hide command."""

    def test_row_hide_file_not_found(self):
        """Test row hide with non-existent file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "nonexistent.xlsx")
            result = runner.invoke(app, ["row", "hide", path, "Sheet1", "1"])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_row_hide_sheet_not_found(self):
        """Test row hide with non-existent sheet returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["row", "hide", path, "NonexistentSheet", "1"])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "Sheet not found" in result.output

    def test_row_hide_invalid_row(self):
        """Test row hide with invalid row number returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["row", "hide", path, "Sheet", "0"])

            assert result.exit_code == ErrorCode.ROW_NOT_FOUND
            assert "Invalid row" in result.output

    def test_row_hide_success(self):
        """Test hiding a row."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A1"] = "Header"
            ws["A2"] = "Data"
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["row", "hide", path, "Sheet1", "1"])

            assert result.exit_code == 0
            assert "Hid row 1" in result.output

            # Verify row is hidden
            wb = openpyxl.load_workbook(path)
            assert wb["Sheet1"].row_dimensions[1].hidden is True
            wb.close()


class TestRowUnhide:
    """Tests for xlforge row unhide command."""

    def test_row_unhide_file_not_found(self):
        """Test row unhide with non-existent file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "nonexistent.xlsx")
            result = runner.invoke(app, ["row", "unhide", path, "Sheet1", "1"])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_row_unhide_sheet_not_found(self):
        """Test row unhide with non-existent sheet returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["row", "unhide", path, "NonexistentSheet", "1"])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "Sheet not found" in result.output

    def test_row_unhide_invalid_row(self):
        """Test row unhide with invalid row number returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["row", "unhide", path, "Sheet", "0"])

            assert result.exit_code == ErrorCode.ROW_NOT_FOUND
            assert "Invalid row" in result.output

    def test_row_unhide_success(self):
        """Test unhiding a row."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.row_dimensions[1].hidden = True
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["row", "unhide", path, "Sheet1", "1"])

            assert result.exit_code == 0
            assert "Unhid row 1" in result.output

            # Verify row is visible
            wb = openpyxl.load_workbook(path)
            assert wb["Sheet1"].row_dimensions[1].hidden is False
            wb.close()


class TestColumnHide:
    """Tests for xlforge column hide command."""

    def test_column_hide_file_not_found(self):
        """Test column hide with non-existent file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "nonexistent.xlsx")
            result = runner.invoke(app, ["column", "hide", path, "Sheet1", "A"])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_column_hide_sheet_not_found(self):
        """Test column hide with non-existent sheet returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["column", "hide", path, "NonexistentSheet", "A"])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "Sheet not found" in result.output

    def test_column_hide_success(self):
        """Test hiding a column."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A1"] = "Header1"
            ws["B1"] = "Header2"
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["column", "hide", path, "Sheet1", "A"])

            assert result.exit_code == 0
            assert "Hid column A" in result.output

            # Verify column is hidden
            wb = openpyxl.load_workbook(path)
            assert wb["Sheet1"].column_dimensions["A"].hidden is True
            wb.close()

    def test_column_hide_lowercase(self):
        """Test hiding a column with lowercase letter."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["column", "hide", path, "Sheet1", "b"])

            assert result.exit_code == 0
            assert "Hid column B" in result.output

            # Verify column is hidden
            wb = openpyxl.load_workbook(path)
            assert wb["Sheet1"].column_dimensions["B"].hidden is True
            wb.close()


class TestColumnUnhide:
    """Tests for xlforge column unhide command."""

    def test_column_unhide_file_not_found(self):
        """Test column unhide with non-existent file returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "nonexistent.xlsx")
            result = runner.invoke(app, ["column", "unhide", path, "Sheet1", "A"])

            assert result.exit_code == ErrorCode.FILE_DOES_NOT_EXIST
            assert "does not exist" in result.output.lower()

    def test_column_unhide_sheet_not_found(self):
        """Test column unhide with non-existent sheet returns error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["column", "unhide", path, "NonexistentSheet", "A"])

            assert result.exit_code == ErrorCode.SHEET_NOT_FOUND
            assert "Sheet not found" in result.output

    def test_column_unhide_success(self):
        """Test unhiding a column."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.column_dimensions["A"].hidden = True
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["column", "unhide", path, "Sheet1", "A"])

            assert result.exit_code == 0
            assert "Unhid column A" in result.output

            # Verify column is visible
            wb = openpyxl.load_workbook(path)
            assert wb["Sheet1"].column_dimensions["A"].hidden is False
            wb.close()

    def test_column_unhide_lowercase(self):
        """Test unhiding a column with lowercase letter."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.column_dimensions["B"].hidden = True
            wb.save(os.path.join(tmpdir, "test.xlsx"))

            path = os.path.join(tmpdir, "test.xlsx")
            result = runner.invoke(app, ["column", "unhide", path, "Sheet1", "b"])

            assert result.exit_code == 0
            assert "Unhid column B" in result.output

            # Verify column is visible
            wb = openpyxl.load_workbook(path)
            assert wb["Sheet1"].column_dimensions["B"].hidden is False
            wb.close()
