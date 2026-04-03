"""Tests for CellRef and related functions."""

import pytest

from xlforge.core.types.cell_ref import (
    CellRef,
    cell_ref_to_row_col,
    row_col_to_cell_ref,
    col_to_index,
    index_to_col,
)


class TestCellRefBasic:
    """Basic CellRef creation and properties."""

    def test_create_simple_cell_ref(self):
        ref = CellRef(sheet="", coord="A1")
        assert ref.sheet == ""
        assert ref.coord == "A1"

    def test_create_cell_ref_with_sheet(self):
        ref = CellRef(sheet="Sheet1", coord="B2")
        assert ref.sheet == "Sheet1"
        assert ref.coord == "B2"

    def test_empty_coord_raises_value_error(self):
        with pytest.raises(ValueError, match="Cell reference cannot be empty"):
            CellRef(sheet="", coord="")


class TestCellRefProperties:
    """Test CellRef properties."""

    def test_row_property_a1(self):
        ref = CellRef(sheet="", coord="A1")
        assert ref.row == 0

    def test_row_property_b2(self):
        ref = CellRef(sheet="", coord="B2")
        assert ref.row == 1

    def test_row_property_z100(self):
        ref = CellRef(sheet="", coord="Z100")
        assert ref.row == 99

    def test_col_property_a(self):
        ref = CellRef(sheet="", coord="A1")
        assert ref.col == 0

    def test_col_property_b(self):
        ref = CellRef(sheet="", coord="B1")
        assert ref.col == 1

    def test_col_property_z(self):
        ref = CellRef(sheet="", coord="Z1")
        assert ref.col == 25

    def test_col_property_aa(self):
        ref = CellRef(sheet="", coord="AA1")
        assert ref.col == 26

    def test_col_property_az(self):
        ref = CellRef(sheet="", coord="AZ1")
        assert ref.col == 51

    def test_col_property_ba(self):
        ref = CellRef(sheet="", coord="BA1")
        assert ref.col == 52

    def test_is_range_false_for_single_cell(self):
        ref = CellRef(sheet="", coord="A1")
        assert ref.is_range is False

    def test_is_range_true_for_range(self):
        ref = CellRef(sheet="", coord="A1:C3")
        assert ref.is_range is True

    def test_end_row_none_for_single_cell(self):
        ref = CellRef(sheet="", coord="A1")
        assert ref.end_row is None

    def test_end_row_for_range(self):
        ref = CellRef(sheet="", coord="A1:C10")
        assert ref.end_row == 9  # 10 - 1

    def test_end_col_none_for_single_cell(self):
        ref = CellRef(sheet="", coord="A1")
        assert ref.end_col is None

    def test_end_col_for_range(self):
        ref = CellRef(sheet="", coord="A1:C3")
        assert ref.end_col == 2  # C = 2


class TestCellRefMethods:
    """Test CellRef methods."""

    def test_to_a1_notation_simple(self):
        ref = CellRef(sheet="", coord="A1")
        assert ref.to_a1_notation() == "A1"

    def test_to_a1_notation_range(self):
        ref = CellRef(sheet="", coord="B2:D5")
        assert ref.to_a1_notation() == "B2:D5"


class TestCellRefStringRepresentation:
    """Test CellRef string representations."""

    def test_str_without_sheet(self):
        ref = CellRef(sheet="", coord="A1")
        assert str(ref) == "A1"

    def test_str_with_sheet(self):
        ref = CellRef(sheet="Sheet1", coord="A1")
        assert str(ref) == "Sheet1!A1"

    def test_repr_without_sheet(self):
        ref = CellRef(sheet="", coord="A1")
        assert repr(ref) == "CellRef(sheet='', coord='A1')"

    def test_repr_with_sheet(self):
        ref = CellRef(sheet="Data", coord="B2")
        assert repr(ref) == "CellRef(sheet='Data', coord='B2')"


class TestCellRefEquality:
    """Test CellRef equality and hashing."""

    def test_eq_true_for_same_values(self):
        ref1 = CellRef(sheet="", coord="A1")
        ref2 = CellRef(sheet="", coord="A1")
        assert ref1 == ref2

    def test_eq_false_for_different_coords(self):
        ref1 = CellRef(sheet="", coord="A1")
        ref2 = CellRef(sheet="", coord="A2")
        assert ref1 != ref2

    def test_eq_false_for_different_sheets(self):
        ref1 = CellRef(sheet="Sheet1", coord="A1")
        ref2 = CellRef(sheet="Sheet2", coord="A1")
        assert ref1 != ref2

    def test_hash_same_values(self):
        ref1 = CellRef(sheet="", coord="A1")
        ref2 = CellRef(sheet="", coord="A1")
        assert hash(ref1) == hash(ref2)

    def test_can_use_in_set(self):
        ref1 = CellRef(sheet="", coord="A1")
        ref2 = CellRef(sheet="", coord="A1")
        ref_set = {ref1, ref2}
        assert len(ref_set) == 1

    def test_can_use_as_dict_key(self):
        ref = CellRef(sheet="", coord="A1")
        d = {ref: "value"}
        assert d[ref] == "value"


class TestCellRefInvalidReferences:
    """Test that invalid cell references raise ValueError."""

    def test_invalid_coord_raises_on_row_access(self):
        ref = CellRef(sheet="", coord="INVALID")
        with pytest.raises(ValueError, match="Invalid cell reference"):
            _ = ref.row

    def test_invalid_coord_raises_on_col_access(self):
        ref = CellRef(sheet="", coord="INVALID")
        with pytest.raises(ValueError, match="Invalid cell reference"):
            _ = ref.col

    def test_empty_string_raises_on_construction(self):
        with pytest.raises(ValueError, match="Cell reference cannot be empty"):
            CellRef(sheet="", coord="")


class TestCellRefModuleFunctions:
    """Test module-level functions for cell reference manipulation."""

    def test_cell_ref_to_row_col_a1(self):
        row, col = cell_ref_to_row_col("A1")
        assert row == 0
        assert col == 0

    def test_cell_ref_to_row_col_b2(self):
        row, col = cell_ref_to_row_col("B2")
        assert row == 1
        assert col == 1

    def test_cell_ref_to_row_col_z100(self):
        row, col = cell_ref_to_row_col("Z100")
        assert row == 99
        assert col == 25

    def test_cell_ref_to_row_col_aa1(self):
        row, col = cell_ref_to_row_col("AA1")
        assert row == 0
        assert col == 26

    def test_cell_ref_to_row_col_invalid_raises(self):
        with pytest.raises(ValueError, match="Invalid cell reference"):
            cell_ref_to_row_col("INVALID")

    def test_row_col_to_cell_ref_a1(self):
        result = row_col_to_cell_ref(0, 0)
        assert result == "A1"

    def test_row_col_to_cell_ref_b2(self):
        result = row_col_to_cell_ref(1, 1)
        assert result == "B2"

    def test_row_col_to_cell_ref_z100(self):
        result = row_col_to_cell_ref(99, 25)
        assert result == "Z100"

    def test_row_col_to_cell_ref_aa1(self):
        result = row_col_to_cell_ref(0, 26)
        assert result == "AA1"


class TestColToIndex:
    """Test column letter to index conversion."""

    def test_single_letters(self):
        assert col_to_index("A") == 0
        assert col_to_index("B") == 1
        assert col_to_index("Z") == 25

    def test_double_letters(self):
        assert col_to_index("AA") == 26
        assert col_to_index("AZ") == 51
        assert col_to_index("BA") == 52

    def test_lowercase_converted_to_uppercase(self):
        assert col_to_index("a") == 0
        assert col_to_index("b") == 1
        assert col_to_index("z") == 25


class TestIndexToCol:
    """Test index to column letter conversion."""

    def test_single_digits(self):
        assert index_to_col(0) == "A"
        assert index_to_col(1) == "B"
        assert index_to_col(25) == "Z"

    def test_double_digits(self):
        assert index_to_col(26) == "AA"
        assert index_to_col(51) == "AZ"
        assert index_to_col(52) == "BA"


class TestRoundTrip:
    """Test round-trip conversions."""

    def test_row_col_roundtrip(self):
        original = "Z100"
        row, col = cell_ref_to_row_col(original)
        result = row_col_to_cell_ref(row, col)
        assert result == original

    def test_index_col_roundtrip(self):
        original = "ABC"
        idx = col_to_index(original)
        result = index_to_col(idx)
        assert result == original
