"""Tests for CellValue and related functions."""

from datetime import datetime

import pytest

from xlforge.core.types.cell_value import (
    CellValue,
    _coerce_string_to_type,
    _infer_type_from_string,
)
from xlforge.core.types.value_type import ValueType


class TestCellValueCreation:
    """Test basic CellValue creation."""

    def test_create_with_raw_and_type(self):
        cv = CellValue(raw="hello", type=ValueType.STRING)
        assert cv.raw == "hello"
        assert cv.type == ValueType.STRING


class TestCellValueFromPython:
    """Test CellValue.from_python() factory method."""

    def test_from_python_none_returns_empty(self):
        cv = CellValue.from_python(None)
        assert cv.raw is None
        assert cv.type == ValueType.EMPTY

    def test_from_python_bool_true(self):
        cv = CellValue.from_python(True)
        assert cv.raw is True
        assert cv.type == ValueType.BOOL

    def test_from_python_bool_false(self):
        cv = CellValue.from_python(False)
        assert cv.raw is False
        assert cv.type == ValueType.BOOL

    def test_from_python_int(self):
        cv = CellValue.from_python(42)
        assert cv.raw == 42
        assert cv.type == ValueType.NUMBER

    def test_from_python_float(self):
        cv = CellValue.from_python(3.14)
        assert cv.raw == 3.14
        assert cv.type == ValueType.NUMBER

    def test_from_python_negative_number(self):
        cv = CellValue.from_python(-100)
        assert cv.raw == -100
        assert cv.type == ValueType.NUMBER

    def test_from_python_datetime(self):
        dt = datetime(2024, 1, 15, 10, 30, 0)
        cv = CellValue.from_python(dt)
        assert cv.raw == dt
        assert cv.type == ValueType.DATE

    def test_from_python_string(self):
        cv = CellValue.from_python("hello")
        assert cv.raw == "hello"
        assert cv.type == ValueType.STRING

    def test_from_python_string_formula(self):
        cv = CellValue.from_python("=A1+B1")
        assert cv.raw == "=A1+B1"
        assert cv.type == ValueType.FORMULA

    def test_from_python_other_type_falls_back_to_string(self):
        class CustomType:
            pass

        obj = CustomType()
        cv = CellValue.from_python(obj)
        assert cv.type == ValueType.STRING


class TestCellValueFromString:
    """Test CellValue.from_string() factory method."""

    def test_from_string_empty_returns_empty(self):
        cv = CellValue.from_string("")
        assert cv.type == ValueType.EMPTY
        assert cv.raw is None

    def test_from_string_none_returns_empty(self):
        cv = CellValue.from_string(None)  # type: ignore
        assert cv.type == ValueType.EMPTY
        assert cv.raw is None

    def test_from_string_number_inferred(self):
        cv = CellValue.from_string("42")
        assert cv.type == ValueType.NUMBER
        assert cv.raw == 42.0

    def test_from_string_negative_number(self):
        cv = CellValue.from_string("-123.45")
        assert cv.type == ValueType.NUMBER
        assert cv.raw == -123.45

    def test_from_string_bool_true_inferred(self):
        cv = CellValue.from_string("TRUE")
        assert cv.type == ValueType.BOOL
        assert cv.raw is True

    def test_from_string_bool_false_inferred(self):
        cv = CellValue.from_string("FALSE")
        assert cv.type == ValueType.BOOL
        assert cv.raw is False

    def test_from_string_bool_lowercase(self):
        cv = CellValue.from_string("true")
        assert cv.type == ValueType.BOOL
        assert cv.raw is True

    def test_from_string_formula(self):
        cv = CellValue.from_string("=SUM(A1:A10)")
        assert cv.type == ValueType.FORMULA
        assert cv.raw == "=SUM(A1:A10)"

    def test_from_string_date_iso_format(self):
        cv = CellValue.from_string("2024-01-15T10:30:00")
        assert cv.type == ValueType.DATE
        assert isinstance(cv.raw, datetime)

    def test_from_string_string_inferred(self):
        cv = CellValue.from_string("hello world")
        assert cv.type == ValueType.STRING
        assert cv.raw == "hello world"

    def test_from_string_with_type_hint_string(self):
        cv = CellValue.from_string("007", type_hint=ValueType.STRING)
        assert cv.type == ValueType.STRING
        assert cv.raw == "007"

    def test_from_string_with_type_hint_number(self):
        cv = CellValue.from_string("123.45", type_hint=ValueType.NUMBER)
        assert cv.type == ValueType.NUMBER
        assert cv.raw == 123.45

    def test_from_string_with_type_hint_number_comma_decimal(self):
        cv = CellValue.from_string("123,45", type_hint=ValueType.NUMBER)
        assert cv.type == ValueType.NUMBER
        assert cv.raw == 123.45

    def test_from_string_with_type_hint_bool_true(self):
        cv = CellValue.from_string("TRUE", type_hint=ValueType.BOOL)
        assert cv.type == ValueType.BOOL
        assert cv.raw is True

    def test_from_string_with_type_hint_bool_yes(self):
        cv = CellValue.from_string("YES", type_hint=ValueType.BOOL)
        assert cv.type == ValueType.BOOL
        assert cv.raw is True

    def test_from_string_with_type_hint_bool_1(self):
        cv = CellValue.from_string("1", type_hint=ValueType.BOOL)
        assert cv.type == ValueType.BOOL
        assert cv.raw is True

    def test_from_string_with_type_hint_date(self):
        cv = CellValue.from_string("2024-01-15T10:30:00", type_hint=ValueType.DATE)
        assert cv.type == ValueType.DATE
        assert isinstance(cv.raw, datetime)

    def test_from_string_with_type_hint_formula(self):
        cv = CellValue.from_string("A1+B1", type_hint=ValueType.FORMULA)
        assert cv.type == ValueType.FORMULA
        assert cv.raw == "=A1+B1"

    def test_from_string_formula_preserved_with_hint(self):
        cv = CellValue.from_string("=A1+B1", type_hint=ValueType.FORMULA)
        assert cv.raw == "=A1+B1"

    def test_from_string_with_type_hint_empty(self):
        cv = CellValue.from_string("", type_hint=ValueType.EMPTY)
        assert cv.type == ValueType.EMPTY
        assert cv.raw is None


class TestCellValueTypeConversions:
    """Test CellValue type conversion methods."""

    def test_as_string_from_string(self):
        cv = CellValue(raw="hello", type=ValueType.STRING)
        assert cv.as_string() == "hello"

    def test_as_string_from_number(self):
        cv = CellValue(raw=42, type=ValueType.NUMBER)
        assert cv.as_string() == "42"

    def test_as_string_from_bool_true(self):
        cv = CellValue(raw=True, type=ValueType.BOOL)
        assert cv.as_string() == "True"

    def test_as_string_from_empty(self):
        cv = CellValue(raw=None, type=ValueType.EMPTY)
        assert cv.as_string() == ""

    def test_as_number_from_number(self):
        cv = CellValue(raw=42.5, type=ValueType.NUMBER)
        assert cv.as_number() == 42.5

    def test_as_number_from_bool_true(self):
        cv = CellValue(raw=True, type=ValueType.BOOL)
        assert cv.as_number() == 1.0

    def test_as_number_from_bool_false(self):
        cv = CellValue(raw=False, type=ValueType.BOOL)
        assert cv.as_number() == 0.0

    def test_as_number_from_string_raises(self):
        cv = CellValue(raw="hello", type=ValueType.STRING)
        with pytest.raises(TypeError, match="Cannot convert ValueType.STRING to number"):
            cv.as_number()

    def test_as_number_from_empty_raises(self):
        cv = CellValue(raw=None, type=ValueType.EMPTY)
        with pytest.raises(TypeError, match="Cannot convert ValueType.EMPTY to number"):
            cv.as_number()

    def test_as_bool_from_bool_true(self):
        cv = CellValue(raw=True, type=ValueType.BOOL)
        assert cv.as_bool() is True

    def test_as_bool_from_bool_false(self):
        cv = CellValue(raw=False, type=ValueType.BOOL)
        assert cv.as_bool() is False

    def test_as_bool_from_number_raises(self):
        cv = CellValue(raw=42, type=ValueType.NUMBER)
        with pytest.raises(TypeError, match="Cannot convert ValueType.NUMBER to bool"):
            cv.as_bool()

    def test_as_bool_from_string_raises(self):
        cv = CellValue(raw="true", type=ValueType.STRING)
        with pytest.raises(TypeError, match="Cannot convert ValueType.STRING to bool"):
            cv.as_bool()

    def test_as_date_from_date(self):
        dt = datetime(2024, 1, 15, 10, 30, 0)
        cv = CellValue(raw=dt, type=ValueType.DATE)
        assert cv.as_date() == dt

    def test_as_date_from_string_raises(self):
        cv = CellValue(raw="2024-01-15", type=ValueType.STRING)
        with pytest.raises(TypeError, match="Cannot convert ValueType.STRING to date"):
            cv.as_date()


class TestCellValueQueries:
    """Test CellValue query methods."""

    def test_is_empty_true_for_empty(self):
        cv = CellValue(raw=None, type=ValueType.EMPTY)
        assert cv.is_empty() is True

    def test_is_empty_false_for_string(self):
        cv = CellValue(raw="hello", type=ValueType.STRING)
        assert cv.is_empty() is False

    def test_is_empty_false_for_number(self):
        cv = CellValue(raw=0, type=ValueType.NUMBER)
        assert cv.is_empty() is False

    def test_is_error_true_for_error(self):
        cv = CellValue(raw="#REF!", type=ValueType.ERROR)
        assert cv.is_error() is True

    def test_is_error_false_for_string(self):
        cv = CellValue(raw="hello", type=ValueType.STRING)
        assert cv.is_error() is False

    def test_is_error_false_for_number(self):
        cv = CellValue(raw=42, type=ValueType.NUMBER)
        assert cv.is_error() is False


class TestCellValueRepr:
    """Test CellValue string representation."""

    def test_repr_string(self):
        cv = CellValue(raw="hello", type=ValueType.STRING)
        assert repr(cv) == "CellValue('hello', ValueType.STRING)"

    def test_repr_number(self):
        cv = CellValue(raw=42, type=ValueType.NUMBER)
        assert repr(cv) == "CellValue(42, ValueType.NUMBER)"

    def test_repr_bool(self):
        cv = CellValue(raw=True, type=ValueType.BOOL)
        assert repr(cv) == "CellValue(True, ValueType.BOOL)"

    def test_repr_empty(self):
        cv = CellValue(raw=None, type=ValueType.EMPTY)
        assert repr(cv) == "CellValue(None, ValueType.EMPTY)"


class TestValueType:
    """Test ValueType enum."""

    def test_value_type_string_exists(self):
        assert ValueType.STRING.value == "string"

    def test_value_type_number_exists(self):
        assert ValueType.NUMBER.value == "number"

    def test_value_type_bool_exists(self):
        assert ValueType.BOOL.value == "bool"

    def test_value_type_date_exists(self):
        assert ValueType.DATE.value == "date"

    def test_value_type_formula_exists(self):
        assert ValueType.FORMULA.value == "formula"

    def test_value_type_empty_exists(self):
        assert ValueType.EMPTY.value == "empty"

    def test_value_type_error_exists(self):
        assert ValueType.ERROR.value == "error"

    def test_value_type_repr(self):
        assert repr(ValueType.STRING) == "ValueType.STRING"


class TestInferTypeFromString:
    """Test type inference from strings."""

    def test_infer_formula(self):
        result = _infer_type_from_string("=SUM(A1:A10)")
        assert result == ValueType.FORMULA

    def test_infer_bool_true_uppercase(self):
        result = _infer_type_from_string("TRUE")
        assert result == ValueType.BOOL

    def test_infer_bool_true_lowercase(self):
        result = _infer_type_from_string("true")
        assert result == ValueType.BOOL

    def test_infer_bool_false(self):
        result = _infer_type_from_string("FALSE")
        assert result == ValueType.BOOL

    def test_infer_number_integer(self):
        result = _infer_type_from_string("42")
        assert result == ValueType.NUMBER

    def test_infer_number_negative(self):
        result = _infer_type_from_string("-123.45")
        assert result == ValueType.NUMBER

    def test_infer_number_decimal(self):
        result = _infer_type_from_string("3.14")
        assert result == ValueType.NUMBER

    def test_infer_date_iso(self):
        result = _infer_type_from_string("2024-01-15T10:30:00")
        assert result == ValueType.DATE

    def test_infer_date_with_z(self):
        result = _infer_type_from_string("2024-01-15T10:30:00Z")
        assert result == ValueType.DATE

    def test_infer_string_plain_text(self):
        result = _infer_type_from_string("hello world")
        assert result == ValueType.STRING

    def test_infer_string_with_numbers(self):
        result = _infer_type_from_string("abc123def")
        assert result == ValueType.STRING

    def test_infer_string_invalid_date_format(self):
        result = _infer_type_from_string("not a date")
        assert result == ValueType.STRING


class TestCoerceStringToType:
    """Test string coercion to specific types."""

    def test_coerce_to_string(self):
        result = _coerce_string_to_type("hello", ValueType.STRING)
        assert result == "hello"

    def test_coerce_to_number(self):
        result = _coerce_string_to_type("123.45", ValueType.NUMBER)
        assert result == 123.45

    def test_coerce_to_number_with_comma(self):
        result = _coerce_string_to_type("123,45", ValueType.NUMBER)
        assert result == 123.45

    def test_coerce_to_bool_true_uppercase(self):
        result = _coerce_string_to_type("TRUE", ValueType.BOOL)
        assert result is True

    def test_coerce_to_bool_false_uppercase(self):
        result = _coerce_string_to_type("FALSE", ValueType.BOOL)
        assert result is False

    def test_coerce_to_bool_1(self):
        result = _coerce_string_to_type("1", ValueType.BOOL)
        assert result is True

    def test_coerce_to_bool_yes(self):
        result = _coerce_string_to_type("YES", ValueType.BOOL)
        assert result is True

    def test_coerce_to_date(self):
        result = _coerce_string_to_type("2024-01-15T10:30:00", ValueType.DATE)
        assert isinstance(result, datetime)

    def test_coerce_to_formula_without_equal(self):
        result = _coerce_string_to_type("A1+B1", ValueType.FORMULA)
        assert result == "=A1+B1"

    def test_coerce_to_formula_with_equal(self):
        result = _coerce_string_to_type("=A1+B1", ValueType.FORMULA)
        assert result == "=A1+B1"

    def test_coerce_to_empty(self):
        result = _coerce_string_to_type("anything", ValueType.EMPTY)
        assert result is None

    def test_coerce_unknown_type_returns_original(self):
        result = _coerce_string_to_type("hello", ValueType.ERROR)
        assert result == "hello"
