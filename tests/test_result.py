"""Tests for Result[T, E] and Maybe[T] types."""

import pytest

from xlforge.result import Ok, Err, Result, Some, Nothing, Maybe
from xlforge.result import is_ok, is_err, is_some, is_nothing


# =============================================================================
# Result[T, E] - Ok Tests
# =============================================================================


class TestResultOk:
    def test_is_ok_returns_true(self):
        result: Result[int, str] = Ok(42)
        assert result.is_ok() is True
        assert result.is_err() is False

    def test_unwrap_returns_value(self):
        result = Ok(42)
        assert result.unwrap() == 42

    def test_unwrap_or_returns_value(self):
        result = Ok(42)
        assert result.unwrap_or(0) == 42

    def test_unwrap_err_raises(self):
        result = Ok(42)
        with pytest.raises(ValueError, match="Called unwrap_err on Ok"):
            result.unwrap_err()

    def test_map_transforms_value(self):
        result = Ok(42)
        mapped = result.map(lambda x: x * 2)
        assert mapped == Ok(84)

    def test_map_preserves_err_on_err(self):
        result = Err("oops")
        mapped = result.map(lambda x: x * 2)
        assert mapped == Err("oops")

    def test_map_err_does_nothing_on_ok(self):
        result = Ok(42)
        mapped = result.map_err(lambda e: f"Error: {e}")
        assert mapped == Ok(42)

    def test_map_err_transforms_error(self):
        result = Err("oops")
        mapped = result.map_err(lambda e: f"Error: {e}")
        assert mapped == Err("Error: oops")

    def test_and_then_chains_ok(self):
        def parse_and_double(x: int) -> Result[int, str]:
            if x > 0:
                return Ok(x * 2)
            return Err("negative")

        result = Ok(21)
        chained = result.and_then(parse_and_double)
        assert chained == Ok(42)

    def test_and_then_preserves_err(self):
        def parse_and_double(x: int) -> Result[int, str]:
            if x > 0:
                return Ok(x * 2)
            return Err("negative")

        result = Err("error")
        chained = result.and_then(parse_and_double)
        assert chained == Err("error")

    def test_or_else_preserves_ok(self):
        def fallback(msg: str) -> Result[int, str]:
            return Ok(-1)

        result = Ok(42)
        chained = result.or_else(fallback)
        assert chained == Ok(42)

    def test_repr(self):
        result = Ok(42)
        assert repr(result) == "Ok(42)"

    def test_ok_with_none_value(self):
        result: Result[None, str] = Ok(None)
        assert result.is_ok()
        assert result.unwrap() is None

    def test_ok_with_complex_value(self):
        data = {"key": "value", "list": [1, 2, 3]}
        result = Ok(data)
        assert result.unwrap() == data


# =============================================================================
# Result[T, E] - Err Tests
# =============================================================================


class TestResultErr:
    def test_is_err_returns_true(self):
        result: Result[int, str] = Err("oops")
        assert result.is_err() is True
        assert result.is_ok() is False

    def test_unwrap_raises(self):
        result = Err("oops")
        with pytest.raises(ValueError, match="Called unwrap on Err"):
            result.unwrap()

    def test_unwrap_or_returns_default(self):
        result = Err("oops")
        assert result.unwrap_or(42) == 42

    def test_unwrap_err_returns_error(self):
        result = Err("oops")
        assert result.unwrap_err() == "oops"

    def test_map_does_nothing_on_err(self):
        result = Err("oops")
        mapped = result.map(lambda x: x * 2)
        assert mapped == Err("oops")

    def test_map_err_does_nothing_on_err(self):
        result = Err("oops")
        mapped = result.map_err(lambda e: f"Error: {e}")
        assert mapped == Err("Error: oops")

    def test_and_then_does_not_call_fn(self):
        def parse(x: int) -> Result[str, str]:
            return Ok(f"value:{x}")

        result = Err("error")
        chained = result.and_then(parse)
        assert chained == Err("error")

    def test_or_else_chains_on_err(self):
        def fallback(msg: str) -> Result[int, str]:
            return Ok(-1)

        result = Err("error")
        chained = result.or_else(fallback)
        assert chained == Ok(-1)

    def test_repr(self):
        result = Err("oops")
        assert repr(result) == "Err('oops')"

    def test_err_with_int_error(self):
        result: Result[int, int] = Err(404)
        assert result.is_err()
        assert result.unwrap_err() == 404

    def test_err_with_complex_error(self):
        error = {"code": 404, "message": "Not found"}
        result: Result[int, dict] = Err(error)
        assert result.is_err()
        assert result.unwrap_err() == error


# =============================================================================
# Result Type Guards
# =============================================================================


class TestResultTypeGuards:
    def test_is_ok_true_for_ok(self):
        result: Result[int, str] = Ok(42)
        assert is_ok(result) is True

    def test_is_ok_false_for_err(self):
        result: Result[int, str] = Err("oops")
        assert is_ok(result) is False

    def test_is_err_true_for_err(self):
        result: Result[int, str] = Err("oops")
        assert is_err(result) is True

    def test_is_err_false_for_ok(self):
        result: Result[int, str] = Ok(42)
        assert is_err(result) is False


# =============================================================================
# Maybe[T] - Some Tests
# =============================================================================


class TestMaybeSome:
    def test_is_some_returns_true(self):
        maybe: Maybe[int] = Some(42)
        assert maybe.is_some() is True
        assert maybe.is_nothing() is False

    def test_unwrap_returns_value(self):
        maybe = Some(42)
        assert maybe.unwrap() == 42

    def test_unwrap_or_returns_value(self):
        maybe = Some(42)
        assert maybe.unwrap_or(0) == 42

    def test_unwrap_none_raises(self):
        maybe = Some(42)
        with pytest.raises(ValueError, match="Called unwrap_none on Some"):
            maybe.unwrap_none()

    def test_map_transforms_value(self):
        maybe = Some(42)
        mapped = maybe.map(lambda x: x * 2)
        assert mapped == Some(84)

    def test_map_preserves_nothing(self):
        maybe = Nothing()
        mapped = maybe.map(lambda x: x * 2)
        assert mapped == Nothing()

    def test_filter_with_matching_predicate(self):
        maybe = Some(42)
        filtered = maybe.filter(lambda x: x > 10)
        assert filtered == Some(42)

    def test_filter_with_non_matching_predicate(self):
        maybe = Some(5)
        filtered = maybe.filter(lambda x: x > 10)
        assert filtered == Nothing()

    def test_and_then_chains_some(self):
        def get_positive(x: int) -> Maybe[int]:
            if x > 0:
                return Some(x * 2)
            return Nothing()

        maybe = Some(21)
        chained = maybe.and_then(get_positive)
        assert chained == Some(42)

    def test_and_then_preserves_nothing(self):
        def get_positive(x: int) -> Maybe[int]:
            if x > 0:
                return Some(x * 2)
            return Nothing()

        maybe = Nothing()
        chained = maybe.and_then(get_positive)
        assert chained == Nothing()

    def test_or_else_preserves_some(self):
        def fallback() -> Maybe[int]:
            return Some(-1)

        maybe = Some(42)
        result = maybe.or_else(fallback)
        assert result == Some(42)

    def test_repr(self):
        maybe = Some(42)
        assert repr(maybe) == "Some(42)"

    def test_some_with_none_value(self):
        maybe: Maybe[None] = Some(None)
        assert maybe.is_some()
        assert maybe.unwrap() is None

    def test_some_with_complex_value(self):
        data = {"key": "value"}
        maybe = Some(data)
        assert maybe.unwrap() == data


# =============================================================================
# Maybe[T] - Nothing Tests
# =============================================================================


class TestMaybeNothing:
    def test_is_some_returns_false(self):
        maybe: Maybe[int] = Nothing()
        assert maybe.is_some() is False
        assert maybe.is_nothing() is True

    def test_is_nothing_returns_true(self):
        maybe: Maybe[int] = Nothing()
        assert maybe.is_nothing() is True

    def test_unwrap_raises(self):
        maybe = Nothing()
        with pytest.raises(ValueError, match="Called unwrap on Nothing"):
            maybe.unwrap()

    def test_unwrap_or_returns_default(self):
        maybe = Nothing()
        assert maybe.unwrap_or(42) == 42

    def test_unwrap_none_returns_nothing(self):
        maybe: Maybe[int] = Nothing()
        assert maybe.unwrap_none() == maybe

    def test_map_does_nothing(self):
        maybe = Nothing()
        mapped = maybe.map(lambda x: x * 2)
        assert mapped == Nothing()

    def test_filter_does_nothing(self):
        maybe = Nothing()
        filtered = maybe.filter(lambda x: x > 10)
        assert filtered == Nothing()

    def test_and_then_does_not_call_fn(self):
        def get_positive(x: int) -> Maybe[int]:
            if x > 0:
                return Some(x * 2)
            return Nothing()

        maybe = Nothing()
        chained = maybe.and_then(get_positive)
        assert chained == Nothing()

    def test_or_else_calls_fallback(self):
        def fallback() -> Maybe[int]:
            return Some(-1)

        maybe = Nothing()
        result = maybe.or_else(fallback)
        assert result == Some(-1)

    def test_repr(self):
        maybe = Nothing()
        assert repr(maybe) == "Nothing"


# =============================================================================
# Maybe Type Guards
# =============================================================================


class TestMaybeTypeGuards:
    def test_is_some_true_for_some(self):
        maybe: Maybe[int] = Some(42)
        assert is_some(maybe) is True

    def test_is_some_false_for_nothing(self):
        maybe: Maybe[int] = Nothing()
        assert is_some(maybe) is False

    def test_is_nothing_true_for_nothing(self):
        maybe: Maybe[int] = Nothing()
        assert is_nothing(maybe) is True

    def test_is_nothing_false_for_some(self):
        maybe: Maybe[int] = Some(42)
        assert is_nothing(maybe) is False


# =============================================================================
# Edge Cases
# =============================================================================


class TestResultEdgeCases:
    def test_chaining_multiple_maps(self):
        result = Ok(2).map(lambda x: x + 1).map(lambda x: x * 2)
        assert result == Ok(6)

    def test_chaining_multiple_and_then(self):
        def add_one(x: int) -> Result[int, str]:
            return Ok(x + 1)

        def multiply_two(x: int) -> Result[int, str]:
            return Ok(x * 2)

        result = Ok(2).and_then(add_one).and_then(multiply_two)
        assert result == Ok(6)

    def test_error_propagation_through_chain(self):
        def fail(x: int) -> Result[int, str]:
            return Err("failed")

        result = Ok(2).and_then(fail).map(lambda x: x * 2)
        assert result == Err("failed")


class TestMaybeEdgeCases:
    def test_chaining_multiple_maps(self):
        maybe = Some(2).map(lambda x: x + 1).map(lambda x: x * 2)
        assert maybe == Some(6)

    def test_chaining_multiple_and_then(self):
        def add_one(x: int) -> Maybe[int]:
            return Some(x + 1)

        def multiply_two(x: int) -> Maybe[int]:
            return Some(x * 2)

        maybe = Some(2).and_then(add_one).and_then(multiply_two)
        assert maybe == Some(6)

    def test_nothing_propagation_through_chain(self):
        def fail(x: int) -> Maybe[int]:
            return Nothing()

        result = Some(2).and_then(fail).map(lambda x: x * 2)
        assert result == Nothing()

    def test_filter_with_predicate_that_always_returns_true(self):
        maybe = Some(42)
        filtered = maybe.filter(lambda x: True)
        assert filtered == Some(42)

    def test_filter_with_predicate_that_always_returns_false(self):
        maybe = Some(42)
        filtered = maybe.filter(lambda x: False)
        assert filtered == Nothing()


# =============================================================================
# Generics with Different Types
# =============================================================================


class TestResultGenerics:
    def test_result_with_string_value(self):
        result: Result[str, int] = Ok("hello")
        assert result.unwrap() == "hello"

    def test_result_with_multiple_error_types(self):
        def to_int(s: str) -> Result[int, ValueError]:
            try:
                return Ok(int(s))
            except ValueError as e:
                return Err(e)

        assert to_int("42").unwrap() == 42
        assert is_err(to_int("not a number"))
