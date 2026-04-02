"""Result[T, E] and Maybe[T] types for explicit error handling.

This module provides functional error handling primitives inspired by
Rust's Result and Option types.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Generic, TypeVar, Callable, Union

if TYPE_CHECKING:
    from typing import TypeGuard

T = TypeVar("T")
E = TypeVar("E")
U = TypeVar("U")
F = TypeVar("F")


# =============================================================================
# Result[T, E] - Represents either success (Ok) or failure (Err)
# =============================================================================


@dataclass(frozen=True, slots=True)
class Ok(Generic[T]):
    """Represents a successful result containing a value of type T."""

    value: T

    def is_ok(self) -> bool:
        return True

    def is_err(self) -> bool:
        return False

    def unwrap(self) -> T:
        """Return the contained value. Raises if Err."""
        return self.value

    def unwrap_or(self, default: T) -> T:
        """Return the contained value or a default."""
        return self.value

    def unwrap_err(self) -> E:
        """Raise ValueError - Ok does not contain an error."""
        raise ValueError(f"Called unwrap_err on Ok({self.value!r})")

    def map(self, fn: Callable[[T], U]) -> Result[U, E]:
        """Transform the contained value with a function."""
        return Ok(fn(self.value))

    def map_err(self, fn: Callable[[E], F]) -> Result[T, F]:
        """Apply function to error (does nothing for Ok)."""
        return Ok(self.value)

    def and_then(self, fn: Callable[[T], Result[U, E]]) -> Result[U, E]:
        """Chain a Result-returning function on the value."""
        return fn(self.value)

    def or_else(self, fn: Callable[[E], Result[T, U]]) -> Result[T, U]:
        """Apply function to error (does nothing for Ok)."""
        return Ok(self.value)

    def __repr__(self) -> str:
        return f"Ok({self.value!r})"


@dataclass(frozen=True, slots=True)
class Err(Generic[E]):
    """Represents a failed result containing an error of type E."""

    error: E

    def is_ok(self) -> bool:
        return False

    def is_err(self) -> bool:
        return True

    def unwrap(self) -> T:
        """Raise ValueError - Err does not contain a value."""
        raise ValueError(f"Called unwrap on Err({self.error!r})")

    def unwrap_or(self, default: T) -> T:
        """Return the default value."""
        return default

    def unwrap_err(self) -> E:
        """Return the contained error."""
        return self.error

    def map(self, fn: Callable[[T], U]) -> Result[U, E]:
        """Apply function to value (does nothing for Err)."""
        return Err(self.error)

    def map_err(self, fn: Callable[[E], F]) -> Result[T, F]:
        """Transform the contained error with a function."""
        return Err(fn(self.error))

    def and_then(self, fn: Callable[[T], Result[U, E]]) -> Result[U, E]:
        """Apply function to value (does nothing for Err)."""
        return Err(self.error)

    def or_else(self, fn: Callable[[E], Result[T, U]]) -> Result[T, U]:
        """Chain a Result-returning function on the error."""
        return fn(self.error)

    def __repr__(self) -> str:
        return f"Err({self.error!r})"


# Union type for Result
Result = Union[Ok[T], Err[E]]


# =============================================================================
# Type Guards
# =============================================================================


def is_ok(result: Result[T, E]) -> TypeGuard[Ok[T]]:
    """Type guard: returns True if result is Ok."""
    return result.is_ok()


def is_err(result: Result[T, E]) -> TypeGuard[Err[E]]:
    """Type guard: returns True if result is Err."""
    return result.is_err()


# =============================================================================
# Maybe[T] - Represents an optional value (Some) or absence (Nothing)
# =============================================================================


@dataclass(frozen=True, slots=True)
class Some(Generic[T]):
    """Represents a present value of type T."""

    value: T

    def is_some(self) -> bool:
        return True

    def is_nothing(self) -> bool:
        return False

    def unwrap(self) -> T:
        """Return the contained value. Raises if Nothing."""
        return self.value

    def unwrap_or(self, default: T) -> T:
        """Return the contained value or a default."""
        return self.value

    def unwrap_none(self) -> T:
        """Raise ValueError - Some contains a value."""
        raise ValueError(f"Called unwrap_none on Some({self.value!r})")

    def map(self, fn: Callable[[T], U]) -> Maybe[U]:
        """Transform the contained value with a function."""
        return Some(fn(self.value))

    def filter(self, predicate: Callable[[T], bool]) -> Maybe[T]:
        """Return Some if predicate matches, else Nothing."""
        if predicate(self.value):
            return Some(self.value)
        return Nothing()

    def and_then(self, fn: Callable[[T], Maybe[U]]) -> Maybe[U]:
        """Chain a Maybe-returning function on the value."""
        return fn(self.value)

    def or_else(self, fn: Callable[[], Maybe[T]]) -> Maybe[T]:
        """Apply function to absence (does nothing for Some)."""
        return Some(self.value)

    def __repr__(self) -> str:
        return f"Some({self.value!r})"


@dataclass(frozen=True, slots=True)
class Nothing(Generic[T]):
    """Represents an absent value."""

    def is_some(self) -> bool:
        return False

    def is_nothing(self) -> bool:
        return True

    def unwrap(self) -> T:
        """Raise ValueError - Nothing does not contain a value."""
        raise ValueError("Called unwrap on Nothing")

    def unwrap_or(self, default: T) -> T:
        """Return the default value."""
        return default

    def unwrap_none(self) -> T:
        """Return the Nothing instance (for compatibility)."""
        return self  # type: ignore[return-value]

    def map(self, fn: Callable[[T], U]) -> Maybe[U]:
        """Apply function to value (does nothing for Nothing)."""
        return Nothing()

    def filter(self, predicate: Callable[[T], bool]) -> Maybe[T]:
        """Return Nothing (predicate cannot match)."""
        return Nothing()

    def and_then(self, fn: Callable[[T], Maybe[U]]) -> Maybe[U]:
        """Apply function to value (does nothing for Nothing)."""
        return Nothing()

    def or_else(self, fn: Callable[[], Maybe[T]]) -> Maybe[T]:
        """Chain a Maybe-returning function to get a fallback."""
        return fn()

    def __repr__(self) -> str:
        return "Nothing"


# Union type for Maybe
Maybe = Union[Some[T], Nothing[T]]


# =============================================================================
# Type Guards
# =============================================================================


def is_some(maybe: Maybe[T]) -> TypeGuard[Some[T]]:
    """Type guard: returns True if maybe is Some."""
    return maybe.is_some()


def is_nothing(maybe: Maybe[T]) -> TypeGuard[Nothing[T]]:
    """Type guard: returns True if maybe is Nothing."""
    return maybe.is_nothing()
