"""
Custom exceptions for validation and parsing errors.
"""


class ValidationError(Exception):
    """Base class for validation errors."""

    pass


class TypeConflictError(ValidationError):
    """Raised when more than one data type field is present."""

    pass


class DeferResolutionError(ValidationError):
    """Raised when a deferred FormulaValue references a missing Parameter."""

    pass
