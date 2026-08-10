"""Errors specific to the main data workbook."""


class NoReservedColumnError(Exception):
    """Raised when a required workbook column is missing."""
