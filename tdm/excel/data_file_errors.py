"""Errors specific to the main data workbook."""

from tdm.domain.errors import UserActionableError


class NoReservedColumnError(UserActionableError):
    """Raised when a required workbook column is missing."""

    code = "WORKBOOK_COLUMN_MISSING"
