"""Exception hierarchy for tropi_excel.

Backend-specific HTTP errors are translated to these so callers can write
provider-agnostic error handling.
"""
from __future__ import annotations


class ExcelError(Exception):
    """Base class for all tropi_excel errors."""


class ExcelApiError(ExcelError):
    """Graph returned an unexpected HTTP error.

    Attributes
    ----------
    status:   The HTTP status code.
    body:     The raw response body (str).
    """

    def __init__(self, status: int, body: str) -> None:
        self.status = status
        self.body = body
        super().__init__(f"Graph Excel API error {status}: {body[:300]}")


class ExcelAuthError(ExcelError):
    """Token acquisition failed."""


class ExcelNotFoundError(ExcelError):
    """The workbook, worksheet, or table was not found (HTTP 404)."""


class ExcelThrottledError(ExcelError):
    """Graph returned 429 or 503 (rate-limited / service unavailable).

    Attributes
    ----------
    retry_after:  Seconds to wait before retrying, or None if not given.
    """

    def __init__(self, message: str, retry_after: float | None = None) -> None:
        self.retry_after = retry_after
        super().__init__(message)


class ExcelResolveError(ExcelError):
    """Could not resolve a path or URL to a (drive_id, item_id) pair."""
