"""tropi_excel — high-level wrapper over the Microsoft Graph Excel REST API.

Allows services to read and write SharePoint workbooks by cell/range/table
without downloading the file — no openpyxl, no upload round-trip.

Quick start::

    from tropi_excel import open_workbook

    wb = open_workbook("https://...sharepoint.com/...")   # or logical path
    with wb.session():
        wb.set_range("Sheet1", "A1", values=[["hello"]])
        result = wb.read_range("Sheet1", "A1")
        print(result["values"])    # [["hello"]]

Auth
----
Set ``GRAPH_EXCEL_BROKER_URL`` + ``GRAPH_EXCEL_BROKER_API_KEY`` in production
to use the broker.  For local testing with the excel-spike device-code cache,
leave those unset; ``LocalTokenProvider`` is used automatically.
"""
from __future__ import annotations

from .auth import BrokerTokenProvider, LocalTokenProvider, TokenProvider, get_token_provider
from .errors import (
    ExcelApiError,
    ExcelAuthError,
    ExcelError,
    ExcelNotFoundError,
    ExcelResolveError,
    ExcelThrottledError,
)
from .excelfile import ExcelFileClient
from .resolve import ExcelFile, invalidate, resolve

__version__ = "0.3.0"


def open_workbook(
    path_or_url: str,
    token_provider: "TokenProvider | None" = None,
) -> ExcelFileClient:
    """Resolve *path_or_url* and return an ``ExcelFileClient``.

    Args:
        path_or_url:    A SharePoint sharing/web URL **or** a logical path
                        (e.g. ``/SiteRoot/Finance/2026/budget.xlsx``).
        token_provider: Optional explicit ``TokenProvider``.  If omitted,
                        ``get_token_provider()`` is called (picks Broker or
                        Local based on env vars).

    Returns:
        An ``ExcelFileClient`` ready for use inside ``with wb.session():``.
    """
    tp = token_provider or get_token_provider()
    ef = resolve(path_or_url, tp)
    return ExcelFileClient(ef, tp)


__all__ = [
    "open_workbook",
    "ExcelFileClient",
    "ExcelFile",
    "resolve",
    "invalidate",
    "TokenProvider",
    "BrokerTokenProvider",
    "LocalTokenProvider",
    "get_token_provider",
    "ExcelError",
    "ExcelApiError",
    "ExcelAuthError",
    "ExcelNotFoundError",
    "ExcelResolveError",
    "ExcelThrottledError",
    "__version__",
]
