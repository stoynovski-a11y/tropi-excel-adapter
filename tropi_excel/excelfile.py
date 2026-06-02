"""ExcelFile — high-level wrapper over the Graph Excel REST API.

Usage::

    from tropi_excel import open_workbook

    wb = open_workbook("https://...sharepoint.com/...")   # or logical path
    with wb.session():
        wb.set_range("Sheet1", "A1:B1", values=[["Hello", None]], formulas=[[None, "=1+1"]])
        data = wb.read_range("Sheet1", "A1:B1")
        print(data["values"])          # [["Hello", 2]]

        wb.add_worksheet("NewSheet")
        wb.add_table_rows("MyTable", [{"A": "val1", "B": "val2"}])
        wb.delete_worksheet("NewSheet")

Session management
------------------
``session()`` is a context manager that creates a Graph workbook session with
``persistChanges=true``.  The session-id is sent in the ``workbook-session-id``
header on every subsequent API call.  If a call returns 404 (session expired),
the session is automatically recreated once and the call is retried.

Locking
-------
A module-level ``threading.Lock`` is held for the duration of each write
operation (``set_range``, ``add_table_rows``, ``add_worksheet``,
``delete_worksheet``).  This serialises writes to the same workbook from
different threads in the same process, preventing interleaved Graph sessions.

Retry
-----
HTTP 429 and 503 honour ``Retry-After``; HTTP 504 is retried on
``createSession``; all transient errors retry up to ``MAX_RETRIES`` times with
exponential back-off.
"""
from __future__ import annotations

import datetime
import logging
import threading
import time
from contextlib import contextmanager
from decimal import Decimal
from typing import Any, Generator

import requests

from .errors import ExcelApiError, ExcelNotFoundError, ExcelThrottledError
from .resolve import ExcelFile

logger = logging.getLogger("tropi_excel")

GRAPH = "https://graph.microsoft.com/v1.0"

MAX_RETRIES = 3
BACKOFF_BASE = 1.0   # seconds
BACKOFF_CAP = 30.0   # seconds

# Module-level write locks: (drive_id, item_id) → Lock
_write_locks: dict[tuple[str, str], threading.Lock] = {}
_write_locks_guard = threading.Lock()


def _get_write_lock(ef: ExcelFile) -> threading.Lock:
    key = (ef.drive_id, ef.item_id)
    with _write_locks_guard:
        if key not in _write_locks:
            _write_locks[key] = threading.Lock()
        return _write_locks[key]


# ---------------------------------------------------------------------------
# ExcelSession — internal; ExcelFile uses it via the session() context manager
# ---------------------------------------------------------------------------

class _ExcelSession:
    """Holds a Graph workbook session for the duration of a ``with`` block."""

    def __init__(self, ef: ExcelFile, token_provider: Any) -> None:
        self._ef = ef
        self._token_provider = token_provider
        self._session_id: str | None = None

    # --- headers -------------------------------------------------------------

    def _base_headers(self) -> dict[str, str]:
        token = self._token_provider.get_token()
        h = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        if self._session_id:
            h["workbook-session-id"] = self._session_id
        return h

    # --- session lifecycle ---------------------------------------------------

    def open(self) -> None:
        self._session_id = self._create_session()

    def close(self) -> None:
        if not self._session_id:
            return
        try:
            requests.post(
                f"{self._ef.item_base_url}/workbook/closeSession",
                headers=self._base_headers(),
                timeout=15,
            )
        except Exception:
            pass  # best-effort close
        self._session_id = None

    def _create_session(self) -> str:
        """POST createSession, retry 504 up to MAX_RETRIES times."""
        for attempt in range(MAX_RETRIES + 1):
            r = requests.post(
                f"{self._ef.item_base_url}/workbook/createSession",
                headers=self._base_headers(),
                json={"persistChanges": True},
                timeout=30,
            )
            if r.ok:
                return r.json()["id"]
            if r.status_code == 504 and attempt < MAX_RETRIES:
                delay = min(BACKOFF_BASE * (2 ** attempt), BACKOFF_CAP)
                logger.warning(
                    "createSession 504, retry %d/%d in %.1fs",
                    attempt + 1, MAX_RETRIES, delay,
                )
                time.sleep(delay)
                continue
            _raise_for_status(r)
        # unreachable, but keeps mypy happy
        raise ExcelApiError(504, "createSession failed after retries")

    # --- request plumbing ----------------------------------------------------

    def request(
        self,
        method: str,
        url: str,
        *,
        json: Any = None,
        _retry_on_404: bool = True,
    ) -> requests.Response:
        """Make an authenticated Graph request, retrying 429/503/504.

        On 404 (session expired), recreate the session once and retry
        when ``_retry_on_404`` is True (default).
        """
        if not url.startswith("http"):
            url = GRAPH + url

        for attempt in range(MAX_RETRIES + 1):
            r = requests.request(
                method,
                url,
                headers=self._base_headers(),
                json=json,
                timeout=60,
            )

            if r.ok:
                return r

            # 404 with a live session → session likely expired, recreate once.
            if r.status_code == 404 and _retry_on_404 and self._session_id:
                logger.warning("Got 404 — assuming session expired, recreating.")
                self._session_id = None
                self._session_id = self._create_session()
                # Retry without the 404-recreation flag (one shot only).
                return self.request(method, url, json=json, _retry_on_404=False)

            # 429 / 503 → throttled.
            if r.status_code in (429, 503):
                if attempt >= MAX_RETRIES:
                    raise ExcelThrottledError(
                        f"Graph {r.status_code} after {MAX_RETRIES} retries",
                        retry_after=None,
                    )
                retry_after_str = r.headers.get("Retry-After")
                retry_after = float(retry_after_str) if retry_after_str else None
                delay = (
                    min(retry_after, BACKOFF_CAP)
                    if retry_after
                    else min(BACKOFF_BASE * (2 ** attempt), BACKOFF_CAP)
                )
                logger.warning(
                    "Graph %d throttled, retry %d/%d in %.1fs",
                    r.status_code, attempt + 1, MAX_RETRIES, delay,
                )
                time.sleep(delay)
                continue

            # Everything else → raise immediately.
            _raise_for_status(r)

        raise ExcelApiError(0, "request loop exhausted without response")  # unreachable


def _raise_for_status(resp: requests.Response) -> None:
    if resp.ok:
        return
    if resp.status_code == 404:
        raise ExcelNotFoundError(resp.text[:300])
    raise ExcelApiError(resp.status_code, resp.text)


# ---------------------------------------------------------------------------
# ExcelFileClient — the main public object returned by open_workbook()
# ---------------------------------------------------------------------------

class ExcelFileClient:
    """High-level workbook client.

    Obtain one via ``tropi_excel.open_workbook()``.  All workbook operations
    must be called inside a ``with wb.session():`` block.
    """

    def __init__(self, ef: ExcelFile, token_provider: Any) -> None:
        self._ef = ef
        self._token_provider = token_provider
        self._session: _ExcelSession | None = None
        self._write_lock = _get_write_lock(ef)
        # Cache: table_name → list of column headers (letter A, B, C… order)
        self._table_header_cache: dict[str, list[str]] = {}
        # Cache: table_name → 0-based worksheet index of the table's first column
        # (e.g. a table whose range starts at column C → 2).  Used to convert a
        # worksheet column letter into a table-relative column index.
        self._table_start_col_cache: dict[str, int] = {}

    # --- session context manager --------------------------------------------

    @contextmanager
    def session(self) -> Generator[None, None, None]:
        """Open a Graph workbook session for the duration of the block."""
        sess = _ExcelSession(self._ef, self._token_provider)
        sess.open()
        self._session = sess
        try:
            yield
        finally:
            sess.close()
            self._session = None

    def _sess(self) -> _ExcelSession:
        if self._session is None:
            raise RuntimeError(
                "No active session. Use 'with wb.session():' before calling workbook methods."
            )
        return self._session

    def _wb_url(self, suffix: str) -> str:
        """Build a workbook endpoint URL."""
        return f"{self._ef.item_base_url}/workbook{suffix}"

    # --- read_range ----------------------------------------------------------

    def read_range(self, sheet: str, address: str) -> dict:
        """Read a range and return the Graph JSON (values, text, formulas, etc.).

        Args:
            sheet:   Worksheet name (e.g. "Sheet1").
            address: A1-notation address (e.g. "A1:C5").

        Returns:
            The Graph range object dict with keys ``values``, ``text``,
            ``formulas``, ``numberFormat``, etc.
        """
        url = self._wb_url(f"/worksheets('{sheet}')/range(address='{address}')")
        r = self._sess().request("GET", url)
        return r.json()

    # --- set_range -----------------------------------------------------------

    def set_range(
        self,
        sheet: str,
        address: str,
        values: list[list] | None = None,
        formulas: list[list] | None = None,
        number_format: list[list] | None = None,
    ) -> dict:
        """Write values and/or formulas into a range.

        Mixed writes (e.g. some cells get a value, others a formula) work by
        passing None in the positions you do not want to override — Graph uses
        null as "leave the existing value unchanged" for the field you omit,
        but here we send explicit nulls in the 2-D array as the spike does.

        Args:
            sheet:         Worksheet name.
            address:       A1-notation address.
            values:        2-D array of cell values (null = leave).
            formulas:      2-D array of formula strings (null = leave).
            number_format: 2-D array of format strings (null = leave).

        Returns:
            The updated Graph range object.
        """
        payload: dict[str, Any] = {}
        if values is not None:
            payload["values"] = _to_excel_grid(values)
        if formulas is not None:
            payload["formulas"] = formulas
        if number_format is not None:
            payload["numberFormat"] = number_format

        if not payload:
            raise ValueError("set_range: at least one of values/formulas/number_format must be set.")

        url = self._wb_url(f"/worksheets('{sheet}')/range(address='{address}')")
        with self._write_lock:
            r = self._sess().request("PATCH", url, json=payload)
        return r.json()

    # --- set_font_color ------------------------------------------------------

    def set_font_color(self, sheet: str, address: str, color: str) -> dict:
        """Set the font color of a range (e.g. flag a cell red).

        Patches the range's ``format/font`` resource — independent of the
        cell's value/formula, so it can be called after writing the value.

        Args:
            sheet:   Worksheet name.
            address: A1-notation address (a single cell ``K9`` or a range).
            color:   Hex color string ``#RRGGBB`` (e.g. ``"#FF0000"`` for red).

        Returns:
            The updated Graph rangeFont object.
        """
        if not (
            isinstance(color, str)
            and len(color) == 7
            and color[0] == "#"
            and all(c in "0123456789abcdefABCDEF" for c in color[1:])
        ):
            raise ValueError(
                f"color must be a '#RRGGBB' hex string; got {color!r}."
            )
        url = self._wb_url(
            f"/worksheets('{sheet}')/range(address='{address}')/format/font"
        )
        with self._write_lock:
            r = self._sess().request("PATCH", url, json={"color": color})
        return r.json()

    # --- table_column --------------------------------------------------------

    def table_column(self, table: str, column_name_or_letter: str) -> list:
        """Return the data values of a table column (excluding the header row).

        Args:
            table:                Table name (as shown in Excel's Name Manager).
            column_name_or_letter: Either the column header text (case-insensitive)
                                   or an Excel column letter (A, B, …).

        Returns:
            A flat list of cell values.
        """
        headers = self._get_table_headers(table)

        # Resolve letter or name to a 0-based TABLE-RELATIVE column index.
        # NB: Graph's /columns/{key} matches by the column's internal *id*
        # (e.g. "3", "46"), which is NOT the positional index — so we must
        # address by position via columns/itemAt(index=N) instead.
        col_upper = column_name_or_letter.strip().upper()
        if len(col_upper) <= 3 and col_upper.isalpha():
            # Worksheet column letter → worksheet-absolute index, then make it
            # table-relative by subtracting the table's start column.
            idx = _col_letter_to_index(col_upper) - self._table_start_col(table)
            if idx < 0 or idx >= len(headers):
                raise ValueError(
                    f"Column letter {column_name_or_letter!r} is out of range "
                    f"for table {table!r} which has {len(headers)} columns."
                )
        else:
            # Column header text → its position within the table headers is
            # already the 0-based table-relative index.
            lower = column_name_or_letter.strip().lower()
            try:
                idx = next(i for i, h in enumerate(headers) if h.lower() == lower)
            except StopIteration:
                raise ValueError(
                    f"Column {column_name_or_letter!r} not found in table {table!r}. "
                    f"Available headers: {headers}"
                ) from None

        # dataBodyRange = the column's data area, excluding header and total rows.
        url = self._wb_url(
            f"/tables('{table}')/columns/itemAt(index={idx})/dataBodyRange"
        )
        r = self._sess().request("GET", url)
        data = r.json()
        rows = data.get("values", [])
        return [row[0] for row in rows if row]

    def _get_table_headers(self, table: str) -> list[str]:
        """Return the ordered list of column header strings for *table*."""
        if table in self._table_header_cache:
            return self._table_header_cache[table]

        url = self._wb_url(f"/tables('{table}')/headerRowRange")
        r = self._sess().request("GET", url)
        data = r.json()
        rows = data.get("values", [[]])
        headers: list[str] = [str(v) for v in (rows[0] if rows else [])]
        self._table_header_cache[table] = headers
        return headers

    def _table_start_col(self, table: str) -> int:
        """Return the 0-based worksheet index of *table*'s first column.

        A table whose range is ``Въвод!A8:BE229`` starts at column A → 0;
        one at ``Sheet!C2:F9`` → 2.  Used to translate a worksheet column
        letter into a table-relative column index for ``columns/itemAt``.
        """
        if table in self._table_start_col_cache:
            return self._table_start_col_cache[table]

        url = self._wb_url(f"/tables('{table}')/range?$select=address")
        r = self._sess().request("GET", url)
        address = r.json().get("address", "")  # e.g. "Въвод!A8:BE229"
        cell = address.split("!")[-1].split(":")[0]  # "A8"
        letters = "".join(ch for ch in cell if ch.isalpha())  # "A"
        start = _col_letter_to_index(letters) if letters else 0
        self._table_start_col_cache[table] = start
        return start

    # --- add_table_rows ------------------------------------------------------

    def add_table_rows(
        self,
        table: str,
        rows: list[dict[str, Any]],
        index: int | None = None,
    ) -> dict:
        """Append (or insert) rows into a table.

        Args:
            table:  Table name.
            rows:   Each dict maps COLUMN LETTER (e.g. "A", "B") to a value.
                    Columns not specified are filled with null (Excel keeps
                    formulas / auto-fills as appropriate).
            index:  0-based row index to insert before.  None → append.

        Returns:
            The Graph tableRow object.
        """
        headers = self._get_table_headers(table)
        n_cols = len(headers)
        start_col = self._table_start_col(table)

        def _row_to_array(row_dict: dict[str, Any]) -> list:
            # The rows/add body is TABLE-relative: position 0 = the table's
            # first column.  Convert each worksheet letter to a table-relative
            # index by subtracting the table's start column.
            arr: list[Any] = [None] * n_cols
            for letter, value in row_dict.items():
                idx = _col_letter_to_index(letter.strip().upper()) - start_col
                if 0 <= idx < n_cols:
                    arr[idx] = _to_excel_value(value)
            return arr

        values = [_row_to_array(r) for r in rows]
        body: dict[str, Any] = {"values": values}
        if index is not None:
            body["index"] = index

        url = self._wb_url(f"/tables('{table}')/rows/add")
        with self._write_lock:
            r = self._sess().request("POST", url, json=body)
        return r.json()

    # --- add_worksheet -------------------------------------------------------

    def add_worksheet(self, name: str) -> dict:
        """Add a new worksheet.

        Args:
            name: The new worksheet's name (must be unique within the workbook).

        Returns:
            The Graph worksheet object.
        """
        url = self._wb_url("/worksheets/add")
        with self._write_lock:
            r = self._sess().request("POST", url, json={"name": name})
        return r.json()

    # --- delete_worksheet ----------------------------------------------------

    def delete_worksheet(self, name: str) -> None:
        """Delete a worksheet by name.

        Args:
            name: The worksheet name to delete.
        """
        url = self._wb_url(f"/worksheets('{name}')")
        with self._write_lock:
            self._sess().request("DELETE", url)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Excel's day 0 is 1899-12-30 (the offset absorbs the fictional 1900-02-29 leap
# day for all dates from 1900-03-01 onward, which covers every real-world date).
_EXCEL_EPOCH = datetime.datetime(1899, 12, 30)


def _to_excel_value(v: Any) -> Any:
    """Coerce a Python value into something the Graph Excel API can JSON-encode.

    The workbook range/table-row endpoints take values as a JSON body, so
    ``datetime``/``date`` (not JSON-serialisable) and ``Decimal`` (rejected by
    requests' encoder) must be converted first.  Dates become Excel serial
    numbers so the cell's existing date number-format renders them correctly;
    Decimals become floats.  Everything else passes through unchanged.
    """
    if isinstance(v, datetime.datetime):
        delta = v - _EXCEL_EPOCH
        return delta.days + (delta.seconds + delta.microseconds / 1e6) / 86400.0
    if isinstance(v, datetime.date):  # plain date (datetime already handled above)
        return (datetime.datetime(v.year, v.month, v.day) - _EXCEL_EPOCH).days
    if isinstance(v, Decimal):
        return float(v)
    return v


def _to_excel_grid(grid: list[list] | None) -> list[list] | None:
    """Apply :func:`_to_excel_value` across a 2-D array (or return None)."""
    if grid is None:
        return None
    return [[_to_excel_value(cell) for cell in row] for row in grid]


def _col_letter_to_index(letter: str) -> int:
    """Convert an Excel column letter to a 0-based column index.

    Examples: A→0, B→1, Z→25, AA→26, AB→27.
    """
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1
