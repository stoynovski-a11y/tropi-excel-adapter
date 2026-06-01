# tropi-excel-adapter

High-level Python wrapper over the **Microsoft Graph Excel REST API** for editing SharePoint workbooks by cell, range, or table — without downloading, modifying locally, and re-uploading.

Mirrors `tropi-storage-adapter` in packaging and style, and optionally reuses its routing module to resolve logical paths via `M365_ROUTES`.

---

## Install

```bash
# From the local dev clone (recommended while not on PyPI)
pip install -e ~/dev/tropi-excel-adapter

# With logical-path routing (requires tropi-storage-adapter)
pip install -e ~/dev/tropi-excel-adapter -e ~/dev/tropi-storage-adapter
```

---

## Auth

### Production — Broker token provider

Set two env vars:

| Variable | Value |
|---|---|
| `GRAPH_EXCEL_BROKER_URL` | Base URL of the internal token broker |
| `GRAPH_EXCEL_BROKER_API_KEY` | API key for the broker |

The broker exposes `GET /token` returning `{"access_token": "...", "expires_at": 1234567890.0}`.

### Local development — MSAL cache (device-code)

Leave `GRAPH_EXCEL_BROKER_URL` unset.  `LocalTokenProvider` picks up the token cache that the [excel-spike](../excel-spike) saved after a one-time device-code login:

```bash
cd ~/dev/excel-spike
python spike.py init   # prints a code — open the URL and sign in
python spike.py run    # polls until signed in, saves token_cache.bin
```

After that, every call to `open_workbook()` works silently (no browser needed) until the refresh token expires (~90 days on first use, indefinite with regular use).

---

## Quick start

```python
from tropi_excel import open_workbook

# Pass a SharePoint sharing URL or web URL.
wb = open_workbook("https://tropicommodity.sharepoint.com/:x:/...")

with wb.session():
    # Write a value and a formula.
    wb.set_range("Sheet1", "A1:B1",
                 values=[["Hello", None]],
                 formulas=[[None, "=1+1"]])

    # Read back.
    data = wb.read_range("Sheet1", "A1:B1")
    print(data["values"])   # [["Hello", 2]]

    # Table operations.
    wb.add_table_rows("Orders", [{"A": "SKU-001", "B": 10, "C": "2026-06-01"}])
    col = wb.table_column("Orders", "A")   # or wb.table_column("Orders", "SKU")

    # Worksheet management.
    wb.add_worksheet("NewSheet")
    wb.delete_worksheet("NewSheet")
```

---

## API reference

### `open_workbook(path_or_url, token_provider=None) → ExcelFileClient`

Resolve a URL or logical path and return a client.  Caches the `(drive_id, item_id)` pair for the process lifetime.

- **URL** (starts with `http`) → resolved via Graph `/shares/u!{base64url}/driveItem`.
- **Logical path** (e.g. `/SiteRoot/Finance/2026/budget.xlsx`) → resolved via tropi-storage-adapter routing (`M365_ROUTES` env var).

---

### `ExcelFileClient`

#### `session()` — context manager

Opens a Graph workbook session (`persistChanges=true`).  The session-id is sent on every API call.  If a call returns 404 (session expired), the session is recreated once and the call retried automatically.

All workbook methods must be called inside `with wb.session():`.

#### `read_range(sheet, address) → dict`

`GET /workbook/worksheets('{sheet}')/range(address='{address}')`.

Returns the full Graph range object.  Useful keys: `values`, `text`, `formulas`, `numberFormat`.

#### `set_range(sheet, address, values=None, formulas=None, number_format=None) → dict`

`PATCH` a range.  Pass 2-D arrays; use `None` in positions you do not want to overwrite.

```python
# Write a value to A1 and a formula to B1 in one call.
wb.set_range("Sheet1", "A1:B1",
             values=[["My text", None]],
             formulas=[[None, "=SUM(C1:C10)"]])
```

#### `table_column(table, column_name_or_letter) → list`

Read all data values from a named table column (header row excluded).

```python
ids = wb.table_column("Orders", "A")       # by column letter
ids = wb.table_column("Orders", "Order ID")  # by header name (case-insensitive)
```

#### `add_table_rows(table, rows, index=None) → dict`

`POST /workbook/tables('{table}')/rows/add`.

`rows` is a list of dicts mapping **column letter → value**.  Unspecified columns are filled with `null` (Excel keeps formulas / auto-fills).

```python
wb.add_table_rows("Orders", [
    {"A": "SKU-001", "C": 15},   # column B stays null
    {"A": "SKU-002", "B": "XL", "C": 5},
])
```

`index` is a 0-based row position to insert before.  Omit (or `None`) to append.

#### `add_worksheet(name) → dict`

Add a new worksheet.  Returns the Graph worksheet object.

#### `delete_worksheet(name) → None`

Delete a worksheet by name.

---

## Logical-path resolution

If you pass a logical path instead of a URL, the adapter uses `tropi-storage-adapter`'s routing:

```bash
export M365_SITE_HOSTNAME=tropicommodity.sharepoint.com
export M365_ROUTES='{"Onedrive Transfer": ["/sites/MultiPACK", "Documents"]}'
```

```python
wb = open_workbook("/Onedrive Transfer/Finance/2026/budget.xlsx")
```

---

## Errors

| Exception | When |
|---|---|
| `ExcelApiError(status, body)` | Graph returned an unexpected HTTP error |
| `ExcelAuthError` | Token acquisition failed |
| `ExcelNotFoundError` | Workbook, worksheet, or table not found (404) |
| `ExcelThrottledError` | Graph 429 / 503 after max retries |
| `ExcelResolveError` | Could not map path/URL to a workbook |

---

## Live test

```bash
cd ~/dev/tropi-excel-adapter
python tests/live_test.py "https://...sharepoint.com/..."
```

The test adds a temporary worksheet, writes a value + formula, reads back and verifies the formula computed correctly, then deletes the worksheet.  No artifacts remain in the workbook after a passing run.

---

## Concurrency

A per-workbook `threading.Lock` serialises write calls (`set_range`, `add_table_rows`, `add_worksheet`, `delete_worksheet`) from different threads within the same process.

---

## Retry behaviour

- HTTP **429 / 503**: honours `Retry-After` header (or falls back to exponential backoff), up to `MAX_RETRIES = 3`.
- HTTP **504** on `createSession`: retried up to `MAX_RETRIES` times.
- Session **expired 404**: session is recreated once, call retried.
