# RUNBOOK — tropi-excel-adapter (shared library)

**This is a Python library, not a service.** It does not run anywhere, has no server, no cron, no Railway project. Services import it (`import tropi_excel`) and pin a git commit in their `requirements.txt`. "Deploying" it = pushing a commit to GitHub, then re-pinning consumers.

## Purpose

Lets the fleet edit Excel workbooks that live on SharePoint **in place** — by cell, range, or table row — through the Microsoft Graph Excel REST API, instead of the old fragile cycle of download → openpyxl edit → re-upload (the #1 source of OOXML corruption). Eight production services (plus the planned invoice-generator) write accounting registers, sales books, and chain order files through this one library, so a bug here can corrupt or block writes fleet-wide.

**Auth is delegated-only**: the Graph workbook API does not support app-only tokens, so production tokens come from the **graph-excel-token-broker** Railway service (a separate repo at `~/dev/graph-excel-token-broker` holding a service-account refresh token; it exposes `GET /token` guarded by `X-API-Key`).

## Public API surface

Everything is exported from `tropi_excel/__init__.py` (current `__version__ = "0.2.1"`).

| Symbol | What it does | File |
|---|---|---|
| `open_workbook(path_or_url, token_provider=None)` | Entry point. Resolves a SharePoint URL **or** logical path to `(drive_id, item_id)`, returns `ExcelFileClient`. Result cached for process life. | `tropi_excel/__init__.py:39` |
| `wb.session(retry_not_found=False)` | Context manager — opens a Graph workbook session (`persistChanges=true`). **All other methods must run inside it.** `retry_not_found=True` = use right after a server-side copy (retries 404 up to ~31.5 s while SharePoint activates the new file). | `excelfile.py:284` |
| `wb.read_range(sheet, address)` | GET a range; returns Graph dict (`values`, `text`, `formulas`, `numberFormat`). | `excelfile.py:327` |
| `wb.set_range(sheet, address, values=, formulas=, number_format=)` | PATCH a range. 2-D arrays; `None` = leave cell alone. | `excelfile.py:344` |
| `wb.set_ranges(ops)` | Many range writes via Graph `$batch` (chunks of 20). Applies all `number_format` writes **before** any values — required for the Text-`@` gotcha below. | `excelfile.py:387` |
| `wb.insert_rows(sheet, start_row, count)` | Insert blank ENTIRE rows, shift Down; new rows inherit the row above's formatting. For sheets without a table. | `excelfile.py:459` |
| `wb.add_table_rows(table, rows, index=None)` | Append/insert table rows; `rows` = dicts of **worksheet column letter → value** (converted internally to table-relative positions). | `excelfile.py:644` |
| `wb.table_column(table, name_or_letter)` | Read one table column's data values. Addresses by position (`itemAt`), never by Graph's internal column id. | `excelfile.py:532` |
| `wb.table_data_start_row(table)` / `wb.table_name_for_sheet(sheet)` | Helpers: 1-based first data row; runtime table-name discovery. | `excelfile.py:612` / `:627` |
| `wb.set_font_color(sheet, address, "#RRGGBB")` | Color a range's font (e.g. flag a cell red). | `excelfile.py:500` |
| `wb.add_worksheet(name)` / `wb.delete_worksheet(name)` | Blank sheet add/delete. **No sheet clone/copy** — Graph REST has none (why warehouse-transfer/kasa use tropi-excelops native-clone instead). | `excelfile.py:689` / `:705` |
| `resolve(path) / invalidate(path)` | Path→item resolution and manual cache drop. | `resolve.py:64` / `:88` |
| Errors: `ExcelApiError(status, body)`, `ExcelAuthError`, `ExcelNotFoundError`, `ExcelThrottledError`, `ExcelResolveError` | All inherit `ExcelError`. | `errors.py` |

Automatic value coercion on write (`excelfile.py:725`): `datetime`/`date` → Excel serial numbers, `Decimal` → float.

## Env vars / config it reads

The library has no config file — it reads env vars at call time:

| Var | What it does | Secret | Notes |
|---|---|---|---|
| `GRAPH_EXCEL_BROKER_URL` | Token broker base URL. **If set → production `BrokerTokenProvider`; if unset → `LocalTokenProvider`** (MSAL cache from `~/dev/excel-spike/token_cache.bin`). | no | Selector logic: `auth.py:202` (`get_token_provider`). Prod value: `https://graph-excel-token-broker-production.up.railway.app` (verified live on Railway sales-autofill env, 2026-06-11; URL answers 401 without a key) |
| `GRAPH_EXCEL_BROKER_API_KEY` | Sent as `X-API-Key` to the broker's `GET /token`. | **yes** | `auth.py:58` |
| `M365_TENANT_ID`, `M365_CLIENT_ID` | Only for local dev (`LocalTokenProvider`); read from env or `~/dev/excel-spike/.env`. | tenant no / client no | `auth.py:125` |
| `M365_SITE_HOSTNAME` | Required for **logical-path** resolution. Actual fleet value: `vaklin.sharepoint.com` (verified live on Railway sales-autofill env, 2026-06-11). The README's `tropicommodity.sharepoint.com` is an example only. | no | `resolve.py:131` |
| `M365_ROUTES`, `M365_STRIP_PREFIX`, `M365_SITE_PATH`, `M365_DEFAULT_LIBRARY` | Logical-path routing, delegated to `tropi_storage.routing` (tropi-storage-adapter must be installed for this mode). | no | `resolve.py:118-148` |

**Rollback flags live in the CONSUMERS, not here**: each service gates its Graph path behind `EXCEL_BACKEND=graph` (flip to `openpyxl` to fall back to download-edit-upload) — keyaccounts uses `EXCEL_GRAPH_CHAINS` (remove a chain name to roll that chain back). Flipping a consumer flag stops it calling this library entirely.

## Consumers and pins

Repo is **public** on GitHub (`stoynovski-a11y/tropi-excel-adapter`) so Railway builds can pip-install it without credentials. Pins as of 2026-06-11 (HEAD = `6bea9c1`):

| Consumer (`~/dev/...`) | Pin |
|---|---|
| svedenie-generator | `@6bea9c1` (HEAD — has the faster 0.5 s not-found retry) |
| railway-bank-parser, railway-invoice-renamer, sales-autofill, railway-keyaccounts, railway-warehouse-receipts, railway-metro, metro-order-parser, invoice-generator (planned, not yet live) | `@321594d` (has the stale-cache self-heal, lacks `6bea9c1` retry tuning — functionally fine) |
| warehouse-transfer, kasa-automation | **NOT consumers** — they use `tropi-excelops` (sheet-clone lib, separate repo) |

To bump a consumer: edit its `requirements.txt` git pin → commit → its normal deploy path (`railway up` or auto-deploy, per that service's runbook).

## "Deploy" (library release flow)

1. Change code on `main`, run tests: `cd ~/dev/tropi-excel-adapter && python -m pytest` (unit tests are offline/mocked).
2. Optional live check against a scratch workbook: `python tests/live_test.py "https://...sharepoint.com/..."` (adds a temp sheet, writes, verifies, deletes — leaves no artifacts).
3. `git push` (user does this) → note the new short SHA.
4. Re-pin consumers one at a time; nothing updates automatically — **a push here changes zero running services until a consumer re-pins and redeploys.**

No Procfile / Dockerfile / railway.json / CI workflows exist in this repo (verified).

## Failure modes & debugging

| Symptom | Where to look | Recovery |
|---|---|---|
| `ExcelAuthError: Token broker returned 401/5xx` in a consumer | Broker service logs (Railway, graph-excel-token-broker); `GET /token` needs valid `X-API-Key` | Check the broker's no-auth `GET /health` endpoint; if the service-account refresh token died, re-mint it via device code: `cd ~/dev/graph-excel-token-broker && python mint_token.py init` (prints device code + URL) then `python mint_token.py run` (prints the new refresh token) → paste it into `GRAPH_EXCEL_REFRESH_TOKEN` on the broker's Railway env. Rotate `GRAPH_EXCEL_BROKER_API_KEY` on broker + all consumers together. |
| `createSession` 404 right after a service deletes + re-copies a file to the **same path** | Consumer logs: `createSession 404 — stale item id, re-resolved ... → retrying` (warning, then success) | Self-heals since `321594d`: cache invalidated, path re-resolved, retried (`excelfile.py:152`, `resolve.py:88`). If a consumer is pinned older than `321594d` it loops 404 forever — re-pin. |
| 404 on a freshly server-side-copied workbook ("new"/draft mode) | Consumer didn't pass `retry_not_found=True` to `session()` | Caller must opt in: `with wb.session(retry_not_found=True):` — gives ~31.5 s activation budget (`excelfile.py:62-68`). |
| `ExcelThrottledError` (Graph 429/503) | Consumer logs: `Graph 429 throttled, retry n/3` | Library honours `Retry-After` and retries 3× automatically (`excelfile.py:224`). Persistent throttling = too many concurrent sessions on one workbook — note the lock only serialises writes **within one process**, not across services. |
| Numbers written where text was expected — product codes lose leading zeros, `XLOOKUP` → `#N/A` | The written register itself | Graph `set_range` coerces numeric-looking strings to numbers. Caller must PATCH `number_format=[["@"]]` (Text) on those cells **before** writing values. `set_ranges()` does this ordering automatically (`excelfile.py:394-399`). |
| Wrong column written in a table | — | Use letters/headers only; the lib already converts worksheet letters to table-relative positions and never uses Graph's internal column ids (`excelfile.py:546-549`). If a table doesn't start at column A, raw indexes will be off — use the helpers. |

**Cost note:** this library calls only Microsoft Graph (included in the M365 subscription) — no per-call paid API, so no spend caps/attempt ledgers live here. Those guards belong to the consumers (Gemini caps in bank-parser / invoice-renamer).

## Gotchas

- **Delegated-only auth.** The Graph workbook API rejects app-only tokens — you cannot reuse tropi-storage-adapter's client-credentials flow for cell edits. That is the entire reason the broker + service account exists.
- **Everything must be inside `with wb.session():`** — calling a method outside raises `RuntimeError` (`excelfile.py:314`). Session expiry mid-block is handled (recreate once + retry).
- **Process-lifetime caches** in `resolve.py`: path→item_id, site-id, drive-id. The item-id cache self-heals on createSession 404 (`321594d`), but the site/drive caches don't — renaming a library or site requires a consumer restart.
- **No worksheet clone.** `add_worksheet` creates a *blank* sheet. Formatted-template cloning is `tropi-excelops` (Office-Scripts-style native clone), a different library.
- **Mid-flight failures are not transactional.** `persistChanges=true` saves as you go; a crashed multi-write flow leaves a partially written workbook. `set_ranges` is designed to be safely re-run (fixed-cell overwrites), but row-inserting flows need their own idempotency (consumers use Redis ledgers for that).
- **Two version numbers disagree**: `pyproject.toml` says `0.1.0`, `tropi_excel/__init__.py` says `0.2.1`. Pin by git SHA (everyone does), not by version.
- **Write lock is per-process only** (`excelfile.py:72`). Two Railway services writing the same workbook concurrently are NOT serialised — Graph sessions can conflict. The fleet design gives each workbook a single writing service (not exhaustively verified across consumers) — keep it that way.
