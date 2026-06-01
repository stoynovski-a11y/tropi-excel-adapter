#!/usr/bin/env python3
"""Live end-to-end test for tropi_excel against a real SharePoint workbook.

Usage
-----
    python tests/live_test.py <file-url>

Where <file-url> is the SharePoint sharing URL (or web URL) of a THROWAWAY
test workbook — the same one used by the excel-spike.  You can also set it
in ~/dev/excel-spike/.env as EXCEL_TEST_FILE.

Pre-requisite: you must have a saved MSAL token cache from the spike's
device-code flow:
    cd ~/dev/excel-spike
    python spike.py init   # prints a device code
    python spike.py run    # polls until you sign in, saves token_cache.bin

What the test does (mirroring the spike's excel_write function):
  1. Resolve the URL to (drive_id, item_id).
  2. Open a workbook session.
  3. add_worksheet("TROPI_LIVE_<timestamp>").
  4. set_range: write a text value to A1 and the formula =1+1 to B1.
  5. read_range A1:B1 — confirm A1 contains the text and B1's value == 2.
  6. delete_worksheet.
  7. Report PASS or FAIL.
"""
import os
import sys
import datetime

# Allow running from the repo root without installing the package.
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from tropi_excel import open_workbook
from tropi_excel.auth import LocalTokenProvider
from tropi_excel.errors import ExcelError


def _load_env(path: str) -> dict:
    result = {}
    try:
        with open(path) as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                result[k.strip()] = v.strip()
    except FileNotFoundError:
        pass
    return result


def main() -> int:
    spike_env = _load_env(os.path.expanduser("~/dev/excel-spike/.env"))

    # URL from CLI arg > EXCEL_TEST_FILE in spike .env > env var
    file_url = (
        sys.argv[1] if len(sys.argv) > 1
        else spike_env.get("EXCEL_TEST_FILE", "").strip()
        or os.getenv("EXCEL_TEST_FILE", "").strip()
    )

    if not file_url:
        print(
            "ERROR: No SharePoint URL supplied.\n"
            "  Pass it as the first argument, or set EXCEL_TEST_FILE in\n"
            "  ~/dev/excel-spike/.env (or as an env var).",
            file=sys.stderr,
        )
        return 1

    print(f"Target file URL: {file_url}")

    timestamp = datetime.datetime.now().strftime("%H%M%S")
    sheet_name = f"TROPI_LIVE_{timestamp}"
    text_value = f"tropi_excel live test {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

    tp = LocalTokenProvider()
    failures = []

    try:
        print("\n[1] Resolving workbook ...", end=" ", flush=True)
        wb = open_workbook(file_url, token_provider=tp)
        print("OK")

        with wb.session():
            print(f"[2] Opening workbook session ... OK")

            # --- add worksheet -------------------------------------------
            print(f"[3] Adding worksheet '{sheet_name}' ...", end=" ", flush=True)
            wb.add_worksheet(sheet_name)
            print("OK")

            # --- set_range -----------------------------------------------
            print("[4] Writing A1 (value) and B1 (formula =1+1) ...", end=" ", flush=True)
            wb.set_range(
                sheet_name,
                "A1:B1",
                values=[[text_value, None]],
                formulas=[[None, "=1+1"]],
            )
            print("OK")

            # --- read_range ----------------------------------------------
            print("[5] Reading back A1:B1 ...", end=" ", flush=True)
            result = wb.read_range(sheet_name, "A1:B1")
            print("OK")

            values = result.get("values", [[]])
            a1 = values[0][0] if values and values[0] else None
            b1 = values[0][1] if values and len(values[0]) > 1 else None

            print(f"    A1 value : {a1!r}")
            print(f"    B1 value : {b1!r}")

            if a1 != text_value:
                msg = f"A1 mismatch: expected {text_value!r}, got {a1!r}"
                print(f"  FAIL: {msg}")
                failures.append(msg)
            else:
                print("  A1 content matches. PASS")

            b1_num = float(b1) if b1 is not None else None
            if b1_num != 2.0:
                msg = f"B1 formula result: expected 2, got {b1!r}"
                print(f"  FAIL: {msg}")
                failures.append(msg)
            else:
                print("  B1 formula computed to 2. PASS")

            # --- delete worksheet ----------------------------------------
            print(f"[6] Deleting worksheet '{sheet_name}' ...", end=" ", flush=True)
            wb.delete_worksheet(sheet_name)
            print("OK")

        print("\n[7] Session closed.")

    except ExcelError as exc:
        print(f"\nEXCEPTION: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"\nUNEXPECTED EXCEPTION: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1

    if failures:
        print("\n=== LIVE TEST FAILED ===")
        for f in failures:
            print(f"  - {f}")
        return 1

    print("\n=== LIVE TEST PASSED ===")
    print(f"  Worksheet '{sheet_name}' was created, written, read back, and deleted.")
    print(f"  The workbook is clean — no test artifacts remain.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
