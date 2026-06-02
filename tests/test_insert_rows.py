"""Unit tests for insert_rows.

Asserts the entire-row address construction (so all columns shift together)
and the shift body, plus the count guard — without a live Graph session.
"""
import threading

import pytest

from tropi_excel.excelfile import ExcelFileClient


class _FakeResp:
    def json(self):
        return {"address": "ok"}


class _FakeSess:
    def __init__(self):
        self.calls = []

    def request(self, method, url, json=None):
        self.calls.append((method, url, json))
        return _FakeResp()


class _FakeEF:
    item_base_url = "https://graph.microsoft.com/v1.0/drives/D/items/I"


def _wb():
    wb = ExcelFileClient.__new__(ExcelFileClient)
    wb._write_lock = threading.Lock()
    wb._session = _FakeSess()
    wb._ef = _FakeEF()
    return wb


def test_insert_rows_builds_entire_row_address_and_shift_down():
    wb = _wb()
    wb.insert_rows("Продажби 2026", 3573, 2)
    method, url, body = wb._session.calls[0]
    assert method == "POST"
    # entire-row address (no column letters) => every column shifts together
    assert "range(address='3573:3574')/insert" in url
    assert body == {"shift": "Down"}


def test_insert_rows_single_row_address():
    wb = _wb()
    wb.insert_rows("S", 10, 1)
    _, url, _ = wb._session.calls[0]
    assert "range(address='10:10')/insert" in url


@pytest.mark.parametrize("bad", [0, -1, -5])
def test_insert_rows_rejects_nonpositive_count(bad):
    wb = ExcelFileClient.__new__(ExcelFileClient)
    with pytest.raises(ValueError):
        wb.insert_rows("S", 5, bad)
