"""Unit tests for set_ranges (Graph $batch) — without a live Graph session.

Asserts: number_format writes are flushed before value writes, sub-request
URLs/bodies are well-formed, chunking respects BATCH_MAX, and a failing
sub-response raises.
"""
import threading

import pytest

from tropi_excel.errors import ExcelApiError
from tropi_excel.excelfile import BATCH_MAX, ExcelFileClient


class _Resp:
    def __init__(self, payload):
        self._payload = payload

    def json(self):
        return self._payload


class _BatchSess:
    """Records each POST /$batch and replies with a matching success set."""
    session_id = "sess-1"

    def __init__(self):
        self.batches = []  # list of request-payloads (the {"requests": [...]})

    def request(self, method, url, json=None):
        self.batches.append((method, url, json))
        n = len(json["requests"]) if json and "requests" in json else 0
        return _Resp({"responses": [{"id": str(i), "status": 200} for i in range(n)]})


class _FakeEF:
    item_base_url = "https://graph.microsoft.com/v1.0/drives/D/items/I"


def _wb():
    wb = ExcelFileClient.__new__(ExcelFileClient)
    wb._write_lock = threading.Lock()
    wb._session = _BatchSess()
    wb._ef = _FakeEF()
    return wb


def _bodies(batch_payload):
    return [r["body"] for r in batch_payload["requests"]]


def test_formats_flushed_before_values():
    wb = _wb()
    wb.set_ranges([
        {"sheet": "S", "address": "A1", "number_format": [["@"]], "values": [["007"]]},
        {"sheet": "S", "address": "B1", "values": [[5]]},  # numeric, no format
    ])
    calls = wb._session.batches
    assert len(calls) == 2  # one format phase, one value phase
    # Phase 1: only numberFormat sub-bodies.
    assert all("numberFormat" in b and "values" not in b for b in _bodies(calls[0][2]))
    # Phase 2: only value sub-bodies, both cells.
    assert all("values" in b and "numberFormat" not in b for b in _bodies(calls[1][2]))
    assert len(calls[1][2]["requests"]) == 2


def test_sub_request_url_is_relative_workbook_path():
    wb = _wb()
    wb.set_ranges([{"sheet": "Заявка София", "address": "D3", "values": [["x"]]}])
    req = wb._session.batches[0][2]["requests"][0]
    assert req["method"] == "PATCH"
    assert req["url"] == "/drives/D/items/I/workbook/worksheets('Заявка София')/range(address='D3')"
    assert req["headers"]["workbook-session-id"] == "sess-1"


def test_chunks_at_batch_max():
    wb = _wb()
    ops = [{"sheet": "S", "address": f"A{i}", "values": [[i]]} for i in range(BATCH_MAX + 5)]
    wb.set_ranges(ops)
    # No formats → only value batches: one full chunk + remainder.
    sizes = [len(b[2]["requests"]) for b in wb._session.batches]
    assert sizes == [BATCH_MAX, 5]


def test_raises_on_sub_failure():
    wb = _wb()

    def failing_request(method, url, json=None):
        return _Resp({"responses": [{"id": "0", "status": 400, "body": {"error": "bad"}}]})

    wb._session.request = failing_request
    with pytest.raises(ExcelApiError):
        wb.set_ranges([{"sheet": "S", "address": "A1", "values": [["x"]]}])


def test_empty_ops_makes_no_calls():
    wb = _wb()
    wb.set_ranges([])
    assert wb._session.batches == []
