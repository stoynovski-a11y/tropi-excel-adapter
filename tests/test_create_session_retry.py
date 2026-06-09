"""Unit tests for createSession retry behaviour.

Covers the opt-in 404-retry path used right after a server-side copy ("new"
mode), where the copied workbook briefly returns 404 itemNotFound until
SharePoint activates it for the Excel API. Without opt-in, a 404 must still
fail fast (so genuine missing-file errors don't hang). 504 retries regardless.
"""
import pytest

import tropi_excel.excelfile as ef_mod
from tropi_excel.errors import ExcelNotFoundError
from tropi_excel.excelfile import _ExcelSession


class _FakeResp:
    def __init__(self, status_code, payload=None):
        self.status_code = status_code
        self._payload = payload or {}
        self.text = "" if status_code < 400 else f"error {status_code}"

    @property
    def ok(self):
        return self.status_code < 400

    def json(self):
        return self._payload


class _FakeEF:
    item_base_url = "https://graph.microsoft.com/v1.0/drives/D/items/I"


class _FakeTokenProvider:
    def get_token(self):
        return "tok"


def _session():
    return _ExcelSession(_FakeEF(), _FakeTokenProvider())


@pytest.fixture(autouse=True)
def _no_sleep(monkeypatch):
    """Make backoff instant and record how many times it slept."""
    sleeps = []
    monkeypatch.setattr(ef_mod.time, "sleep", lambda s: sleeps.append(s))
    return sleeps


def _patch_responses(monkeypatch, responses):
    """Make requests.post return each response in turn; record call count."""
    state = {"i": 0}

    def fake_post(*args, **kwargs):
        r = responses[min(state["i"], len(responses) - 1)]
        state["i"] += 1
        return r

    monkeypatch.setattr(ef_mod.requests, "post", fake_post)
    return state


def test_404_then_ok_when_opted_in(monkeypatch, _no_sleep):
    state = _patch_responses(monkeypatch, [
        _FakeResp(404), _FakeResp(404), _FakeResp(200, {"id": "S"}),
    ])
    sess = _session()
    assert sess._create_session(retry_not_found=True) == "S"
    assert state["i"] == 3          # two retries, then success
    assert len(_no_sleep) == 2      # slept before each retry


def test_404_fails_fast_when_not_opted_in(monkeypatch, _no_sleep):
    state = _patch_responses(monkeypatch, [_FakeResp(404), _FakeResp(200, {"id": "S"})])
    sess = _session()
    with pytest.raises(ExcelNotFoundError):
        sess._create_session()              # default: retry_not_found=False
    assert state["i"] == 1                   # no retry on 404
    assert len(_no_sleep) == 0


def test_504_retries_even_without_opt_in(monkeypatch, _no_sleep):
    _patch_responses(monkeypatch, [_FakeResp(504), _FakeResp(200, {"id": "S"})])
    sess = _session()
    assert sess._create_session() == "S"
    assert len(_no_sleep) == 1


def test_persistent_404_raises_after_budget(monkeypatch, _no_sleep):
    state = _patch_responses(monkeypatch, [_FakeResp(404)])
    sess = _session()
    with pytest.raises(ExcelNotFoundError):
        sess._create_session(retry_not_found=True)
    # budget = CREATE_SESSION_NOT_FOUND_RETRIES retries + 1 final attempt
    assert state["i"] == ef_mod.CREATE_SESSION_NOT_FOUND_RETRIES + 1
    assert len(_no_sleep) == ef_mod.CREATE_SESSION_NOT_FOUND_RETRIES
