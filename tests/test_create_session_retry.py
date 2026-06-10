"""Unit tests for createSession retry behaviour.

Covers the opt-in 404-retry path used right after a server-side copy ("new"
mode), where the copied workbook briefly returns 404 itemNotFound until
SharePoint activates it for the Excel API. Without opt-in, a 404 must still
fail fast (so genuine missing-file errors don't hang). 504 retries regardless.

Not-found retries use NOT_FOUND_BACKOFF_BASE (0.5 s) as the starting delay,
doubling each time: 0.5, 1, 2, 4, 8, 16 (6 sleeps, ~31.5 s total budget).
504 retries keep BACKOFF_BASE (1.0 s) as the starting delay.
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
    # First delay must use NOT_FOUND_BACKOFF_BASE (0.5 s), not BACKOFF_BASE (1.0 s)
    assert _no_sleep[0] == pytest.approx(ef_mod.NOT_FOUND_BACKOFF_BASE)
    # Second delay doubles: 0.5 * 2^1 = 1.0 s
    assert _no_sleep[1] == pytest.approx(ef_mod.NOT_FOUND_BACKOFF_BASE * 2)


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
    # 504 path must keep BACKOFF_BASE (1.0 s) — NOT_FOUND_BACKOFF_BASE must not bleed in
    assert _no_sleep[0] == pytest.approx(ef_mod.BACKOFF_BASE)


def test_persistent_404_raises_after_budget(monkeypatch, _no_sleep):
    state = _patch_responses(monkeypatch, [_FakeResp(404)])
    sess = _session()
    with pytest.raises(ExcelNotFoundError):
        sess._create_session(retry_not_found=True)
    # budget = CREATE_SESSION_NOT_FOUND_RETRIES retries + 1 final attempt
    assert state["i"] == ef_mod.CREATE_SESSION_NOT_FOUND_RETRIES + 1
    assert len(_no_sleep) == ef_mod.CREATE_SESSION_NOT_FOUND_RETRIES


def test_not_found_backoff_sequence(monkeypatch, _no_sleep):
    """Full not-found retry backoff sequence: 0.5, 1, 2, 4, 8, 16 (6 sleeps)."""
    state = _patch_responses(monkeypatch, [_FakeResp(404)])
    sess = _session()
    with pytest.raises(ExcelNotFoundError):
        sess._create_session(retry_not_found=True)

    assert len(_no_sleep) == ef_mod.CREATE_SESSION_NOT_FOUND_RETRIES  # 6
    base = ef_mod.NOT_FOUND_BACKOFF_BASE
    cap = ef_mod.BACKOFF_CAP
    expected = [min(base * (2 ** i), cap) for i in range(ef_mod.CREATE_SESSION_NOT_FOUND_RETRIES)]
    for i, (actual, exp) in enumerate(zip(_no_sleep, expected)):
        assert actual == pytest.approx(exp), f"sleep[{i}]: expected {exp}, got {actual}"


def test_504_backoff_unchanged(monkeypatch, _no_sleep):
    """504 retry uses BACKOFF_BASE (1.0 s), not NOT_FOUND_BACKOFF_BASE."""
    # Three 504s then success — checks the first three delays.
    _patch_responses(monkeypatch, [
        _FakeResp(504), _FakeResp(504), _FakeResp(504), _FakeResp(200, {"id": "S"}),
    ])
    sess = _session()
    assert sess._create_session() == "S"
    assert len(_no_sleep) == 3
    base = ef_mod.BACKOFF_BASE
    cap = ef_mod.BACKOFF_CAP
    expected = [min(base * (2 ** i), cap) for i in range(3)]
    for i, (actual, exp) in enumerate(zip(_no_sleep, expected)):
        assert actual == pytest.approx(exp), f"sleep[{i}]: expected {exp}, got {actual}"
