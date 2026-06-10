"""Unit tests for the stale-item-id cache-heal in createSession.

When a caller deletes a SharePoint file and re-copies a template to the same
path, the file gets a new item_id.  The next createSession returns 404
itemNotFound.  The fix: on 404, invalidate the cache, re-resolve the path, and
if the item_id changed, retry immediately with the fresh item.

Covers:
  a. 404 + new item_id → retried with fresh id, succeeds, 2 POSTs, 0 sleeps.
  b. 404 + same item_id from re_resolve → raises ExcelNotFoundError, 1 POST.
  c. 404 + re_resolve returns None → raises ExcelNotFoundError fast.
  d. Heal runs at most once: [404, 404] with new id each time → still raises
     after second 404 (2 POSTs total).
  e. resolve() cache: second call hits cache; invalidate() + re-resolve gets
     fresh item_id.
"""
import sys
import pytest

import tropi_excel.excelfile as ef_mod
# Import the submodule before tropi_excel.__init__ shadows the name with the
# function.  Retrieve the real module object via sys.modules.
import tropi_excel.resolve  # noqa: F401 — ensure the submodule is loaded
resolve_mod = sys.modules["tropi_excel.resolve"]

from tropi_excel.errors import ExcelNotFoundError
from tropi_excel.excelfile import _ExcelSession
from tropi_excel.resolve import ExcelFile, invalidate, resolve


# ---------------------------------------------------------------------------
# Shared test doubles (copied from test_create_session_retry.py style)
# ---------------------------------------------------------------------------

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


class _FakeTokenProvider:
    def get_token(self):
        return "tok"


def _patch_responses(monkeypatch, responses):
    """Make requests.post return each response in turn; record call count."""
    state = {"i": 0}

    def fake_post(*args, **kwargs):
        r = responses[min(state["i"], len(responses) - 1)]
        state["i"] += 1
        return r

    monkeypatch.setattr(ef_mod.requests, "post", fake_post)
    return state


@pytest.fixture(autouse=True)
def _no_sleep(monkeypatch):
    """Make backoff instant and record how many times it slept."""
    sleeps = []
    monkeypatch.setattr(ef_mod.time, "sleep", lambda s: sleeps.append(s))
    return sleeps


def _make_ef(item_id, source_key="https://x/file"):
    return ExcelFile(
        drive_id="D1",
        item_id=item_id,
        source_key=source_key,
    )


# ---------------------------------------------------------------------------
# a. 404 + new item_id → succeeds, 2 POSTs, 0 sleeps
# ---------------------------------------------------------------------------

def test_heal_new_item_id_succeeds(monkeypatch, _no_sleep):
    state = _patch_responses(monkeypatch, [
        _FakeResp(404),
        _FakeResp(200, {"id": "SESSION1"}),
    ])

    new_ef = _make_ef("I_NEW")
    calls = []

    def re_resolve():
        calls.append(1)
        return new_ef

    sess = _ExcelSession(_make_ef("I_OLD"), _FakeTokenProvider(), re_resolve=re_resolve)
    result = sess._create_session(retry_not_found=False)

    assert result == "SESSION1"
    assert state["i"] == 2          # two POST calls
    assert len(_no_sleep) == 0      # no sleep on the heal path
    assert len(calls) == 1          # re_resolve called exactly once


# ---------------------------------------------------------------------------
# b. 404 + same item_id from re_resolve → raises, 1 POST
# ---------------------------------------------------------------------------

def test_heal_same_item_id_raises(monkeypatch, _no_sleep):
    state = _patch_responses(monkeypatch, [
        _FakeResp(404),
        _FakeResp(200, {"id": "SESSION1"}),   # would succeed if retried — but shouldn't be
    ])

    same_ef = _make_ef("I_SAME")

    def re_resolve():
        return same_ef

    sess = _ExcelSession(_make_ef("I_SAME"), _FakeTokenProvider(), re_resolve=re_resolve)
    with pytest.raises(ExcelNotFoundError):
        sess._create_session(retry_not_found=False)

    assert state["i"] == 1          # no extra POST after same-id heal
    assert len(_no_sleep) == 0


# ---------------------------------------------------------------------------
# c. 404 + re_resolve returns None → raises ExcelNotFoundError fast
# ---------------------------------------------------------------------------

def test_heal_re_resolve_none_raises(monkeypatch, _no_sleep):
    state = _patch_responses(monkeypatch, [
        _FakeResp(404),
        _FakeResp(200, {"id": "SESSION1"}),   # should not be reached
    ])

    def re_resolve():
        return None

    sess = _ExcelSession(_make_ef("I_OLD"), _FakeTokenProvider(), re_resolve=re_resolve)
    with pytest.raises(ExcelNotFoundError):
        sess._create_session(retry_not_found=False)

    assert state["i"] == 1          # only the initial 404 call
    assert len(_no_sleep) == 0


# ---------------------------------------------------------------------------
# d. Heal runs at most once: [404, 404, ...] with new id → raises after 2nd 404
# ---------------------------------------------------------------------------

def test_heal_runs_at_most_once(monkeypatch, _no_sleep):
    state = _patch_responses(monkeypatch, [
        _FakeResp(404),   # triggers heal → new item
        _FakeResp(404),   # second 404 with new item — heal already used
    ])

    call_count = []

    def re_resolve():
        call_count.append(1)
        return _make_ef(f"I_NEW_{len(call_count)}")

    sess = _ExcelSession(_make_ef("I_OLD"), _FakeTokenProvider(), re_resolve=re_resolve)
    with pytest.raises(ExcelNotFoundError):
        sess._create_session(retry_not_found=False)

    assert state["i"] == 2          # first 404 (heal) + second 404 (raises)
    assert len(call_count) == 1     # heal invoked only once
    assert len(_no_sleep) == 0


# ---------------------------------------------------------------------------
# e. resolve() cache test
# ---------------------------------------------------------------------------

@pytest.fixture(autouse=True)
def _clear_resolve_cache():
    """Ensure a clean resolve cache for each test that uses it."""
    resolve_mod._cache.clear()
    yield
    resolve_mod._cache.clear()


def test_resolve_cache_and_invalidate(monkeypatch):
    """Second resolve() call hits cache; after invalidate() gets fresh item_id."""
    call_count = []

    def fake_get(url, headers=None, timeout=None):
        call_count.append(url)
        # Return current item_id based on how many times we've been called
        item_id = "I1" if len(call_count) <= 1 else "I2"
        return _FakeResp(200, {"parentReference": {"driveId": "D1"}, "id": item_id})

    monkeypatch.setattr(resolve_mod.requests, "get", fake_get)

    provider = _FakeTokenProvider()
    url = "https://x/file"

    # First call — goes to network
    ef1 = resolve(url, provider)
    assert ef1.item_id == "I1"
    assert ef1.source_key == url
    assert len(call_count) == 1

    # Second call — hits cache, no HTTP request
    ef2 = resolve(url, provider)
    assert ef2.item_id == "I1"
    assert len(call_count) == 1     # still just 1 call

    # Invalidate, then re-resolve — goes to network again, gets new id
    invalidate(url)
    ef3 = resolve(url, provider)
    assert ef3.item_id == "I2"
    assert ef3.source_key == url
    assert len(call_count) == 2     # second HTTP call made
