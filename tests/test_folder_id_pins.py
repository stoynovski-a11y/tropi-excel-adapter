"""Unit tests for M365_FOLDER_IDS folder-ID pins in logical-path resolution.

The pin layer makes the Excel-fill path rename/move-proof: a workbook under a
pinned folder is addressed by the folder's stable driveItem id
(``/drives/{drive}/items/{anchor}:/{remainder}``) instead of by name
(``/drives/{drive}/root:/{path}``).  Mirrors tropi_storage's _build_item_url.

Covers:
  - _match_pin: exact anchor, under anchor, sibling false-match guard, no pin
  - _resolve_from_path: pinned path → items/{id}:/rel URL
  - _resolve_from_path: unpinned path → root:/path URL (byte-identical legacy)
  - _pins_for_drive: keeps only pins routable to the queried drive, caches
"""
import sys

import pytest

import tropi_excel.resolve  # noqa: F401 — ensure the submodule is loaded
resolve_mod = sys.modules["tropi_excel.resolve"]

from tropi_excel.resolve import _match_pin


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


@pytest.fixture(autouse=True)
def _clear_caches():
    """Reset every module-level cache so tests don't bleed into each other."""
    for d in (
        resolve_mod._cache,
        resolve_mod._site_id_cache,
        resolve_mod._drive_id_cache,
        resolve_mod._pins_by_drive,
    ):
        d.clear()
    yield
    for d in (
        resolve_mod._cache,
        resolve_mod._site_id_cache,
        resolve_mod._drive_id_cache,
        resolve_mod._pins_by_drive,
    ):
        d.clear()


# ---------------------------------------------------------------------------
# _match_pin — pure logic
# ---------------------------------------------------------------------------

PINS = [("/Domestic/01 KeyAccounts", "ANCHOR1"), ("/Метро", "ANCHOR2")]


def test_match_pin_under_anchor():
    assert _match_pin("/Domestic/01 KeyAccounts/Билла/f.xlsx", PINS) == (
        "ANCHOR1",
        "/Билла/f.xlsx",
    )


def test_match_pin_exact_anchor():
    assert _match_pin("/Метро", PINS) == ("ANCHOR2", "/")


def test_match_pin_sibling_false_match_guarded():
    # "/Метро2" must NOT match the "/Метро" pin.
    assert _match_pin("/Метро2/x.xlsx", PINS) is None


def test_match_pin_no_pins():
    assert _match_pin("/anything/x.xlsx", []) is None


def test_match_pin_unrelated_path():
    assert _match_pin("/Export/05 INV/x.xlsx", PINS) is None


# ---------------------------------------------------------------------------
# _resolve_from_path — pinned vs name-based URL building
# ---------------------------------------------------------------------------

def _stub_routing(monkeypatch, *, item_path, drive_id="D1"):
    """Stub route + drive resolution so we isolate the URL-building branch."""
    monkeypatch.setattr(
        resolve_mod, "resolve_route",
        lambda path, *a, **k: ("/sites/x", "Lib", item_path),
    )
    monkeypatch.setattr(
        resolve_mod, "_resolve_drive_id",
        lambda *a, **k: drive_id,
    )
    monkeypatch.setenv("M365_SITE_HOSTNAME", "host.sharepoint.com")


def test_resolve_uses_pin_when_path_under_anchor(monkeypatch):
    _stub_routing(monkeypatch, item_path="/Domestic/01 KeyAccounts/Билла/f.xlsx")
    monkeypatch.setattr(
        resolve_mod, "_pins_for_drive",
        lambda *a, **k: [("/Domestic/01 KeyAccounts", "ANCHOR1")],
    )

    captured = {}

    def fake_get(url, headers=None, timeout=None):
        captured["url"] = url
        return _FakeResp(200, {"id": "FILEID"})

    monkeypatch.setattr(resolve_mod.requests, "get", fake_get)

    drive_id, item_id = resolve_mod._resolve_from_path("/whatever", "tok")
    assert (drive_id, item_id) == ("D1", "FILEID")
    assert captured["url"] == (
        f"{resolve_mod.GRAPH}/drives/D1/items/ANCHOR1:/"
        + "%D0%91%D0%B8%D0%BB%D0%BB%D0%B0/f.xlsx"
    )
    assert "root:" not in captured["url"]


def test_resolve_name_based_when_no_pin(monkeypatch):
    _stub_routing(monkeypatch, item_path="/Export/05 INV/f.xlsx")
    monkeypatch.setattr(resolve_mod, "_pins_for_drive", lambda *a, **k: [])

    captured = {}

    def fake_get(url, headers=None, timeout=None):
        captured["url"] = url
        return _FakeResp(200, {"id": "FILEID"})

    monkeypatch.setattr(resolve_mod.requests, "get", fake_get)

    drive_id, item_id = resolve_mod._resolve_from_path("/whatever", "tok")
    assert (drive_id, item_id) == ("D1", "FILEID")
    assert "/items/" not in captured["url"]
    assert captured["url"].startswith(f"{resolve_mod.GRAPH}/drives/D1/root:/")


# ---------------------------------------------------------------------------
# _pins_for_drive — drive filtering + caching
# ---------------------------------------------------------------------------

def test_pins_for_drive_filters_and_caches(monkeypatch):
    # Two configured pins in different libraries → different drives.
    monkeypatch.setattr(
        resolve_mod, "load_folder_pins",
        lambda: {"/LibA/Anchor": "IDA", "/LibB/Other": "IDB"},
    )

    def fake_route(logical, *a, **k):
        seg = logical.strip("/").split("/", 1)
        lib = seg[0]
        rest = "/" + (seg[1] if len(seg) > 1 else "")
        return ("/sites/x", lib, rest)

    monkeypatch.setattr(resolve_mod, "resolve_route", fake_route)

    drive_calls = {"n": 0}

    def fake_drive_id(sp, dn, hostname, token):
        drive_calls["n"] += 1
        return {"LibA": "D1", "LibB": "D2"}[dn]

    monkeypatch.setattr(resolve_mod, "_resolve_drive_id", fake_drive_id)

    pins_d1 = resolve_mod._pins_for_drive(
        "D1", None, None, None, None, "host", "tok"
    )
    assert pins_d1 == [("/Anchor", "IDA")]

    # Second call for the same drive is served from cache (no new drive lookups).
    before = drive_calls["n"]
    pins_d1_again = resolve_mod._pins_for_drive(
        "D1", None, None, None, None, "host", "tok"
    )
    assert pins_d1_again == [("/Anchor", "IDA")]
    assert drive_calls["n"] == before  # cached, no extra resolution

    pins_d2 = resolve_mod._pins_for_drive(
        "D2", None, None, None, None, "host", "tok"
    )
    assert pins_d2 == [("/Other", "IDB")]


def test_pins_for_drive_empty_when_unset(monkeypatch):
    monkeypatch.setattr(resolve_mod, "load_folder_pins", lambda: {})
    assert resolve_mod._pins_for_drive(
        "D1", None, None, None, None, "host", "tok"
    ) == []
