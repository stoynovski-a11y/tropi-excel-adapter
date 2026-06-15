"""Resolve a logical path or SharePoint URL to (drive_id, item_id).

Two resolution strategies:

1. URL  (starts with "http")
   Uses the Graph /shares/{encoded_url}/driveItem endpoint — exactly as the
   excel-spike's ``item_from_url`` function. Works for any SharePoint
   sharing/web URL without knowing the site topology.

2. Logical path  (e.g. "/Top A/Finance/2026/budget.xlsx")
   Uses tropi-storage-adapter's routing module to map the first path segment
   to (site_path, library_name), then resolves the site ID + drive ID via
   Graph, then fetches the item ID via
   ``GET /sites/{siteId}/drives/{driveId}/root:/{path}:``.

   Folder-ID pins
   --------------
   If the optional ``M365_FOLDER_IDS`` env var is set (the SAME var the
   tropi_storage backend honours — JSON mapping a logical folder path to a
   stable driveItem id), a path that falls under a pinned folder is addressed
   by ``/drives/{drive_id}/items/{anchor_id}:/{remainder}`` instead of by name
   (``/drives/{drive_id}/root:/{path}``).  Because the driveItem id survives a
   folder rename/move within the drive, the Excel-fill path becomes
   rename-proof too — matching the storage adapter.  Unset/empty => normal
   name-based addressing (byte-identical to the pre-pin behaviour).

Both strategies cache results in a module-level dict keyed by the input
path/url so repeated calls to the same file are free.
"""
from __future__ import annotations

import base64
import os
import threading
import urllib.parse
from dataclasses import dataclass
from typing import TYPE_CHECKING

import requests

from .errors import ExcelApiError, ExcelResolveError

if TYPE_CHECKING:
    from .auth import TokenProvider

GRAPH = "https://graph.microsoft.com/v1.0"

# tropi-storage-adapter routing — imported at module level so it can be
# monkeypatched in tests and so URL-only callers (who never need it) still work
# when the package is absent.
try:
    from tropi_storage.routing import (
        load_folder_pins,
        load_routes,
        load_strip_prefix,
        resolve_route,
    )
    from tropi_storage.path_utils import normalize_path

    _HAS_STORAGE = True
except ImportError:  # pragma: no cover - exercised only when the dep is missing
    load_folder_pins = load_routes = load_strip_prefix = resolve_route = None  # type: ignore
    normalize_path = None  # type: ignore
    _HAS_STORAGE = False

# Module-level resolution cache: key → (drive_id, item_id)
_cache: dict[str, tuple[str, str]] = {}
_cache_lock = threading.Lock()

# Module-level site-id / drive-id caches (same pattern as graph_backend.py)
_site_id_cache: dict[str, str] = {}
_drive_id_cache: dict[tuple[str, str | None], str] = {}
# Per-drive resolved folder pins: drive_id → [(anchor_item_path, anchor_item_id)]
_pins_by_drive: dict[str, list[tuple[str, str]]] = {}
_id_cache_lock = threading.Lock()


@dataclass(frozen=True)
class ExcelFile:
    """Resolved workbook reference: drive ID + item ID."""

    drive_id: str
    item_id: str
    source_key: str | None = None

    @property
    def item_base_url(self) -> str:
        return f"{GRAPH}/drives/{self.drive_id}/items/{self.item_id}"


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def resolve(path_or_url: str, token_provider: "TokenProvider") -> ExcelFile:
    """Resolve *path_or_url* to an ``ExcelFile``.

    Results are cached for the lifetime of the process.  Pass a URL or a
    logical path — the function auto-detects.
    """
    key = path_or_url.strip()
    with _cache_lock:
        if key in _cache:
            drive_id, item_id = _cache[key]
            return ExcelFile(drive_id=drive_id, item_id=item_id, source_key=key)

    token = token_provider.get_token()

    if key.lower().startswith("http"):
        drive_id, item_id = _resolve_from_url(key, token)
    else:
        drive_id, item_id = _resolve_from_path(key, token)

    with _cache_lock:
        _cache[key] = (drive_id, item_id)
    return ExcelFile(drive_id=drive_id, item_id=item_id, source_key=key)


def invalidate(path_or_url: str) -> None:
    """Drop the cached (drive_id, item_id) for a path, if present."""
    key = path_or_url.strip()
    with _cache_lock:
        _cache.pop(key, None)


# ---------------------------------------------------------------------------
# URL strategy (spike's item_from_url, exactly)
# ---------------------------------------------------------------------------

def _resolve_from_url(url: str, token: str) -> tuple[str, str]:
    """SharePoint sharing URL / webUrl -> (driveId, itemId)."""
    share_id = "u!" + base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    r = requests.get(
        f"{GRAPH}/shares/{share_id}/driveItem",
        headers=_auth_headers(token),
        timeout=30,
    )
    _raise_for_status(r)
    j = r.json()
    drive_id: str = j["parentReference"]["driveId"]
    item_id: str = j["id"]
    return drive_id, item_id


# ---------------------------------------------------------------------------
# Logical-path strategy
# ---------------------------------------------------------------------------

def _resolve_from_path(path: str, token: str) -> tuple[str, str]:
    """Logical path → (drive_id, item_id) via tropi-storage-adapter routing."""
    if not _HAS_STORAGE:
        raise ExcelResolveError(
            "tropi-storage-adapter is not installed; "
            "install it or pass a SharePoint URL instead of a logical path."
        )

    routes = load_routes()
    strip_prefix = load_strip_prefix()
    hostname = os.getenv("M365_SITE_HOSTNAME", "")
    default_site = os.getenv("M365_SITE_PATH", "") or None
    default_drive = os.getenv("M365_DEFAULT_LIBRARY", "") or None

    if not hostname:
        raise ExcelResolveError(
            "M365_SITE_HOSTNAME is not set — required for logical-path resolution."
        )

    try:
        site_path, drive_name, item_path = resolve_route(
            path, routes, default_site, default_drive, strip_prefix=strip_prefix
        )
    except Exception as exc:
        raise ExcelResolveError(
            f"Cannot route path {path!r}: {exc}"
        ) from exc

    drive_id = _resolve_drive_id(site_path, drive_name, hostname, token)

    p = normalize_path(item_path)
    if p == "/":
        raise ExcelResolveError(
            f"Path {path!r} resolved to the library root — specify a file path."
        )

    item_url = _build_item_url(
        drive_id, p, routes, default_site, default_drive, strip_prefix, hostname, token
    )
    r = requests.get(item_url, headers=_auth_headers(token), timeout=30)
    _raise_for_status(r)
    item_id: str = r.json()["id"]
    return drive_id, item_id


def _build_item_url(
    drive_id: str,
    item_path: str,
    routes,
    default_site: str | None,
    default_drive: str | None,
    strip_prefix: str | None,
    hostname: str,
    token: str,
) -> str:
    """Return the Graph URL that fetches *item_path* within *drive_id*.

    Normally ``/drives/{drive_id}/root:/{path}`` (name-based).  When the path
    falls under a folder pinned via ``M365_FOLDER_IDS`` it becomes
    ``/drives/{drive_id}/items/{anchor_id}:/{remainder}`` — addressed by the
    stable anchor id, so a rename/move of the pinned folder (or any ancestor)
    does not break resolution.  Mirrors tropi_storage's ``_build_item_url``.
    """
    pins = _pins_for_drive(
        drive_id, routes, default_site, default_drive, strip_prefix, hostname, token
    )
    match = _match_pin(item_path, pins)
    if match is not None:
        anchor_id, rel = match
        if rel == "/":
            # The path IS the pinned folder itself (no sub-path).
            return f"{GRAPH}/drives/{drive_id}/items/{anchor_id}"
        enc = urllib.parse.quote(rel.lstrip("/"), safe="/")
        return f"{GRAPH}/drives/{drive_id}/items/{anchor_id}:/{enc}"

    encoded = urllib.parse.quote(item_path.lstrip("/"), safe="/")
    return f"{GRAPH}/drives/{drive_id}/root:/{encoded}"


def _match_pin(
    item_path: str, pins: list[tuple[str, str]]
) -> tuple[str, str] | None:
    """Return (anchor_item_id, relative_path) if *item_path* is at/under a pin.

    *relative_path* keeps a leading slash; it is "/" when *item_path* IS the
    pinned anchor.  Returns None when no pin matches.  The ``anchor + "/"``
    guard prevents a sibling false match (anchor ``/A/B`` must not match
    ``/A/BC/...``).
    """
    for anchor, item_id in pins:
        if item_path == anchor:
            return item_id, "/"
        if anchor != "/" and item_path.startswith(anchor + "/"):
            return item_id, item_path[len(anchor):]
    return None


def _pins_for_drive(
    drive_id: str,
    routes,
    default_site: str | None,
    default_drive: str | None,
    strip_prefix: str | None,
    hostname: str,
    token: str,
) -> list[tuple[str, str]]:
    """Return [(anchor_item_path, item_id)] for pins that live in *drive_id*.

    Each configured pin's logical path is routed to (site, library) → drive id;
    only those whose drive matches *drive_id* are kept (the same
    ``M365_FOLDER_IDS`` value is shared across services with differing route
    tables, so non-routable pins are skipped).  Resolved once per drive and
    cached.  Empty list when ``M365_FOLDER_IDS`` is unset.
    """
    with _id_cache_lock:
        cached = _pins_by_drive.get(drive_id)
    if cached is not None:
        return cached

    resolved: list[tuple[str, str]] = []
    for anchor_logical, item_id in load_folder_pins().items():
        try:
            sp, dn, anchor_item_path = resolve_route(
                anchor_logical, routes, default_site, default_drive,
                strip_prefix=strip_prefix,
            )
            if _resolve_drive_id(sp, dn, hostname, token) == drive_id:
                resolved.append((normalize_path(anchor_item_path), item_id))
        except Exception:
            continue

    with _id_cache_lock:
        _pins_by_drive[drive_id] = resolved
    return resolved


def _resolve_site_id(site_path: str, hostname: str, token: str) -> str:
    with _id_cache_lock:
        if site_path in _site_id_cache:
            return _site_id_cache[site_path]

    sp = site_path if site_path.startswith("/") else "/" + site_path
    r = requests.get(
        f"{GRAPH}/sites/{hostname}:{sp}",
        headers=_auth_headers(token),
        timeout=30,
    )
    _raise_for_status(r)
    site_id: str = r.json()["id"]
    with _id_cache_lock:
        _site_id_cache[site_path] = site_id
    return site_id


def _resolve_drive_id(
    site_path: str,
    drive_name: str | None,
    hostname: str,
    token: str,
) -> str:
    key = (site_path, drive_name)
    with _id_cache_lock:
        if key in _drive_id_cache:
            return _drive_id_cache[key]

    site_id = _resolve_site_id(site_path, hostname, token)
    h = _auth_headers(token)

    if not drive_name:
        r = requests.get(f"{GRAPH}/sites/{site_id}/drive", headers=h, timeout=30)
        _raise_for_status(r)
        drive_id: str = r.json()["id"]
    else:
        drive_id = _find_drive_by_name(site_id, drive_name, h)

    with _id_cache_lock:
        _drive_id_cache[key] = drive_id
    return drive_id


def _find_drive_by_name(site_id: str, drive_name: str, headers: dict) -> str:
    url = f"{GRAPH}/sites/{site_id}/drives"
    while url:
        r = requests.get(url, headers=headers, timeout=30)
        _raise_for_status(r)
        data = r.json()
        for entry in data.get("value", []):
            if entry.get("name") == drive_name:
                return entry["id"]
        url = data.get("@odata.nextLink", "")
    raise ExcelResolveError(
        f"Library {drive_name!r} not found in site_id={site_id!r}."
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _auth_headers(token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def _raise_for_status(resp: requests.Response) -> None:
    if resp.ok:
        return
    raise ExcelApiError(resp.status_code, resp.text)
