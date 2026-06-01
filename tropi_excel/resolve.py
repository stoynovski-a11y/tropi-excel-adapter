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

Both strategies cache results in a module-level dict keyed by
(drive_id, item_id) so repeated calls to the same file are free.
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

# Module-level resolution cache: key → (drive_id, item_id)
_cache: dict[str, tuple[str, str]] = {}
_cache_lock = threading.Lock()

# Module-level site-id / drive-id caches (same pattern as graph_backend.py)
_site_id_cache: dict[str, str] = {}
_drive_id_cache: dict[tuple[str, str | None], str] = {}
_id_cache_lock = threading.Lock()


@dataclass(frozen=True)
class ExcelFile:
    """Resolved workbook reference: drive ID + item ID."""

    drive_id: str
    item_id: str

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
            return ExcelFile(drive_id=drive_id, item_id=item_id)

    token = token_provider.get_token()

    if key.lower().startswith("http"):
        drive_id, item_id = _resolve_from_url(key, token)
    else:
        drive_id, item_id = _resolve_from_path(key, token)

    with _cache_lock:
        _cache[key] = (drive_id, item_id)
    return ExcelFile(drive_id=drive_id, item_id=item_id)


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
    try:
        from tropi_storage.routing import load_routes, load_strip_prefix, resolve_route
        from tropi_storage.path_utils import normalize_path
    except ImportError as exc:
        raise ExcelResolveError(
            "tropi-storage-adapter is not installed; "
            "install it or pass a SharePoint URL instead of a logical path."
        ) from exc

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
        from tropi_storage.exceptions import BackendError
        site_path, drive_name, item_path = resolve_route(
            path, routes, default_site, default_drive, strip_prefix=strip_prefix
        )
    except Exception as exc:
        raise ExcelResolveError(
            f"Cannot route path {path!r}: {exc}"
        ) from exc

    drive_id = _resolve_drive_id(site_path, drive_name, hostname, token)

    # Fetch the item by path under the drive.
    p = normalize_path(item_path)
    if p == "/":
        raise ExcelResolveError(
            f"Path {path!r} resolved to the library root — specify a file path."
        )
    encoded = urllib.parse.quote(p.lstrip("/"), safe="/")
    r = requests.get(
        f"{GRAPH}/drives/{drive_id}/root:/{encoded}",
        headers=_auth_headers(token),
        timeout=30,
    )
    _raise_for_status(r)
    item_id: str = r.json()["id"]
    return drive_id, item_id


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
