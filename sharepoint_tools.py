"""SharePoint tool implementations — READ-ONLY.

Opt-in: registered only when M365_ENABLE_SHAREPOINT=1. Scope: Sites.Read.All
(delegated — the user only ever sees sites they already have access to).
No write tools on purpose: a document library is shared team state, and the
OneDrive tools already cover writing to the user's own files.

Navigation model: site -> drive (document library) -> folder -> file.
Lists (SharePoint lists, not libraries) are read via sharepoint_list_items.
"""

from __future__ import annotations

from urllib.parse import urlparse

import graph

_SELECT = "id,name,file,folder,size,lastModifiedDateTime,parentReference,webUrl"


def _clamp(limit: int, hi: int = 200) -> int:
    return min(max(limit, 1), hi)


def _site(s: dict) -> dict:
    return {
        "id": s.get("id"),
        "name": s.get("displayName") or s.get("name"),
        "description": s.get("description"),
        "link": s.get("webUrl"),
    }


def _item(item: dict) -> dict:
    return {
        "id": item.get("id"),
        "name": item.get("name"),
        "mime": (item.get("file") or {}).get("mimeType") or ("folder" if "folder" in item else None),
        "is_folder": "folder" in item,
        "modified": item.get("lastModifiedDateTime"),
        "size": item.get("size"),
        "drive_id": (item.get("parentReference") or {}).get("driveId"),
        "parent_id": (item.get("parentReference") or {}).get("id"),
        "link": item.get("webUrl"),
    }


def _item_path(drive_id: str, item_id: str) -> str:
    base = f"/drives/{drive_id}"
    return f"{base}/root" if item_id in ("root", "", None) else f"{base}/items/{item_id}"


# --- Sites ---------------------------------------------------------------------


def search_sites(query: str, account: str | None = None, limit: int = 20) -> list[dict]:
    """Find sites by keyword. query='*' lists every site the user can see."""
    sites = graph.get_all("/sites", account=account, params={"search": query},
                          limit=_clamp(limit))
    return [_site(s) for s in sites]


def get_site(url: str, account: str | None = None) -> dict:
    """Resolve a SharePoint URL (as pasted from the browser) to a site.
    e.g. https://contoso.sharepoint.com/sites/Marketing"""
    parsed = urlparse(url if "://" in url else f"https://{url}")
    if not parsed.hostname or not parsed.hostname.endswith(".sharepoint.com"):
        raise ValueError("Expected a *.sharepoint.com URL")
    parts = [p for p in parsed.path.split("/") if p]
    # Keep only /sites/<name> or /teams/<name>; drop library/page segments.
    if len(parts) >= 2 and parts[0].lower() in ("sites", "teams"):
        path = f"/sites/{parsed.hostname}:/{parts[0]}/{parts[1]}"
    else:
        path = f"/sites/{parsed.hostname}"  # tenant root site
    return _site(graph.request("GET", path, account=account))


def list_drives(site_id: str, account: str | None = None) -> list[dict]:
    """Document libraries of a site ('Documents', 'Shared Documents', …)."""
    drives = graph.get_all(f"/sites/{site_id}/drives", account=account, limit=100)
    return [{"id": d.get("id"), "name": d.get("name"), "type": d.get("driveType"),
             "link": d.get("webUrl")} for d in drives]


# --- Files ---------------------------------------------------------------------


def list_folder(
    drive_id: str,
    folder_id: str = "root",
    account: str | None = None,
    limit: int = 50,
) -> list[dict]:
    items = graph.get_all(
        f"{_item_path(drive_id, folder_id)}/children", account=account,
        params={"$select": _SELECT, "$top": _clamp(limit)}, limit=_clamp(limit),
    )
    return [_item(i) for i in items]


def search_files(
    drive_id: str,
    query: str,
    account: str | None = None,
    limit: int = 10,
) -> list[dict]:
    """Search one document library (filenames + content)."""
    safe = query.replace("'", "''")
    items = graph.get_all(
        f"/drives/{drive_id}/root/search(q='{safe}')", account=account,
        params={"$select": _SELECT, "$top": _clamp(limit, 50)}, limit=_clamp(limit, 50),
    )
    return [_item(i) for i in items]


def read_file(
    drive_id: str,
    item_id: str,
    account: str | None = None,
    include_content: bool = False,
    max_chars: int = 50_000,
) -> dict:
    """Metadata by default; content is opt-in (text decodes to UTF-8,
    binaries return a size-only placeholder)."""
    path = _item_path(drive_id, item_id)
    meta = graph.request("GET", path, account=account, params={"$select": _SELECT})
    if not include_content or "folder" in meta:
        return _item(meta)
    raw = graph.request("GET", f"{path}/content", account=account, raw=True)
    try:
        text = raw.decode("utf-8")
        truncated = len(text) > max_chars
        text = text[:max_chars]
    except UnicodeDecodeError:
        text = f"(binary, {len(raw)} bytes — open the link or convert first)"
        truncated = False
    return {**_item(meta), "content": text, "truncated": truncated, "bytes": len(raw)}


# --- Lists ---------------------------------------------------------------------


def list_lists(site_id: str, account: str | None = None) -> list[dict]:
    """SharePoint lists of a site. Hidden/system lists are skipped."""
    lists = graph.get_all(
        f"/sites/{site_id}/lists", account=account,
        params={"$select": "id,displayName,description,webUrl,list"}, limit=200,
    )
    out = []
    for lst in lists:
        info = lst.get("list") or {}
        if info.get("hidden"):
            continue
        out.append({"id": lst.get("id"), "name": lst.get("displayName"),
                    "template": info.get("template"),
                    "description": lst.get("description"), "link": lst.get("webUrl")})
    return out


def list_items(
    site_id: str,
    list_id: str,
    account: str | None = None,
    limit: int = 50,
) -> list[dict]:
    """Rows of a SharePoint list, with their column values."""
    items = graph.get_all(
        f"/sites/{site_id}/lists/{list_id}/items", account=account,
        params={"$expand": "fields", "$top": _clamp(limit)}, limit=_clamp(limit),
    )
    out = []
    for it in items:
        fields = {k: v for k, v in (it.get("fields") or {}).items()
                  if not k.startswith(("@", "_")) and k not in ("id",)}
        out.append({"id": it.get("id"), "modified": it.get("lastModifiedDateTime"),
                    "link": it.get("webUrl"), "fields": fields})
    return out
