"""SharePoint tools against the fake Graph transport."""

from __future__ import annotations

import pytest

import sharepoint_tools

FILE = {
    "id": "I1", "name": "acta.txt", "file": {"mimeType": "text/plain"}, "size": 5,
    "lastModifiedDateTime": "2026-11-01T00:00:00Z",
    "parentReference": {"driveId": "D1", "id": "P1"}, "webUrl": "https://sp.example/acta",
}


@pytest.mark.parametrize("url,expected", [
    ("https://contoso.sharepoint.com/sites/Marketing/Shared%20Documents/Forms/AllItems.aspx",
     "/sites/contoso.sharepoint.com:/sites/Marketing"),
    ("contoso.sharepoint.com/teams/Ventas", "/sites/contoso.sharepoint.com:/teams/Ventas"),
    ("https://contoso.sharepoint.com/", "/sites/contoso.sharepoint.com"),
])
def test_get_site_parses_browser_urls(fake_graph, url, expected):
    fake_graph.queue("GET", expected, {"id": "S1", "displayName": "Marketing"})
    out = sharepoint_tools.get_site(url)
    assert fake_graph.calls[0]["path"] == expected
    assert out["id"] == "S1"


def test_get_site_rejects_non_sharepoint(fake_graph):
    with pytest.raises(ValueError):
        sharepoint_tools.get_site("https://example.com/sites/x")


def test_list_folder_root_vs_item(fake_graph):
    sharepoint_tools.list_folder("D1")
    sharepoint_tools.list_folder("D1", folder_id="F9")
    assert fake_graph.calls[0]["path"] == "/drives/D1/root/children"
    assert fake_graph.calls[1]["path"] == "/drives/D1/items/F9/children"


def test_search_files_escapes_quotes(fake_graph):
    fake_graph.queue("GET", "/drives/D1/root/search(q='o''brien')", {"value": [FILE]})
    out = sharepoint_tools.search_files("D1", "o'brien")
    assert out[0]["drive_id"] == "D1"


def test_read_file_content_is_opt_in(fake_graph):
    fake_graph.queue("GET", "/drives/D1/items/I1", FILE)
    assert "content" not in sharepoint_tools.read_file("D1", "I1")
    fake_graph.queue("GET", "/drives/D1/items/I1", FILE)
    fake_graph.queue("GET", "/drives/D1/items/I1/content", b"hola mundo")
    out = sharepoint_tools.read_file("D1", "I1", include_content=True, max_chars=4)
    assert out["content"] == "hola" and out["truncated"] is True


def test_list_lists_skips_hidden(fake_graph):
    fake_graph.queue("GET", "/sites/S1/lists", {"value": [
        {"id": "L1", "displayName": "Tareas", "list": {"template": "genericList", "hidden": False}},
        {"id": "L2", "displayName": "Sistema", "list": {"hidden": True}},
    ]})
    assert [x["id"] for x in sharepoint_tools.list_lists("S1")] == ["L1"]


def test_list_items_drops_odata_and_system_fields(fake_graph):
    fake_graph.queue("GET", "/sites/S1/lists/L1/items", {"value": [
        {"id": "1", "fields": {"@odata.etag": "x", "_UIVersionString": "1.0",
                               "id": "1", "Title": "Llamar a Ana", "Estado": "Abierta"}},
    ]})
    out = sharepoint_tools.list_items("S1", "L1")
    assert out[0]["fields"] == {"Title": "Llamar a Ana", "Estado": "Abierta"}
