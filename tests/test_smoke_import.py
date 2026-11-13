"""Server-import smoke test.

Imports the FastMCP server entry module and asserts it constructs. This is the
runtime-verification gate the install step alone does not provide: a dependency
bump (e.g. a fastmcp / msal major) that breaks tool registration or server
construction fails here instead of shipping silently.
"""
from __future__ import annotations

import asyncio
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def test_server_module_constructs() -> None:
    from fastmcp import FastMCP
    import server
    assert isinstance(server.mcp, FastMCP), f"server.mcp is not FastMCP: {type(server.mcp)!r}"


def test_all_tools_registered_with_schemas() -> None:
    import server
    tools = asyncio.run(server.mcp.list_tools())
    names = {t.name for t in tools}
    assert len(names) == 35, f"expected 35 tools, got {len(names)}"
    for family, expect in [("m365_account_", 4), ("outlook_", 11), ("mscal_", 7), ("onedrive_", 13)]:
        got = sum(1 for n in names if n.startswith(family))
        assert got == expect, f"{family}*: expected {expect}, got {got}"
    # The observability wrapper must not erase real parameter schemas.
    search = next(t for t in tools if t.name == "outlook_search")
    assert "query" in search.parameters["properties"]
    assert "kwargs" not in search.parameters["properties"]


def _tool_names(env: dict) -> set[str]:
    """Import the server in a clean subprocess with the given flags."""
    import json, os, subprocess
    code = ("import asyncio, json, server; "
            "print(json.dumps(sorted(t.name for t in asyncio.run(server.mcp.list_tools()))))")
    base = {k: v for k, v in os.environ.items() if not k.startswith("M365_ENABLE_")}
    out = subprocess.run([sys.executable, "-c", code], cwd=ROOT, env={**base, **env},
                         capture_output=True, text=True, check=True)
    return set(json.loads(out.stdout.strip().splitlines()[-1]))


def test_optional_families_hidden_by_default() -> None:
    names = _tool_names({})
    assert not any(n.startswith(("teams_", "sharepoint_")) for n in names)


def test_optional_families_registered_when_enabled() -> None:
    names = _tool_names({"M365_ENABLE_TEAMS": "1", "M365_ENABLE_SHAREPOINT": "1"})
    assert len(names) == 35 + 6 + 8
    assert "teams_read_channel" not in names  # needs its own admin-consent flag
    names = _tool_names({"M365_ENABLE_TEAMS": "1", "M365_ENABLE_TEAMS_CHANNEL_READ": "1"})
    assert "teams_read_channel" in names
