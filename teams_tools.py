"""Microsoft Teams tool implementations — chats and channels.

Opt-in: these tools are registered only when M365_ENABLE_TEAMS=1 (see
accounts.active_scopes). Reading CHANNEL messages needs ChannelMessage.Read.All,
which requires tenant-admin consent, so it sits behind a second flag
(M365_ENABLE_TEAMS_CHANNEL_READ=1). Everything else works with user consent.

Design rules (same as the mail tools):
- List tools return compact shapes; message bodies come back as plain text.
- System events (member added, call ended…) are dropped from reads.
- Sends are irreversible: every send takes dry_run.
"""

from __future__ import annotations

import html
import re

import graph

_TAG = re.compile(r"<[^>]+>")
_BLANKS = re.compile(r"\n{3,}")


def _text(body: dict | None) -> str:
    """Graph returns Teams bodies as HTML; flatten to plain text."""
    body = body or {}
    content = body.get("content") or ""
    if body.get("contentType") == "html":
        content = re.sub(r"<br\s*/?>|</p>|</div>", "\n", content, flags=re.I)
        content = html.unescape(_TAG.sub("", content))
    return _BLANKS.sub("\n\n", content).strip()


def _sender(msg: dict) -> str | None:
    frm = msg.get("from") or {}
    who = frm.get("user") or frm.get("application") or {}
    return who.get("displayName")


def _message(msg: dict) -> dict:
    return {
        "id": msg.get("id"),
        "from": _sender(msg),
        "created": msg.get("createdDateTime"),
        "text": _text(msg.get("body")),
        "reply_to_id": msg.get("replyToId"),
        "link": msg.get("webUrl"),
    }


def _messages(raw: list[dict], limit: int) -> list[dict]:
    real = [m for m in raw if m.get("messageType") == "message" and not m.get("deletedDateTime")]
    return [_message(m) for m in real[:limit]]


def _clamp(limit: int, hi: int = 50) -> int:
    return min(max(limit, 1), hi)


def _body(text: str) -> dict:
    return {"body": {"contentType": "text", "content": text}}


# --- Chats --------------------------------------------------------------------


def list_chats(account: str | None = None, limit: int = 20) -> list[dict]:
    """The user's 1:1, group and meeting chats, most recently active first."""
    chats = graph.get_all(
        "/me/chats", account=account,
        params={"$expand": "members", "$top": 50},
        limit=_clamp(limit, 200),
    )
    out = []
    for c in chats:
        members = [m.get("displayName") for m in c.get("members", []) if m.get("displayName")]
        out.append({
            "id": c.get("id"),
            "type": c.get("chatType"),  # oneOnOne | group | meeting
            "topic": c.get("topic"),
            "members": members,
            "last_updated": c.get("lastUpdatedDateTime"),
            "link": c.get("webUrl"),
        })
    out.sort(key=lambda c: c["last_updated"] or "", reverse=True)
    return out


def read_chat(chat_id: str, account: str | None = None, limit: int = 20) -> list[dict]:
    """Latest messages in a chat (newest first), as plain text."""
    raw = graph.get_all(
        f"/chats/{chat_id}/messages", account=account,
        params={"$top": _clamp(limit)}, limit=_clamp(limit),
    )
    return _messages(raw, _clamp(limit))


def send_chat_message(
    chat_id: str,
    text: str,
    account: str | None = None,
    dry_run: bool = False,
) -> dict:
    """Send a message in a chat. DESTRUCTIVE — irreversible once sent."""
    if not text.strip():
        raise ValueError("text is empty")
    if dry_run:
        return {"dry_run": True, "chat_id": chat_id, "text_preview": text[:500],
                "status": "NOT SENT — dry_run=True"}
    sent = graph.request("POST", f"/chats/{chat_id}/messages", account=account, json_body=_body(text))
    return {"chat_id": chat_id, "message_id": sent.get("id"), "status": "sent"}


# --- Teams and channels ---------------------------------------------------------


def list_teams(account: str | None = None) -> list[dict]:
    teams = graph.get_all("/me/joinedTeams", account=account, limit=200)
    return [{"id": t.get("id"), "name": t.get("displayName"),
             "description": t.get("description")} for t in teams]


def list_channels(team_id: str, account: str | None = None) -> list[dict]:
    chans = graph.get_all(f"/teams/{team_id}/channels", account=account, limit=200)
    return [{"id": c.get("id"), "name": c.get("displayName"),
             "type": c.get("membershipType"), "link": c.get("webUrl")} for c in chans]


def read_channel(
    team_id: str,
    channel_id: str,
    account: str | None = None,
    limit: int = 20,
    include_replies: bool = False,
) -> list[dict]:
    """Latest top-level posts in a channel. Needs ChannelMessage.Read.All
    (tenant-admin consent). include_replies expands each post's thread."""
    params: dict = {"$top": _clamp(limit)}
    if include_replies:
        params["$expand"] = "replies"
    raw = graph.get_all(
        f"/teams/{team_id}/channels/{channel_id}/messages", account=account,
        params=params, limit=_clamp(limit),
    )
    out = []
    for m in raw:
        if m.get("messageType") != "message" or m.get("deletedDateTime"):
            continue
        post = _message(m)
        post["subject"] = m.get("subject")
        if include_replies:
            post["replies"] = _messages(m.get("replies", []), 50)
        out.append(post)
    return out[:_clamp(limit)]


def send_channel_message(
    team_id: str,
    channel_id: str,
    text: str,
    account: str | None = None,
    reply_to_message_id: str | None = None,
    dry_run: bool = False,
) -> dict:
    """Post in a channel, or reply to a post when reply_to_message_id is set.
    DESTRUCTIVE — irreversible once sent."""
    if not text.strip():
        raise ValueError("text is empty")
    path = f"/teams/{team_id}/channels/{channel_id}/messages"
    if reply_to_message_id:
        path += f"/{reply_to_message_id}/replies"
    if dry_run:
        return {"dry_run": True, "team_id": team_id, "channel_id": channel_id,
                "reply_to_message_id": reply_to_message_id, "text_preview": text[:500],
                "status": "NOT SENT — dry_run=True"}
    sent = graph.request("POST", path, account=account, json_body=_body(text))
    return {"team_id": team_id, "channel_id": channel_id,
            "message_id": sent.get("id"), "status": "sent"}
