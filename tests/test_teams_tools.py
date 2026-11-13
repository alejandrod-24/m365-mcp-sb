"""Teams tools against the fake Graph transport."""

from __future__ import annotations

import pytest

import teams_tools

MSG = {
    "id": "M1", "messageType": "message", "createdDateTime": "2026-11-12T15:00:00Z",
    "from": {"user": {"displayName": "Ana"}},
    "body": {"contentType": "html", "content": "<p>Hola &amp; bienvenidos</p><p>Segunda línea</p>"},
    "webUrl": "https://teams.example/m1",
}
SYSTEM = {"id": "S1", "messageType": "systemEventMessage", "body": {"content": ""}}
DELETED = {**MSG, "id": "D1", "deletedDateTime": "2026-11-12T16:00:00Z"}


def test_read_chat_flattens_html_and_drops_system_and_deleted(fake_graph):
    fake_graph.queue("GET", "/chats/C1/messages", {"value": [SYSTEM, MSG, DELETED]})
    out = teams_tools.read_chat("C1")
    assert len(out) == 1
    assert out[0]["from"] == "Ana"
    assert out[0]["text"] == "Hola & bienvenidos\nSegunda línea"


def test_list_chats_sorted_by_activity_with_members(fake_graph):
    fake_graph.queue("GET", "/me/chats", {"value": [
        {"id": "old", "chatType": "group", "lastUpdatedDateTime": "2026-01-01T00:00:00Z",
         "members": [{"displayName": "Ana"}, {"displayName": "Beto"}]},
        {"id": "new", "chatType": "oneOnOne", "lastUpdatedDateTime": "2026-11-01T00:00:00Z",
         "members": [{"displayName": "Carla"}]},
    ]})
    out = teams_tools.list_chats()
    assert [c["id"] for c in out] == ["new", "old"]
    assert out[1]["members"] == ["Ana", "Beto"]
    assert fake_graph.calls[0]["params"]["$expand"] == "members"


def test_send_chat_dry_run_makes_no_call(fake_graph):
    out = teams_tools.send_chat_message("C1", "hola", dry_run=True)
    assert out["status"].startswith("NOT SENT")
    assert fake_graph.calls == []


def test_send_chat_posts_text_body(fake_graph):
    fake_graph.queue("POST", "/chats/C1/messages", {"id": "M9"})
    out = teams_tools.send_chat_message("C1", "hola")
    assert out == {"chat_id": "C1", "message_id": "M9", "status": "sent"}
    assert fake_graph.calls[0]["json_body"] == {"body": {"contentType": "text", "content": "hola"}}


def test_send_rejects_empty_text(fake_graph):
    with pytest.raises(ValueError):
        teams_tools.send_chat_message("C1", "   ")


def test_channel_reply_targets_replies_endpoint(fake_graph):
    fake_graph.queue("POST", "/teams/T1/channels/CH1/messages/M1/replies", {"id": "R1"})
    out = teams_tools.send_channel_message("T1", "CH1", "ok", reply_to_message_id="M1")
    assert out["message_id"] == "R1"
    assert fake_graph.calls[0]["path"] == "/teams/T1/channels/CH1/messages/M1/replies"


def test_read_channel_expands_replies(fake_graph):
    fake_graph.queue("GET", "/teams/T1/channels/CH1/messages",
                     {"value": [{**MSG, "subject": "Plan", "replies": [MSG, SYSTEM]}]})
    out = teams_tools.read_channel("T1", "CH1", include_replies=True)
    assert out[0]["subject"] == "Plan"
    assert len(out[0]["replies"]) == 1
    assert fake_graph.calls[0]["params"]["$expand"] == "replies"
