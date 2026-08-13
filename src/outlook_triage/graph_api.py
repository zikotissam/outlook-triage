from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone
from typing import Any, Callable
from urllib.parse import urlencode

import requests

from .models import Attachment, Message

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

WELL_KNOWN_FOLDERS = {
    "inbox": "inbox",
    "spam": "junkemail",
    "junk": "junkemail",
    "junkemail": "junkemail",
    "sent": "sentitems",
    "sentitems": "sentitems",
    "draft": "drafts",
    "drafts": "drafts",
    "trash": "deleteditems",
    "deleteditems": "deleteditems",
    "archive": "archive",
    "outbox": "outbox",
    "conversationhistory": "conversationhistory",
    "searchfolders": "searchfolders",
}

SELECT_FIELDS = [
    "id",
    "conversationId",
    "from",
    "toRecipients",
    "subject",
    "bodyPreview",
    "receivedDateTime",
    "isRead",
    "isDraft",
    "importance",
    "inferenceClassification",
    "categories",
    "hasAttachments",
    "parentFolderId",
    "internetMessageHeaders",
]

META_HEADERS = ["List-Unsubscribe", "Precedence", "X-Auto-Response-Suppress", "X-Feedback-ID", "X-Mailer", "Auto-Submitted"]


def _resolve_window(since: datetime | None, until: datetime | None, hours: int | None) -> tuple[datetime | None, datetime | None]:
    if hours is not None:
        if since is not None or until is not None:
            raise ValueError("--hours cannot be combined with --since/--until")
        since = datetime.now(timezone.utc) - timedelta(hours=hours)
    return since, until


def _utc_midnight(dt: datetime) -> datetime:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def build_filter(since: datetime | None = None, until: datetime | None = None, hours: int | None = None, unread_only: bool = False, extra: str | None = None) -> str | None:
    since, until = _resolve_window(since, until, hours)
    parts: list[str] = []
    if since:
        parts.append(f"receivedDateTime ge {_utc_midnight(since).isoformat()}")
    if until:
        parts.append(f"receivedDateTime lt {_utc_midnight(until).isoformat()}")
    if unread_only:
        parts.append("isRead eq false")
    if extra and extra.strip():
        parts.append(extra.strip())
    if not parts:
        return None
    return " and ".join(parts)


class GraphClient:
    def __init__(self, get_token: Callable[[], str]):
        self.get_token = get_token

    def _request(self, url: str, params: dict[str, str] | None = None, headers: dict[str, str] | None = None) -> dict:
        req_headers = {"Authorization": f"Bearer {self.get_token()}"}
        if headers:
            req_headers.update(headers)
        resp = requests.get(url, params=params, headers=req_headers, timeout=30)
        if resp.status_code == 401:
            raise SystemExit("Authentication expired. Re-run `outlook-triage auth-run` for this account.")
        if resp.status_code == 403:
            raise SystemExit(f"Access denied (403). Check the Mail.Read permission is granted: {resp.text[:300]}")
        if resp.status_code != 200:
            raise SystemExit(f"Graph API error {resp.status_code}: {resp.text[:500]}")
        return resp.json()

    def _mail_folder_id(self, folder: str) -> str | None:
        key = folder.strip().lower()
        if key in WELL_KNOWN_FOLDERS:
            return WELL_KNOWN_FOLDERS[key]
        data = self._request(
            f"{GRAPH_BASE}/me/mailFolders",
            params={"$filter": f"displayName eq '{folder.strip()}'", "$select": "id,displayName"},
        )
        for f in data.get("value", []):
            if f.get("displayName", "").lower() == key:
                return f["id"]
        return None

    def fetch(
        self,
        since: datetime | None = None,
        until: datetime | None = None,
        hours: int | None = None,
        max_results: int = 100,
        include_body: bool = False,
        folder: str | None = None,
        unread_only: bool = False,
        extra: str | None = None,
    ) -> list[Message]:
        filter_str = build_filter(since, until, hours, unread_only=unread_only, extra=extra)

        pseudo = folder.strip().lower() if folder else ""
        if pseudo in ("focused", "other"):
            folder_id = None
            cls_filter = f"inferenceClassification eq '{pseudo}'"
            filter_str = " and ".join(x for x in [filter_str, cls_filter] if x)
        elif pseudo in ("all", "allitems", "allmail", "everything"):
            folder_id = None
        else:
            folder_id = self._mail_folder_id(folder) if folder else "inbox"

        if folder_id:
            url = f"{GRAPH_BASE}/me/mailFolders/{folder_id}/messages"
        else:
            url = f"{GRAPH_BASE}/me/messages"

        params: dict[str, str] = {}
        params["$select"] = ",".join(SELECT_FIELDS)
        if include_body:
            params["$select"] += ",body"
        if filter_str:
            params["$filter"] = filter_str
        params["$orderby"] = "receivedDateTime desc"
        params["$top"] = str(max_results)
        params["$expand"] = "attachments($select=name,contentType,size)"

        req_headers = {}
        if include_body:
            req_headers["Prefer"] = "outlook.body-content-type: text"

        messages: list[Message] = []
        next_url: str | None = url + "?" + urlencode(params)
        while next_url and len(messages) < max_results:
            data = self._request(next_url, headers=req_headers)
            for entry in data.get("value", []):
                messages.append(_parse_message(entry, include_body=include_body))
            next_url = _next_link(data)
        return messages


def _next_link(data: dict) -> str | None:
    link = data.get("@odata.nextLink")
    if not link:
        return None
    # Cap at the requested limit to honor --limit on subsequent pages.
    return link


def _parse_message(entry: dict, include_body: bool = False) -> Message:
    headers: dict[str, str] = {}
    for h in entry.get("internetMessageHeaders", []) or []:
        name = (h.get("name") or "").lower()
        if name:
            headers[name] = h.get("value", "")

    def head(name: str) -> str:
        return headers.get(name.lower(), "")

    from_addr = _parse_recipient(entry.get("from") or entry.get("sender"))
    to = ", ".join(_parse_recipient(r) for r in (entry.get("toRecipients") or []))
    date = _parse_datetime(entry.get("receivedDateTime"))
    body = ""
    if include_body:
        body = (entry.get("body") or {}).get("content", "") or ""

    attachments = [_parse_attachment(a) for a in (entry.get("attachments") or [])]

    return Message(
        id=entry.get("id", ""),
        conversation_id=entry.get("conversationId") or entry.get("id", ""),
        from_=from_addr,
        from_name=_name_part(entry.get("from") or entry.get("sender")),
        to=to,
        subject=entry.get("subject", "") or "",
        snippet=entry.get("bodyPreview", "") or "",
        date=date,
        importance=entry.get("importance", "normal") or "normal",
        inference_classification=entry.get("inferenceClassification"),
        categories=list(entry.get("categories") or []),
        is_read=bool(entry.get("isRead", True)),
        is_draft=bool(entry.get("isDraft", False)),
        headers=headers,
        body=body,
        has_attachments=bool(entry.get("hasAttachments", False)),
        attachments=attachments,
        unsubscribe_url=_parse_unsubscribe(headers.get("list-unsubscribe")),
    )


def _parse_recipient(r: dict | None) -> str:
    if not r:
        return ""
    ea = r.get("emailAddress") or {}
    name = ea.get("name")
    address = ea.get("address", "")
    if name and name != address:
        return f"{name} <{address}>"
    return address


def _name_part(r: dict | None) -> str:
    if not r:
        return ""
    ea = r.get("emailAddress") or {}
    name = ea.get("name") or ""
    if name:
        return name
    return ea.get("address", "")


def _parse_attachment(a: dict) -> Attachment:
    return Attachment(
        name=a.get("name", ""),
        content_type=a.get("contentType", ""),
        size=a.get("size"),
    )


def _parse_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def _parse_unsubscribe(header: str | None) -> str | None:
    if not header:
        return None
    for match in re.finditer(r"<(https?://[^>]+)>", header):
        return match.group(1)
    for match in re.finditer(r"<mailto:[^>]+>", header):
        return match.group(0)
    return None
