from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class Attachment:
    name: str
    content_type: str
    size: int | None = None


@dataclass
class Message:
    id: str
    conversation_id: str
    from_: str
    from_name: str
    to: str
    subject: str
    snippet: str
    date: datetime | None
    importance: str = "normal"
    inference_classification: str | None = None
    categories: list[str] = field(default_factory=list)
    is_read: bool = True
    is_draft: bool = False
    headers: dict[str, str] = field(default_factory=dict)
    body: str = ""
    has_attachments: bool = False
    attachments: list[Attachment] = field(default_factory=list)
    unsubscribe_url: str | None = None

    def header(self, name: str) -> str | None:
        return self.headers.get(name)

    @property
    def is_unread(self) -> bool:
        return not self.is_read

    def to_features(self, include_body: bool = False) -> dict:
        out = {
            "id": self.id,
            "conversation_id": self.conversation_id,
            "from": self.from_,
            "from_name": self.from_name,
            "to": self.to,
            "subject": self.subject,
            "snippet": self.snippet,
            "date": self.date.isoformat() if self.date else None,
            "is_unread": self.is_unread,
            "is_draft": self.is_draft,
            "importance": self.importance,
            "inference_classification": self.inference_classification,
            "categories": self.categories,
            "has_attachment": self.has_attachments,
            "attachments": [
                {
                    "name": a.name,
                    "content_type": a.content_type,
                    "size": a.size,
                }
                for a in self.attachments
            ],
            "unsubscribe_url": self.unsubscribe_url,
            "headers": {
                k: v
                for k, v in self.headers.items()
                if k
                in (
                    "list-unsubscribe",
                    "precedence",
                    "x-auto-response-suppress",
                    "x-feedback-id",
                    "x-mailer",
                    "auto-submitted",
                )
            },
        }
        if include_body:
            out["body"] = self.body
        return out
