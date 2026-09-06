from __future__ import annotations

import json
from typing import Any

# Chat history is stubbed on the pane, but a client can still POST a huge body
# straight at uvicorn. Keep the request small enough to be a single Claude turn.
MAX_CHAT_MESSAGES = 40
MAX_MESSAGE_CHARS = 200_000
MAX_WORKBOOK_HINTS = 64
MAX_HINT_CHARS = 128
MAX_BODY_BYTES = 1_048_576


def message_payload_size(content: str | list[dict[str, Any]]) -> int:
    if isinstance(content, str):
        return len(content)
    return len(json.dumps(content, ensure_ascii=False))


def reject_content_length(header: str | None) -> int | None:
    """Return 400/413 when Content-Length is invalid or over the budget."""
    if not header:
        return None
    try:
        size = int(header)
    except ValueError:
        return 400
    if size > MAX_BODY_BYTES:
        return 413
    return None


def sanitize_workbook_hints(hints: list[str] | None) -> list[str]:
    """Drop control characters so a sheet name cannot break the system prompt."""
    if not hints:
        return []
    cleaned: list[str] = []
    for raw in hints[:MAX_WORKBOOK_HINTS]:
        name = "".join(char for char in raw if char.isprintable()).strip()[:MAX_HINT_CHARS]
        if name:
            cleaned.append(name)
    return cleaned
