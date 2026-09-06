from __future__ import annotations

from typing import Any

from app.config import settings
from app.models import ChatRequest, ChatResponse, ToolCall
from app.tools import MAX_READ_CELLS, TOOLS

SYSTEM_PROMPT = f"""You are Crunched, an AI analyst that lives in Excel.

You never see the workbook yourself. You request data through tools. The Excel task pane executes those tools and returns bounded results.

Tools:
- list_workbook_meta: sheets, used-range addresses, row/column counts. Use this first on an unfamiliar book.
- read_range: values and formulas for a specific address.
- write_range: write a 2D values array.
- get_selection: the user's current selection, when they say "this table" or similar.

Rules:
- Metadata first. Never ask for an entire used range on a large sheet.
- Prefer headers, a small sample, and named or mentioned ranges.
- Reads over {MAX_READ_CELLS} cells are truncated. Say so when that happens.
- If a tool errors, recover with a smaller or corrected request.
- When you have enough to answer, reply in plain language. Do not mention tool JSON.
"""


def _anthropic_messages(request: ChatRequest) -> list[dict[str, Any]]:
    return [{"role": message.role, "content": message.content} for message in request.messages]


def run_turn(
    request: ChatRequest,
    client: Any | None = None,
    api_key: str | None = None,
    model: str | None = None,
) -> ChatResponse:
    key = settings.anthropic_api_key if api_key is None else api_key
    if client is None and not key:
        return ChatResponse(type="error", message="ANTHROPIC_API_KEY is not configured on the server.")

    if client is None:
        from anthropic import Anthropic

        client = Anthropic(api_key=key)

    system = SYSTEM_PROMPT
    if request.workbook_hint:
        system += "\nKnown sheet names: " + ", ".join(request.workbook_hint)

    kwargs: dict[str, Any] = {
        "model": model or settings.anthropic_model,
        "max_tokens": 4096,
        "system": system,
        "messages": _anthropic_messages(request),
    }
    if not request.force_text:
        kwargs["tools"] = TOOLS

    response = client.messages.create(**kwargs)
    if getattr(response, "stop_reason", None) == "tool_use":
        tool_calls = [
            ToolCall(id=block.id, name=block.name, input=dict(block.input or {}))
            for block in response.content
            if getattr(block, "type", None) == "tool_use"
        ]
        if tool_calls:
            return ChatResponse(type="tool_calls", tool_calls=tool_calls)

    text_parts = [
        block.text for block in response.content if getattr(block, "type", None) == "text"
    ]
    text = "".join(text_parts).strip() or "I did not get a text reply from the model."
    return ChatResponse(type="message", text=text)
