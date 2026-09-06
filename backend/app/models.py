from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from app.limits import (
    MAX_CHAT_MESSAGES,
    MAX_HINT_CHARS,
    MAX_MESSAGE_CHARS,
    MAX_WORKBOOK_HINTS,
    message_payload_size,
)


class ChatMessage(BaseModel):
    role: Literal["user", "assistant"]
    content: str | list[dict[str, Any]]

    @field_validator("content")
    @classmethod
    def content_fits_budget(cls, value: str | list[dict[str, Any]]) -> str | list[dict[str, Any]]:
        if message_payload_size(value) > MAX_MESSAGE_CHARS:
            raise ValueError(f"message content exceeds {MAX_MESSAGE_CHARS} characters")
        return value


class ToolCall(BaseModel):
    id: str
    name: str
    input: dict[str, Any] = Field(default_factory=dict)


class ChatRequest(BaseModel):
    messages: list[ChatMessage] = Field(min_length=1, max_length=MAX_CHAT_MESSAGES)
    workbook_hint: list[str] | None = Field(default=None, max_length=MAX_WORKBOOK_HINTS)
    force_text: bool = False

    @field_validator("workbook_hint")
    @classmethod
    def hint_strings_are_short(cls, value: list[str] | None) -> list[str] | None:
        if value is None:
            return value
        for hint in value:
            if len(hint) > MAX_HINT_CHARS:
                raise ValueError(f"workbook hint exceeds {MAX_HINT_CHARS} characters")
        return value


class ChatResponse(BaseModel):
    type: Literal["tool_calls", "message", "error"]
    tool_calls: list[ToolCall] | None = None
    text: str | None = None
    message: str | None = None
