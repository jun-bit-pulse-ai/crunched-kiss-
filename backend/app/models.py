from typing import Any, Literal

from pydantic import BaseModel, Field


class ChatMessage(BaseModel):
    role: Literal["user", "assistant"]
    content: str | list[dict[str, Any]]


class ToolCall(BaseModel):
    id: str
    name: str
    input: dict[str, Any] = Field(default_factory=dict)


class ChatRequest(BaseModel):
    messages: list[ChatMessage]
    workbook_hint: list[str] | None = None
    force_text: bool = False


class ChatResponse(BaseModel):
    type: Literal["tool_calls", "message", "error"]
    tool_calls: list[ToolCall] | None = None
    text: str | None = None
    message: str | None = None
