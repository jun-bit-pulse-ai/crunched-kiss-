from types import SimpleNamespace

from app.agent import SYSTEM_PROMPT, run_turn
from app.models import ChatRequest


class FakeMessages:
    def __init__(self, response: object) -> None:
        self.response = response
        self.calls: list[dict] = []

    def create(self, **kwargs):
        self.calls.append(kwargs)
        return self.response


class FakeAnthropic:
    def __init__(self, response: object) -> None:
        self.messages = FakeMessages(response)


def _text_response(text: str) -> SimpleNamespace:
    return SimpleNamespace(
        stop_reason="end_turn",
        content=[SimpleNamespace(type="text", text=text)],
    )


def _tool_response(tool_id: str, name: str, tool_input: dict) -> SimpleNamespace:
    return SimpleNamespace(
        stop_reason="tool_use",
        content=[
            SimpleNamespace(type="tool_use", id=tool_id, name=name, input=tool_input),
        ],
    )


def test_system_prompt_asks_for_formulas_on_small_error_checks() -> None:
    assert "values and formulas" in SYSTEM_PROMPT


def test_text_response_maps_to_message() -> None:
    client = FakeAnthropic(_text_response("A1 is 42"))
    result = run_turn(
        ChatRequest(messages=[{"role": "user", "content": "What is in A1?"}]),
        client=client,
    )
    assert result.type == "message"
    assert result.text == "A1 is 42"


def test_tool_use_maps_to_tool_calls() -> None:
    client = FakeAnthropic(_tool_response("toolu_1", "read_range", {"sheet": "Sheet1", "address": "A1"}))
    result = run_turn(
        ChatRequest(messages=[{"role": "user", "content": "Read A1"}]),
        client=client,
    )
    assert result.type == "tool_calls"
    assert result.tool_calls is not None
    assert result.tool_calls[0].id == "toolu_1"
    assert result.tool_calls[0].name == "read_range"
    assert result.tool_calls[0].input == {"sheet": "Sheet1", "address": "A1"}


def test_force_text_omits_tools() -> None:
    client = FakeAnthropic(_text_response("Stopping here."))
    run_turn(
        ChatRequest(
            messages=[{"role": "user", "content": "wrap up"}],
            force_text=True,
        ),
        client=client,
    )
    assert "tools" not in client.messages.calls[0]


def test_workbook_hint_control_characters_are_stripped() -> None:
    client = FakeAnthropic(_text_response("ok"))
    run_turn(
        ChatRequest(
            messages=[{"role": "user", "content": "hi"}],
            workbook_hint=["Budget", "Data\ninject"],
        ),
        client=client,
    )
    system = client.messages.calls[0]["system"]
    assert "Known sheet names: Budget, Datainject" in system
    assert "\ninject" not in system


def test_unknown_tool_names_are_dropped() -> None:
    client = FakeAnthropic(_tool_response("toolu_x", "delete_workbook", {}))
    result = run_turn(
        ChatRequest(messages=[{"role": "user", "content": "wipe it"}]),
        client=client,
    )
    assert result.type == "message"


def test_missing_api_key_returns_error() -> None:
    result = run_turn(
        ChatRequest(messages=[{"role": "user", "content": "hi"}]),
        client=None,
        api_key="",
    )
    assert result.type == "error"
    assert result.message is not None
