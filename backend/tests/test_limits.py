from app.limits import (
    MAX_BODY_BYTES,
    MAX_HINT_CHARS,
    MAX_WORKBOOK_HINTS,
    message_payload_size,
    reject_content_length,
    sanitize_workbook_hints,
)


def test_message_payload_size_counts_text() -> None:
    assert message_payload_size("hello") == 5


def test_message_payload_size_counts_json_blocks() -> None:
    assert message_payload_size([{"type": "tool_result", "content": "x"}]) > 1


def test_reject_content_length() -> None:
    assert reject_content_length(None) is None
    assert reject_content_length("12") is None
    assert reject_content_length(str(MAX_BODY_BYTES)) is None
    assert reject_content_length(str(MAX_BODY_BYTES + 1)) == 413
    assert reject_content_length("nope") == 400


def test_sanitize_workbook_hints_strips_control_and_caps() -> None:
    assert sanitize_workbook_hints(None) == []
    assert sanitize_workbook_hints(["Budget", "  Data\n\ninject  "]) == ["Budget", "Datainject"]
    long_name = "A" * (MAX_HINT_CHARS + 20)
    assert sanitize_workbook_hints([long_name]) == ["A" * MAX_HINT_CHARS]
    many = [f"S{i}" for i in range(MAX_WORKBOOK_HINTS + 5)]
    assert len(sanitize_workbook_hints(many)) == MAX_WORKBOOK_HINTS
