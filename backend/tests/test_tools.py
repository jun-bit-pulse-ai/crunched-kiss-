from app.tools import MAX_FIND_RESULTS, MAX_READ_CELLS, TOOLS, tool_names


def test_max_read_cells_is_2000() -> None:
    assert MAX_READ_CELLS == 2000


def test_closed_tool_allowlist() -> None:
    assert tool_names() == [
        "list_workbook_meta",
        "read_range",
        "write_range",
        "get_selection",
        "find",
    ]


def test_read_range_schema_requires_sheet_and_address() -> None:
    read_range = next(tool for tool in TOOLS if tool["name"] == "read_range")
    required = read_range["input_schema"]["required"]
    assert "sheet" in required
    assert "address" in required


def test_write_range_schema_requires_values() -> None:
    write_range = next(tool for tool in TOOLS if tool["name"] == "write_range")
    assert "values" in write_range["input_schema"]["required"]
    assert "values" in write_range["input_schema"]["properties"]


def test_find_schema_requires_query_only() -> None:
    find = next(tool for tool in TOOLS if tool["name"] == "find")
    schema = find["input_schema"]
    assert schema["required"] == ["query"]
    assert set(schema["properties"]) == {"query", "sheet", "match_case", "complete_match"}
    assert schema["additionalProperties"] is False


def test_find_description_states_the_result_cap() -> None:
    find = next(tool for tool in TOOLS if tool["name"] == "find")
    assert str(MAX_FIND_RESULTS) in find["description"]


def test_list_workbook_meta_mentions_the_header_preview() -> None:
    meta = next(tool for tool in TOOLS if tool["name"] == "list_workbook_meta")
    assert "header" in meta["description"].lower()
