from app.tools import MAX_READ_CELLS, TOOLS, tool_names


def test_max_read_cells_is_2000() -> None:
    assert MAX_READ_CELLS == 2000


def test_closed_tool_allowlist() -> None:
    assert tool_names() == [
        "list_workbook_meta",
        "read_range",
        "write_range",
        "get_selection",
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
