MAX_READ_CELLS = 2000

TOOLS: list[dict] = [
    {
        "name": "list_workbook_meta",
        "description": (
            "List worksheets with used-range addresses and dimensions. "
            "Call this first on an unfamiliar workbook. This is O(sheets), not O(cells)."
        ),
        "input_schema": {
            "type": "object",
            "properties": {},
            "additionalProperties": False,
        },
    },
    {
        "name": "read_range",
        "description": (
            f"Read values and formulas from a range. Reads over {MAX_READ_CELLS} cells "
            "are truncated by Excel; never request an entire huge used range."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet": {"type": "string"},
                "address": {
                    "type": "string",
                    "description": "A1-style address such as A1 or A1:D20",
                },
            },
            "required": ["sheet", "address"],
            "additionalProperties": False,
        },
    },
    {
        "name": "write_range",
        "description": "Write a 2D values array starting at the given address.",
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet": {"type": "string"},
                "address": {"type": "string"},
                "values": {
                    "type": "array",
                    "items": {"type": "array"},
                    "description": "Row-major 2D array",
                },
            },
            "required": ["sheet", "address", "values"],
            "additionalProperties": False,
        },
    },
    {
        "name": "get_selection",
        "description": "Return the user's current selection address and a small preview.",
        "input_schema": {
            "type": "object",
            "properties": {},
            "additionalProperties": False,
        },
    },
]


def tool_names() -> list[str]:
    return [tool["name"] for tool in TOOLS]
