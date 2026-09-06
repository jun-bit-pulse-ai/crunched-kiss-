MAX_READ_CELLS = 2000
MAX_FIND_RESULTS = 50

TOOLS: list[dict] = [
    {
        "name": "list_workbook_meta",
        "description": (
            "List worksheets with used-range addresses, dimensions, and a header preview "
            "(the first row, capped at 20 columns). Call this first on an unfamiliar "
            "workbook. This is O(sheets), not O(cells)."
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
    {
        "name": "find",
        "description": (
            "Search for text and get back matching cell addresses, at most "
            f"{MAX_FIND_RESULTS}. Use this to locate a label such as 'Revenue' or 'Total' "
            "before reading, instead of scanning a large range. Searches every sheet "
            "unless one is named."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Text to look for in cell contents"},
                "sheet": {
                    "type": "string",
                    "description": "Restrict the search to one sheet. Omit to search all sheets.",
                },
                "match_case": {"type": "boolean", "description": "Case-sensitive search"},
                "complete_match": {
                    "type": "boolean",
                    "description": "Match the whole cell rather than a substring",
                },
            },
            "required": ["query"],
            "additionalProperties": False,
        },
    },
]


def tool_names() -> list[str]:
    return [tool["name"] for tool in TOOLS]
