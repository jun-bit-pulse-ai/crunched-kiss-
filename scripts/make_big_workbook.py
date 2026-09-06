"""Build scripts/big.xlsx: a 1,000,000-cell Data sheet plus a small Budget sheet with
deliberate mistakes, to show the agent inspects big workbooks without reading them whole."""

from pathlib import Path
from openpyxl import Workbook

# Resolve output path from the script's own location, so it works from anywhere.
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_PATH = SCRIPT_DIR / "big.xlsx"

ROWS, COLS = 5000, 200

wb = Workbook(write_only=True)

# Sheet 1: Data — 1,000,000 numeric cells
data = wb.create_sheet("Data")
data.append([f"Metric_{c}" for c in range(1, COLS + 1)])
for r in range(2, ROWS + 1):
    data.append([r * c for c in range(1, COLS + 1)])

# Sheet 2: Budget — small, with planted errors
budget = wb.create_sheet("Budget")
budget.append(["Line item", "2024", "2025", "2026"])
budget.append(["Revenue", 1000, 1200, 1500])
budget.append(["COGS", 400, 450, 500])
# 2026 Gross profit is hard-coded: a planted error (should be =D2-D3)
budget.append(["Gross profit", "=B2-B3", "=C2-C3", 1000])
budget.append(["Margin", "=B4/B2", "=C4/C2", "=D4/D2"])
# Per unit divides by empty cells below: will yield #DIV/0!
budget.append(["Per unit", "=B4/B9", "=C4/C9", "=D4/D9"])

wb.save(OUTPUT_PATH)
print(f"wrote {OUTPUT_PATH}")
