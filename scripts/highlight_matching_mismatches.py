"""
Copy the "green" highlight from the master Fee_Comparison_Mismatches workbook's
Mismatches sheet onto the matching rows of another Fee_Comparison_Mismatches
workbook (matched by Sheet/Field/PT/Specialty/Procedure Code).

Usage:
    python scripts/highlight_matching_mismatches.py [other.xlsx] [master.xlsx] [output.xlsx]

If other.xlsx is omitted, defaults to OTHER_FILE_PATH below.
If master.xlsx is omitted, defaults to MASTER_FILE_PATH below.
If output.xlsx is omitted, the result is saved back to MASTER_FILE_PATH,
replacing it -- the newly highlighted file becomes the new master.
"""

import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill

MASTER_FILE_PATH = (
    r"C:\Users\JOGILL\OneDrive - HC-SC PHAC-ASPC\Desktop\OHB\Data"
    r"\Output_2026 Fee Comparison\Fee_Comparison_Mismatches_master.xlsx"
)
OTHER_FILE_PATH = (
    r"C:\Users\JOGILL\OneDrive - HC-SC PHAC-ASPC\Desktop\OHB\Data"
    r"\Output_2026 Fee Comparison\Fee_Comparison_Mismatches.xlsx"
)

SHEET_NAME = "Mismatches"
# Key columns used to match a row across the two workbooks. "Ground Truth"
# and "Generated" are deliberately excluded since "Generated" is exactly the
# value expected to potentially differ between the two runs.
KEY_COLUMNS = ["Sheet", "Field", "PT", "Specialty", "Procedure Code"]


def _is_green_highlight(cell) -> bool:
    fill = cell.fill
    if fill is None or fill.patternType != "solid":
        return False
    fg = fill.fgColor
    if fg is None or fg.type != "theme":
        return False
    if fg.theme != 6:
        return False
    tint = fg.tint or 0.0
    return round(tint, 1) == 0.6


def _header_index_map(sheet):
    header = [c.value for c in next(sheet.iter_rows(min_row=1, max_row=1))]
    return {name: idx for idx, name in enumerate(header)}


def _row_key(row_cells, col_index):
    return tuple(row_cells[col_index[name]].value for name in KEY_COLUMNS)


def build_highlighted_key_set(master_path: Path) -> tuple[set, PatternFill]:
    wb = openpyxl.load_workbook(master_path)
    sheet = wb[SHEET_NAME]
    col_index = _header_index_map(sheet)

    highlight_fill = None
    keys = set()
    for row in sheet.iter_rows(min_row=2):
        if not row[0].value and all(c.value is None for c in row):
            continue
        first_cell = row[col_index["Sheet"]]
        if _is_green_highlight(first_cell):
            keys.add(_row_key(row, col_index))
            if highlight_fill is None:
                highlight_fill = PatternFill(
                    fill_type=first_cell.fill.patternType,
                    fgColor=first_cell.fill.fgColor,
                    bgColor=first_cell.fill.bgColor,
                )
    wb.close()
    return keys, highlight_fill


def apply_highlight(other_path: Path, keys: set, highlight_fill: PatternFill, output_path: Path) -> int:
    wb = openpyxl.load_workbook(other_path)
    sheet = wb[SHEET_NAME]
    col_index = _header_index_map(sheet)

    matched = 0
    for row in sheet.iter_rows(min_row=2):
        if all(c.value is None for c in row):
            continue
        key = _row_key(row, col_index)
        if key in keys:
            matched += 1
            for cell in row:
                cell.fill = highlight_fill

    wb.save(output_path)
    wb.close()
    return matched


def main():
    if len(sys.argv) not in (1, 2, 3, 4):
        print(
            "Usage: python highlight_matching_mismatches.py [other.xlsx] "
            "[master.xlsx] [output.xlsx]"
        )
        sys.exit(1)

    other_path = Path(sys.argv[1]) if len(sys.argv) >= 2 else Path(OTHER_FILE_PATH)
    master_path = Path(sys.argv[2]) if len(sys.argv) >= 3 else Path(MASTER_FILE_PATH)
    output_path = Path(sys.argv[3]) if len(sys.argv) == 4 else Path(MASTER_FILE_PATH)

    keys, highlight_fill = build_highlighted_key_set(master_path)
    print(f"Found {len(keys)} highlighted rows in master's {SHEET_NAME} sheet.")

    if highlight_fill is None:
        print("No highlighted rows found in master file; nothing to copy.")
        sys.exit(0)

    matched = apply_highlight(other_path, keys, highlight_fill, output_path)
    print(f"Highlighted {matched} matching row(s) in {other_path.name}.")
    print(f"Wrote result to: {output_path}")


if __name__ == "__main__":
    main()