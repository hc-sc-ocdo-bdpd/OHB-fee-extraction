"""
Copy the "green" highlight from the master Fee_Comparison_Mismatches workbook's
Mismatches sheet onto the matching rows of another Fee_Comparison_Mismatches
workbook (matched by Sheet/Field/PT/Specialty/Procedure Code).

Which year's files are used comes from scripts/config.py (YEAR); nothing in
this file needs editing to run a different year.

Usage:
    python scripts/highlight_matching_mismatches.py [other.xlsx] [master.xlsx] [output.xlsx]

If other.xlsx is omitted, defaults to that year's Fee_Comparison_Mismatches.xlsx.
If master.xlsx is omitted, defaults to that year's Fee_Comparison_Mismatches_master.xlsx.
If output.xlsx is omitted, the result is saved back over the master,
replacing it -- the newly highlighted file becomes the new master.
"""

import os
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config

# Both default paths follow config.YEAR -- change the year there, not here.
# The row-matching logic itself is year-agnostic (rows are keyed on
# Sheet/Field/PT/Specialty/Procedure Code, and the "Field" values embed the
# year but both workbooks come from the same run, so they agree).
MASTER_FILE_PATH = config.mismatch_master()
OTHER_FILE_PATH = config.mismatch_report()

SHEET_NAME = "Mismatches"
# Key columns used to match a row across the two workbooks. "Ground Truth"
# and "Generated" are deliberately excluded since "Generated" is exactly the
# value expected to potentially differ between the two runs.
KEY_COLUMNS = ["Sheet", "Field", "PT", "Specialty", "Procedure Code"]


def _excel_lock_file(path: Path) -> Path:
    """Where Excel puts its lock file while `path` is open ("~$name.xlsx")."""
    return path.with_name("~$" + path.name)


def _refuse_if_open_in_excel(path: Path) -> None:
    """Stop before writing a workbook Excel currently has open.

    Overwriting a file Excel is holding is how this script produced an
    unopenable workbook: on Windows the write can land on a file Excel still
    owns, and in a OneDrive/SharePoint-synced folder the sync client sees the
    file change underneath it and resolves it as a *merge conflict* rather
    than a clean update. Failing here with the reason is far better than
    handing back a workbook that won't open."""
    lock = _excel_lock_file(path)
    if lock.exists():
        raise SystemExit(
            f"ERROR: {path.name} looks like it is open in Excel (found {lock.name}).\n"
            f"  Close it and re-run -- writing over a workbook Excel still has open\n"
            f"  corrupts it, and in a synced folder (OneDrive/SharePoint) it shows up\n"
            f"  as a merge conflict the next time you open the file."
        )


def _save_atomically(wb, output_path: Path) -> None:
    """Write the workbook to a temp file in the same folder, then move it into
    place. A direct wb.save() onto the destination leaves a half-written file
    if anything interrupts it (an exception, or a sync client uploading
    mid-write), and a half-written .xlsx is exactly what Excel reports as
    damaged. os.replace is atomic within one filesystem, so the destination is
    either the old file or the complete new one, never a partial one."""
    tmp = output_path.with_name(output_path.name + ".tmp")
    try:
        wb.save(tmp)
        os.replace(tmp, output_path)
    finally:
        if tmp.exists():
            tmp.unlink()


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

    _save_atomically(wb, output_path)
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

    print(f"Propagating {config.YEAR} mismatch highlights")
    print(f"  Master (source of highlights) : {master_path}")
    print(f"  Other  (rows to highlight)    : {other_path}")
    print(f"  Output                        : {output_path}")
    for label, path in (("Master", master_path), ("Other", other_path)):
        if not path.exists():
            raise SystemExit(f"ERROR: {label} workbook not found: {path}")
    _refuse_if_open_in_excel(output_path)
    print()

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