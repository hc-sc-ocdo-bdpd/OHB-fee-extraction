"""
Build a "Fee Comparison" workbook carrying TWO rate years side by side.

Each sheet gets one row per (province, [sub-specialty,] procedure code), with
that code's CDCP and PT fees for both years:

    <past> CDCP Fee | <current> CDCP Fee | <past> PT Fee | <current> PT Fee

Both years are populated from their own data folders. Which two years those
are is the only thing you change to run this again next year -- set
config.PAST_YEAR and config.CURRENT_YEAR and everything else follows: folder
paths, CDCP price-file names, header labels and the output filename.

HOW IT FITS TOGETHER
    config.py          the two years, and every path derived from them
    fee_dataset.py     loads ONE year's fees for ONE sheet -> plain dicts
    sheet_builders.py  turns rows into a styled sheet matching the template
    this script        asks fee_dataset for each year, merges, writes

That split is the point: a year is a parameter to fee_dataset, not a special
case in the layout code. fee_extraction.py and cdcp_loader.py are used
exactly as they are -- fee_dataset points them at the right year's folders
rather than changing how they read a file.

SCOPE: this script only BUILDS the workbook. Checking the result against the
ground truth -- row matching, per-field match rates, mismatch reporting -- is
compare_fee_files.py's job. Run it afterwards.

The per-province lines printed during a run are build provenance, not
evaluation: they say how many codes came out of each CDCP price file and
which fee-guide files supplied the PT fees, so a missing or unreadable source
is visible while the build is happening.

Covers all five specialty categories: DH, GP, SP and DD, with GP and SP each
split into a main sheet for every province except QC plus a separate
"QC GP"/"QC SP" sheet -- QC's CDCP and PT data use different code sets and
the template keeps them structurally separate.

A year whose folders aren't present yet is not an error: its columns come out
"N/A" and the build continues, so a pair can be built before the second
year's guides have all arrived.

Known gaps and deliberate deviations from the template:
  - CDCP claim-count data isn't available from these files, so claim-count
    columns are written as 0. The weighting formulas are still written, so
    the workbook is ready for real claim counts later.
  - QC GP / QC SP's fee columns were formulas referencing external linked
    workbooks that aren't in our Data folder (e.g.
    "=VLOOKUP(C2,[4]GP!$C:$F,4,FALSE)"). These are replaced with plain
    literal values from our own extraction.
  - GP's template formulas use Excel Table structured references
    ("Table1[[#This Row],[...]]") and its last column was already a broken
    #REF! in the template itself. Both are replaced with equivalent plain
    cell-reference formulas.

Output:
  <Data>/<current>/<current>_Output/<current> Fee Comparisons - generated.xlsx

Usage:
    python scripts/build_fee_comparison.py
    python scripts/build_fee_comparison.py --years 2026 2027
"""

import argparse
import re
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config
from fee_dataset import SHEETS, load_both_years, merged_keys
from sheet_builders import (
    build_dd_sheet, build_dh_sheet, build_gp_sheet, build_qc_gp_sheet,
    build_qc_sp_sheet, build_sp_sheet, copy_claim_lines_sheet, set_year_label_map,
)

_YEAR_IN_TEXT_RE = re.compile(r"\b(20\d{2})\b")

# Which builder writes each sheet. Keeping this a lookup rather than a chain
# of ifs means adding a sheet is one entry here plus one loader in
# fee_dataset, with nothing else to touch.
_SHEET_BUILDERS = {
    "DH": build_dh_sheet,
    "GP": build_gp_sheet,
    "QC GP": build_qc_gp_sheet,
    "SP": build_sp_sheet,
    "QC SP": build_qc_sp_sheet,
    "DD": build_dd_sheet,
}


# ---------------------------------------------------------------------------
# Turning two years of loaded data into rows
# ---------------------------------------------------------------------------

def _pair(by_year: dict, key) -> tuple:
    """(past_year_value, current_year_value) for one identity."""
    past, current = config.years()
    return (by_year[past].get(key), by_year[current].get(key))


def build_rows(by_year: dict) -> list[tuple]:
    """One row per identity either year knows about, each carrying both years.

    The row is the identity's key parts followed by the CDCP pair and the PT
    pair, which is exactly what every sheet_builders function expects -- so
    the same merge works for all six sheets despite their differing keys.
    """
    cdcp_by_year = {y: d.cdcp for y, d in by_year.items()}
    pt_by_year = {y: d.pt for y, d in by_year.items()}
    return [
        (*key, _pair(cdcp_by_year, key), _pair(pt_by_year, key))
        for key in merged_keys(by_year)
    ]


def build_sheet(wb_new, template_wb, sheet: str) -> int:
    """Load both years for one sheet, merge them, and write it."""
    print(f"\n=== {sheet} ===")
    by_year = load_both_years(sheet)
    for year in config.years():
        data = by_year[year]
        print(f"  {data.summary()}")
        for note in data.notes:
            print(note)

    rows = build_rows(by_year)
    _SHEET_BUILDERS[sheet](wb_new, template_wb[sheet], rows)
    print(f"  -> {len(rows)} rows written")
    return len(rows)


# ---------------------------------------------------------------------------
# Template year labels
# ---------------------------------------------------------------------------

def template_years(template_wb) -> list[int]:
    """The distinct years the template's own header rows are labelled with,
    oldest first.

    Detected rather than configured, so the template can be relabelled or
    replaced without a matching config edit. Only the first two rows of each
    sheet are scanned: that's where the year labels live, and it keeps a
    stray year inside the data from being mistaken for a column label.
    """
    found: set[int] = set()
    for name in config.REQUIRED_TEMPLATE_SHEETS:
        if name not in template_wb.sheetnames:
            continue
        ws = template_wb[name]
        for row in ws.iter_rows(min_row=1, max_row=2):
            for cell in row:
                if isinstance(cell.value, str):
                    found.update(int(m) for m in _YEAR_IN_TEXT_RE.findall(cell.value))
    return sorted(found)


def build_year_label_map(template_wb) -> dict[int, int]:
    """Map each year in the template's headers to the year it should read as.

    The template's older year becomes config.PAST_YEAR and its newer one
    config.CURRENT_YEAR. A template already labelled for the target pair
    yields an identity map, i.e. no relabelling.
    """
    years = template_years(template_wb)
    target_past, target_current = config.years()
    if len(years) >= 2:
        return {years[0]: target_past, years[-1]: target_current}
    if len(years) == 1:
        return {years[0]: target_current}
    return {}


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def resolve_template() -> Path:
    """The template workbook, checked before the build rather than during it.

    A workbook that isn't really the template used to surface as
    "KeyError: 'Worksheet Claim Lines does not exist.'" partway through,
    which says nothing about the actual problem.
    """
    template = config.template_workbook()
    if not template.exists():
        looked = "\n    ".join(str(p) for p in config.template_candidates())
        raise SystemExit(
            f"\nERROR: no template workbook found. A template must contain the sheets "
            f"{', '.join(config.REQUIRED_TEMPLATE_SHEETS)}.\n"
            f"  Put it in {config.DATA_DIR / 'templates'} -- a workbook there is used "
            f"under any filename, and wins over one found anywhere else.\n"
            f"  Looked in:\n    {looked}\n"
            f"  Or set config.TEMPLATE_FILE (or the OHB_TEMPLATE environment variable) "
            f"to point at it directly."
        )
    if not config._has_required_sheets(template):
        raise SystemExit(
            f"\nERROR: {template.name} is missing one or more required sheets "
            f"({', '.join(config.REQUIRED_TEMPLATE_SHEETS)}), so it can't be used as the "
            f"template.\n  Set config.TEMPLATE_FILE (or OHB_TEMPLATE) to the real template."
        )
    return template


def print_banner(template: Path, output: Path) -> None:
    """Printed up front so a mis-resolved folder is obvious immediately rather
    than surfacing as a run full of "no PT fee guide found" notes."""
    past, current = config.years()
    print(f"=== Building Fee Comparison: {past} and {current} ===")
    for year in (past, current):
        cdcp, guides = config.cdcp_dir(year), config.pt_guides_dir(year)
        missing = "   <- NOT FOUND; this year's columns will be N/A"
        print(f"  {year} CDCP price files : {cdcp}{'' if cdcp.exists() else missing}")
        print(f"  {year} PT fee guides    : {guides}{'' if guides.exists() else missing}")
    where = "Data/templates" if config.templates_dir() in template.parents else "fallback search"
    print(f"  Template              : {template}   [{where}]")
    print(f"  Output                : {output}")


def main(argv=None) -> None:
    parser = argparse.ArgumentParser(
        description="Build a Fee Comparison workbook carrying two rate years.")
    parser.add_argument(
        "--years", nargs=2, type=int, metavar=("PAST", "CURRENT"),
        help="override config.PAST_YEAR / config.CURRENT_YEAR for this run")
    args = parser.parse_args(argv)

    if args.years:
        config.PAST_YEAR, config.CURRENT_YEAR = sorted(args.years)

    template_path = resolve_template()
    output_path = config.output_workbook()
    print_banner(template_path, output_path)

    template_wb = openpyxl.load_workbook(template_path, data_only=False)

    # One template serves every pair of years; only its labels are rewritten.
    year_map = build_year_label_map(template_wb)
    set_year_label_map(year_map)
    if year_map and any(k != v for k, v in year_map.items()):
        relabel = ", ".join(f"{k} -> {v}" for k, v in sorted(year_map.items()))
        print(f"  Header years          : {relabel}")
    else:
        print(f"  Header years          : {config.past_year()}, {config.current_year()}"
              f" (template already matches)")

    wb_new = openpyxl.Workbook()
    wb_new.remove(wb_new.active)
    copy_claim_lines_sheet(wb_new, template_wb["Claim Lines"])

    total = sum(build_sheet(wb_new, template_wb, sheet) for sheet in SHEETS)

    # A year being built for the first time won't have an output folder yet.
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb_new.save(output_path)
    print(f"\nSaved {output_path}  ({total} rows across {len(SHEETS)} sheets)")
    print("Run compare_fee_files.py to check this output against the ground truth.")


if __name__ == "__main__":
    main()
