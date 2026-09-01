"""
Consolidate CDCP price file data and PT (provincial/territorial) association fee
guide data into a "Fee Comparison" workbook, matching the shape of the
template workbook (see config.template_workbook).

SCOPE: this script only BUILDS the workbook. It does not check the result
against anything -- comparing the generated output to the ground truth (row
matching, per-field match rates, mismatch/missing/extra reporting) is
compare_fee_files.py's job, and duplicating any of it here just produced a
second, differently-defined set of numbers to reconcile. Run
compare_fee_files.py after this to evaluate the output.

The per-province lines printed during a run are build provenance, not
evaluation: they say how many codes came out of each CDCP price file and
which fee-guide files supplied the PT fees, so a province with a missing or
unreadable source is visible while the build is happening.

Covers all five specialty categories: DH, DD, GP, SP (each split into a main
sheet for all provinces except QC, plus a separate "QC GP"/"QC SP" sheet,
matching the template -- QC's CDCP/PT data uses different code sets and the
template kept it structurally separate).

For each province, PT fees are resolved from whatever fee-guide file(s)
actually exist for that province/specialty (see scripts/fee_extraction.py):
xlsx/xlsm/xls spreadsheets first, then csv, then docx, then pdf -- trying
each in turn and filling in only the codes still missing at each step, so
multiple partial sources combine. A province with no fee guide file at all
gets "N/A" PT fees throughout, same as the template.

Known gaps and deliberate deviations from the template (see conversation for
the full rationale):
  - None of the input files contain 2025 fees, so all '2025 CDCP Fee' /
    '2025 PT Fee' columns are written as "N/A".
  - CDCP claim-count data isn't available from these files, so claim-count
    columns are written as 0. The weighting formulas are still written so
    the workbook is ready to be populated with real claim counts later.
  - QC GP / QC SP's CDCP and PT fee columns were formulas referencing
    *external linked workbooks* in the template that aren't present in our
    Data folder (e.g. "=VLOOKUP(C2,[4]GP!$C:$F,4,FALSE)"). These are
    replaced with plain literal values from our own extraction.
  - GP's template formulas use Excel Table structured references
    ("Table1[[#This Row],[...]]") and its last column was already a broken
    #REF! in the template itself. Both are replaced with equivalent plain
    cell-reference formulas.
  - SP procedure codes commonly repeat under several different
    sub-specialties (e.g. one code billed under both Periodontics and Oral
    Surgery) with potentially different fees. Where a province's PT guide is
    split into per-sub-specialty files (identifiable by filename, e.g. "ON PA
    Fee Guide.xlsx" or "MDA 2026 Periodontics....xlsx"), that file is
    preferred for its matching sub-specialty's codes over a general/combined
    guide -- see fee_extraction.load_pt_fees_by_subspecialty. Where a
    province's guide is one combined file with no per-sub-specialty split
    (e.g. one PDF listing the same code under several specialty sections),
    the first fee found is used; exact sub-specialty attribution isn't
    attempted in that case.

Output:
  - <Data>/<year>/<year>_Output/<year> Fee Comparisons - generated.xlsx
"""

import re
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config
from fee_extraction import load_pt_fees, load_pt_fees_by_subspecialty, load_pt_dd_fees, resolve_dd_role_values
from cdcp_loader import load_cdcp_simple_fees, load_cdcp_sp_rows, load_cdcp_dd_fees
from sheet_builders import (
    build_dh_sheet, build_gp_sheet, build_qc_gp_sheet,
    build_sp_sheet, build_qc_sp_sheet, build_dd_sheet, copy_claim_lines_sheet,
    set_year_label_map,
)

# Every path below derives from config.YEAR -- change the year there, not
# here. See scripts/config.py for how each folder name is resolved (the
# naming convention differs between years).
BASE_DIR = config.BASE_DIR
DATA_DIR = config.DATA_DIR

YEAR = config.YEAR
CDCP_DIR = config.cdcp_dir()
PT_GUIDES_DIR = config.pt_guides_dir()
TEMPLATE_WORKBOOK = config.template_workbook()
OUTPUT_WORKBOOK = config.output_workbook()

# Provinces/territories in the order they'll appear in the output, matching
# the template's Claim Lines sheet. Not every province has CDCP data for
# every specialty; those are skipped automatically based on what's actually
# in each CDCP price file.
ALL_PROVINCES = ["AB", "BC", "MB", "NB", "NL", "NS", "ON", "PE", "QC", "SK", "NT", "NU", "YT"]
NON_QC_PROVINCES = [p for p in ALL_PROVINCES if p != "QC"]

def resolve_pt_fees(specialty_dir: Path, province: str, known_codes: set[str], label: str):
    pt_fees, sources_used, files_found = load_pt_fees(specialty_dir, province, known_codes, verbose=False)
    source_desc = ", ".join(f"{name} ({n})" for name, n in sources_used) or "no PT fee guide found"
    print(f"{province} {label}: {len(known_codes)} CDCP codes, {len(pt_fees)} PT fees extracted -- {source_desc}")
    if not files_found:
        print(f"    Note: no PT fee guide file found for {province}/{label}; PT fees will be 'N/A'.")
    return pt_fees


def process_dh(wb_new, template_ws) -> None:
    specialty_dir = PT_GUIDES_DIR / "DH"
    rows = []
    for province in ALL_PROVINCES:
        cdcp_fees = load_cdcp_simple_fees(CDCP_DIR, province, "DH", "DH")
        if not cdcp_fees:
            continue
        pt_fees = resolve_pt_fees(specialty_dir, province, set(cdcp_fees.keys()), "DH")
        rows.append((province, "DH", cdcp_fees, pt_fees))

    dh_rows = [
        (province, code, cdcp_fees.get(code), pt_fees.get(code))
        for province, _, cdcp_fees, pt_fees in rows
        for code in sorted(cdcp_fees.keys())
    ]
    build_dh_sheet(wb_new, template_ws, dh_rows)


def process_gp(wb_new, template_ws) -> None:
    specialty_dir = PT_GUIDES_DIR / "GP"
    rows = []
    for province in NON_QC_PROVINCES:
        cdcp_fees = load_cdcp_simple_fees(CDCP_DIR, province, "GP", "GP", require_fee=False)
        if not cdcp_fees:
            continue
        pt_fees = resolve_pt_fees(specialty_dir, province, set(cdcp_fees.keys()), "GP")
        for code in sorted(cdcp_fees.keys()):
            rows.append((province, code, cdcp_fees.get(code), pt_fees.get(code)))

    build_gp_sheet(wb_new, template_ws, rows)


def process_qc_gp(wb_new, template_ws) -> None:
    specialty_dir = PT_GUIDES_DIR / "GP"
    cdcp_fees = load_cdcp_simple_fees(CDCP_DIR, "QC", "GP", "GP", require_fee=False)
    if not cdcp_fees:
        build_qc_gp_sheet(wb_new, template_ws, [])
        return
    pt_fees = resolve_pt_fees(specialty_dir, "QC", set(cdcp_fees.keys()), "QC GP")
    rows = [(code, cdcp_fees.get(code), pt_fees.get(code)) for code in sorted(cdcp_fees.keys())]
    build_qc_gp_sheet(wb_new, template_ws, rows)


def _codes_by_subspecialty(sp_rows: list[tuple[str, str, float | None]]) -> dict[str, set[str]]:
    codes_by_sub: dict[str, set[str]] = {}
    for code, sub, _fee in sp_rows:
        codes_by_sub.setdefault(sub, set()).add(code)
    return codes_by_sub


def process_sp(wb_new, template_ws) -> None:
    specialty_dir = PT_GUIDES_DIR / "SP"
    rows = []
    for province in NON_QC_PROVINCES:
        sp_rows = load_cdcp_sp_rows(CDCP_DIR, province, require_fee=False)
        if not sp_rows:
            continue
        pt_fees, sources_used, files_found = load_pt_fees_by_subspecialty(
            specialty_dir, province, _codes_by_subspecialty(sp_rows),
            gp_specialty_dir=PT_GUIDES_DIR / "GP", verbose=False,
        )
        source_desc = ", ".join(f"{name} ({n})" for name, n in sources_used) or "no PT fee guide found"
        print(f"{province} SP: {len(sp_rows)} CDCP codes, {len(pt_fees)} PT fees extracted -- {source_desc}")
        if not files_found:
            print(f"    Note: no PT fee guide file found for {province}/SP; PT fees will be 'N/A'.")
        for code, sub_specialty, cdcp_fee in sorted(sp_rows, key=lambda t: (t[1], t[0])):
            rows.append((province, sub_specialty, code, cdcp_fee, pt_fees.get((code, sub_specialty))))

    build_sp_sheet(wb_new, template_ws, rows)


def process_qc_sp(wb_new, template_ws) -> None:
    specialty_dir = PT_GUIDES_DIR / "SP"
    sp_rows = load_cdcp_sp_rows(CDCP_DIR, "QC", require_fee=False)
    if not sp_rows:
        build_qc_sp_sheet(wb_new, template_ws, [])
        return
    pt_fees, sources_used, files_found = load_pt_fees_by_subspecialty(
        specialty_dir, "QC", _codes_by_subspecialty(sp_rows),
        gp_specialty_dir=PT_GUIDES_DIR / "GP", verbose=False,
    )
    source_desc = ", ".join(f"{name} ({n})" for name, n in sources_used) or "no PT fee guide found"
    print(f"QC QC SP: {len(sp_rows)} CDCP codes, {len(pt_fees)} PT fees extracted -- {source_desc}")
    if not files_found:
        print("    Note: no PT fee guide file found for QC/QC SP; PT fees will be 'N/A'.")
    rows = [
        (sub_specialty, code, cdcp_fee, pt_fees.get((code, sub_specialty)))
        for code, sub_specialty, cdcp_fee in sorted(sp_rows, key=lambda t: (t[1], t[0]))
    ]
    build_qc_sp_sheet(wb_new, template_ws, rows)


def resolve_pt_dd_fees(specialty_dir: Path, province: str, known_codes: set[str]):
    role_fees, sources_used, files_found = load_pt_dd_fees(specialty_dir, province, known_codes, verbose=False)
    source_desc = ", ".join(f"{name} ({n})" for name, n in sources_used) or "no PT fee guide found"
    print(f"{province} DD: {len(known_codes)} CDCP codes, {len(role_fees)} PT fees extracted -- {source_desc}")
    if not files_found:
        print(f"    Note: no PT fee guide file found for {province}/DD; PT fees will be 'N/A'.")
    return role_fees


def process_dd(wb_new, template_ws) -> None:
    specialty_dir = PT_GUIDES_DIR / "DD"
    rows = []
    for province in ALL_PROVINCES:
        cdcp_fees = load_cdcp_dd_fees(CDCP_DIR, province)
        if not cdcp_fees:
            continue
        role_fees = resolve_pt_dd_fees(specialty_dir, province, set(cdcp_fees.keys()))
        for code in sorted(cdcp_fees.keys()):
            pt_prof, pt_lab, pt_combo = resolve_dd_role_values(role_fees.get(code, {}))
            rows.append((province, code, cdcp_fees[code], (pt_prof, pt_lab, pt_combo)))

    build_dd_sheet(wb_new, template_ws, rows)


_YEAR_IN_TEXT_RE = re.compile(r"\b(20\d{2})\b")


def template_years(template_wb) -> list[int]:
    """The distinct years the template's own header rows are labelled with,
    oldest first.

    Detected rather than configured so the template can be relabelled or
    replaced without a matching config edit -- only the years we want *out*
    (config.PAST_YEAR / CURRENT_YEAR) need stating. Only the first two rows
    of each sheet are scanned: that's where the year labels live, and it
    keeps a stray year inside the data (a procedure description mentioning
    2019, say) from being mistaken for a column label."""
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
    """Map each year in the template's headers to the year it should read as
    in the generated workbook.

    The output carries two years side by side, so the template's older year
    becomes config.past_year() and its newer one config.current_year(). A
    template already labelled for the target years yields an identity map
    (i.e. no relabelling), which is the normal 2026 case."""
    years = template_years(template_wb)
    target_past, target_current = config.past_year(), config.current_year()
    if len(years) >= 2:
        # Oldest -> past, newest -> current. Anything between (templates
        # shouldn't have any, but be explicit) is left alone rather than
        # guessed at.
        return {years[0]: target_past, years[-1]: target_current}
    if len(years) == 1:
        return {years[0]: target_current}
    return {}


def main() -> None:
    # Printed up front so a mis-resolved folder is obvious immediately rather
    # than surfacing as a run full of "no PT fee guide found" notes. The
    # legacy 2026 folder names carry no year in them, so if a year's own
    # folder is missing the resolver can fall back to a differently-yeared
    # one -- this banner is what makes that visible.
    print(f"=== Building {YEAR} Fee Comparison ===")
    print(f"  CDCP price files : {CDCP_DIR}")
    print(f"  PT fee guides    : {PT_GUIDES_DIR}")
    print(f"  Template         : {TEMPLATE_WORKBOOK}")
    print(f"  Output           : {OUTPUT_WORKBOOK}")
    for label, path in (("CDCP price files", CDCP_DIR), ("PT fee guides", PT_GUIDES_DIR)):
        if not path.exists():
            print(f"  WARNING: {label} not found at the path above.")

    # Checked here, not deep in the build: a workbook that isn't really the
    # template used to surface as "KeyError: 'Worksheet Claim Lines does not
    # exist.'" partway through, which says nothing about the actual problem.
    if not TEMPLATE_WORKBOOK.exists():
        looked = "\n    ".join(str(p) for p in config.template_candidates())
        raise SystemExit(
            f"\nERROR: no template workbook found. A template must contain the sheets "
            f"{', '.join(config.REQUIRED_TEMPLATE_SHEETS)}.\n"
            f"  Looked in:\n    {looked}\n"
            f"  Set config.TEMPLATE_FILE (or the OHB_TEMPLATE environment variable) "
            f"to point at it directly."
        )
    if not config._has_required_sheets(TEMPLATE_WORKBOOK):
        raise SystemExit(
            f"\nERROR: {TEMPLATE_WORKBOOK.name} is missing one or more required sheets "
            f"({', '.join(config.REQUIRED_TEMPLATE_SHEETS)}), so it can't be used as the "
            f"template.\n  Set config.TEMPLATE_FILE (or OHB_TEMPLATE) to the real template."
        )

    template_wb = openpyxl.load_workbook(TEMPLATE_WORKBOOK, data_only=False)

    # One template serves every year; only its year labels are rewritten.
    year_map = build_year_label_map(template_wb)
    set_year_label_map(year_map)
    if year_map and any(k != v for k, v in year_map.items()):
        relabel = ", ".join(f"{k} -> {v}" for k, v in sorted(year_map.items()))
        print(f"  Header years     : {relabel}")
    else:
        print(f"  Header years     : {config.past_year()}, {config.current_year()} (template already matches)")
    print()

    wb_new = openpyxl.Workbook()
    wb_new.remove(wb_new.active)
    copy_claim_lines_sheet(wb_new, template_wb["Claim Lines"])

    print("=== DH ===")
    process_dh(wb_new, template_wb["DH"])

    print("\n=== GP ===")
    process_gp(wb_new, template_wb["GP"])

    print("\n=== QC GP ===")
    process_qc_gp(wb_new, template_wb["QC GP"])

    print("\n=== SP ===")
    process_sp(wb_new, template_wb["SP"])

    print("\n=== QC SP ===")
    process_qc_sp(wb_new, template_wb["QC SP"])

    print("\n=== DD ===")
    process_dd(wb_new, template_wb["DD"])

    # A year being built for the first time won't have an output folder yet.
    OUTPUT_WORKBOOK.parent.mkdir(parents=True, exist_ok=True)
    wb_new.save(OUTPUT_WORKBOOK)
    print(f"\nSaved {OUTPUT_WORKBOOK}")
    print("Run compare_fee_files.py to check this output against the ground truth.")


if __name__ == "__main__":
    main()
