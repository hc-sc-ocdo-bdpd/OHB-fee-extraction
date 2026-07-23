"""
Consolidate CDCP price file data and PT (provincial/territorial) association fee
guide data into a "Fee Comparison" workbook, matching the shape of
Data/Output_2026 Fee Comparison/2026 Fee Comparisons v2.xlsx.

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
  - Data/Output_2026 Fee Comparison/2026 Fee Comparisons - generated.xlsx
"""

import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fee_extraction import load_pt_fees, load_pt_fees_by_subspecialty, load_pt_dd_fees, resolve_dd_role_values
from cdcp_loader import load_cdcp_simple_fees, load_cdcp_sp_rows, load_cdcp_dd_fees
from sheet_builders import (
    build_dh_sheet, build_gp_sheet, build_qc_gp_sheet,
    build_sp_sheet, build_qc_sp_sheet, build_dd_sheet, copy_claim_lines_sheet,
)

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "Data"

CDCP_DIR = DATA_DIR / "Input_CDCP price files"
PT_GUIDES_DIR = DATA_DIR / "Input_PT association fee guides"
TEMPLATE_WORKBOOK = DATA_DIR / "Output_2026 Fee Comparison" / "2026 Fee Comparisons v2.xlsx"
OUTPUT_WORKBOOK = DATA_DIR / "Output_2026 Fee Comparison" / "2026 Fee Comparisons - generated.xlsx"

# Provinces/territories in the order they'll appear in the output, matching
# the template's Claim Lines sheet. Not every province has CDCP data for
# every specialty; those are skipped automatically based on what's actually
# in each CDCP price file.
ALL_PROVINCES = ["AB", "BC", "MB", "NB", "NL", "NS", "ON", "PE", "QC", "SK", "NT", "NU", "YT"]
NON_QC_PROVINCES = [p for p in ALL_PROVINCES if p != "QC"]

# A PT fee more than this many times larger/smaller than the matching CDCP
# fee is flagged for manual review rather than trusted silently -- it's
# usually a sign the extractor grabbed the wrong number from an ambiguous
# source layout.
PLAUSIBILITY_MIN_RATIO = 0.15
PLAUSIBILITY_MAX_RATIO = 6.0


def report_plausibility(cdcp_fees: dict[str, float], pt_fees: dict[str, float]) -> None:
    suspects = []
    for code, pt_fee in pt_fees.items():
        cdcp_fee = cdcp_fees.get(code)
        if not cdcp_fee:
            continue
        ratio = pt_fee / cdcp_fee
        if ratio < PLAUSIBILITY_MIN_RATIO or ratio > PLAUSIBILITY_MAX_RATIO:
            suspects.append((code, pt_fee, cdcp_fee, ratio))
    if suspects:
        print(f"    REVIEW: {len(suspects)} code(s) with an implausible PT/CDCP fee ratio "
              f"(possible extraction error, verify manually):")
        for code, pt_fee, cdcp_fee, ratio in suspects:
            print(f"      {code}: PT=${pt_fee:.2f} CDCP=${cdcp_fee:.2f} (ratio {ratio:.2f}x)")


def resolve_pt_fees(specialty_dir: Path, province: str, known_codes: set[str], label: str):
    pt_fees, sources_used, files_found = load_pt_fees(specialty_dir, province, known_codes, verbose=False)
    source_desc = ", ".join(f"{name} ({n})" for name, n in sources_used) or "no PT fee guide found"
    print(f"{province} {label}: {len(known_codes)} CDCP codes, {len(pt_fees)} PT fees matched -- {source_desc}")
    if not files_found:
        print(f"    Note: no PT fee guide file found for {province}/{label}; PT fees will be 'N/A'.")
    return pt_fees


def process_dh(wb_new, template_ws) -> tuple[int, int]:
    specialty_dir = PT_GUIDES_DIR / "DH"
    rows = []
    total_codes = total_matched = 0
    for province in ALL_PROVINCES:
        cdcp_fees = load_cdcp_simple_fees(CDCP_DIR, province, "DH", "DH")
        if not cdcp_fees:
            continue
        pt_fees = resolve_pt_fees(specialty_dir, province, set(cdcp_fees.keys()), "DH")
        report_plausibility(cdcp_fees, pt_fees)
        total_codes += len(cdcp_fees)
        total_matched += len(pt_fees)
        rows.append((province, "DH", cdcp_fees, pt_fees))

    dh_rows = [
        (province, code, cdcp_fees.get(code), pt_fees.get(code))
        for province, _, cdcp_fees, pt_fees in rows
        for code in sorted(cdcp_fees.keys())
    ]
    build_dh_sheet(wb_new, template_ws, dh_rows)
    return total_codes, total_matched


def process_gp(wb_new, template_ws) -> tuple[int, int]:
    specialty_dir = PT_GUIDES_DIR / "GP"
    rows = []
    total_codes = total_matched = 0
    for province in NON_QC_PROVINCES:
        cdcp_fees = load_cdcp_simple_fees(CDCP_DIR, province, "GP", "GP", require_fee=False)
        if not cdcp_fees:
            continue
        pt_fees = resolve_pt_fees(specialty_dir, province, set(cdcp_fees.keys()), "GP")
        report_plausibility(cdcp_fees, pt_fees)
        total_codes += len(cdcp_fees)
        total_matched += len(pt_fees)
        for code in sorted(cdcp_fees.keys()):
            rows.append((province, code, cdcp_fees.get(code), pt_fees.get(code)))

    build_gp_sheet(wb_new, template_ws, rows)
    return total_codes, total_matched


def process_qc_gp(wb_new, template_ws) -> tuple[int, int]:
    specialty_dir = PT_GUIDES_DIR / "GP"
    cdcp_fees = load_cdcp_simple_fees(CDCP_DIR, "QC", "GP", "GP", require_fee=False)
    if not cdcp_fees:
        build_qc_gp_sheet(wb_new, template_ws, [])
        return 0, 0
    pt_fees = resolve_pt_fees(specialty_dir, "QC", set(cdcp_fees.keys()), "QC GP")
    report_plausibility(cdcp_fees, pt_fees)
    rows = [(code, cdcp_fees.get(code), pt_fees.get(code)) for code in sorted(cdcp_fees.keys())]
    build_qc_gp_sheet(wb_new, template_ws, rows)
    return len(cdcp_fees), len(pt_fees)


def _codes_by_subspecialty(sp_rows: list[tuple[str, str, float | None]]) -> dict[str, set[str]]:
    codes_by_sub: dict[str, set[str]] = {}
    for code, sub, _fee in sp_rows:
        codes_by_sub.setdefault(sub, set()).add(code)
    return codes_by_sub


def process_sp(wb_new, template_ws) -> tuple[int, int]:
    specialty_dir = PT_GUIDES_DIR / "SP"
    rows = []
    total_codes = total_matched = 0
    for province in NON_QC_PROVINCES:
        sp_rows = load_cdcp_sp_rows(CDCP_DIR, province, require_fee=False)
        if not sp_rows:
            continue
        pt_fees, sources_used, files_found = load_pt_fees_by_subspecialty(
            specialty_dir, province, _codes_by_subspecialty(sp_rows),
            gp_specialty_dir=PT_GUIDES_DIR / "GP", verbose=False,
        )
        source_desc = ", ".join(f"{name} ({n})" for name, n in sources_used) or "no PT fee guide found"
        print(f"{province} SP: {len(sp_rows)} CDCP codes, {len(pt_fees)} PT fees matched -- {source_desc}")
        if not files_found:
            print(f"    Note: no PT fee guide file found for {province}/SP; PT fees will be 'N/A'.")
        cdcp_fees_by_key = {(code, sub): fee for code, sub, fee in sp_rows}
        report_plausibility(cdcp_fees_by_key, pt_fees)
        total_codes += len(sp_rows)
        total_matched += sum(1 for code, sub, _ in sp_rows if (code, sub) in pt_fees)
        for code, sub_specialty, cdcp_fee in sorted(sp_rows, key=lambda t: (t[1], t[0])):
            rows.append((province, sub_specialty, code, cdcp_fee, pt_fees.get((code, sub_specialty))))

    build_sp_sheet(wb_new, template_ws, rows)
    return total_codes, total_matched


def process_qc_sp(wb_new, template_ws) -> tuple[int, int]:
    specialty_dir = PT_GUIDES_DIR / "SP"
    sp_rows = load_cdcp_sp_rows(CDCP_DIR, "QC", require_fee=False)
    if not sp_rows:
        build_qc_sp_sheet(wb_new, template_ws, [])
        return 0, 0
    pt_fees, sources_used, files_found = load_pt_fees_by_subspecialty(
        specialty_dir, "QC", _codes_by_subspecialty(sp_rows),
        gp_specialty_dir=PT_GUIDES_DIR / "GP", verbose=False,
    )
    source_desc = ", ".join(f"{name} ({n})" for name, n in sources_used) or "no PT fee guide found"
    print(f"QC QC SP: {len(sp_rows)} CDCP codes, {len(pt_fees)} PT fees matched -- {source_desc}")
    if not files_found:
        print("    Note: no PT fee guide file found for QC/QC SP; PT fees will be 'N/A'.")
    cdcp_fees_by_key = {(code, sub): fee for code, sub, fee in sp_rows}
    report_plausibility(cdcp_fees_by_key, pt_fees)
    rows = [
        (sub_specialty, code, cdcp_fee, pt_fees.get((code, sub_specialty)))
        for code, sub_specialty, cdcp_fee in sorted(sp_rows, key=lambda t: (t[1], t[0]))
    ]
    build_qc_sp_sheet(wb_new, template_ws, rows)
    matched = sum(1 for code, sub, _ in sp_rows if (code, sub) in pt_fees)
    return len(sp_rows), matched


def resolve_pt_dd_fees(specialty_dir: Path, province: str, known_codes: set[str]):
    role_fees, sources_used, files_found = load_pt_dd_fees(specialty_dir, province, known_codes, verbose=False)
    source_desc = ", ".join(f"{name} ({n})" for name, n in sources_used) or "no PT fee guide found"
    print(f"{province} DD: {len(known_codes)} CDCP codes, {len(role_fees)} PT fees matched -- {source_desc}")
    if not files_found:
        print(f"    Note: no PT fee guide file found for {province}/DD; PT fees will be 'N/A'.")
    return role_fees


def process_dd(wb_new, template_ws) -> tuple[int, int]:
    specialty_dir = PT_GUIDES_DIR / "DD"
    rows = []
    total_codes = total_matched = 0
    for province in ALL_PROVINCES:
        cdcp_fees = load_cdcp_dd_fees(CDCP_DIR, province)
        if not cdcp_fees:
            continue
        prof_fees_only = {code: prof for code, (prof, _lab) in cdcp_fees.items()}
        role_fees = resolve_pt_dd_fees(specialty_dir, province, set(cdcp_fees.keys()))
        pt_combo_fees = {}
        for code, values in role_fees.items():
            combo = resolve_dd_role_values(values)[2]
            if combo is not None:
                pt_combo_fees[code] = combo
        report_plausibility(prof_fees_only, pt_combo_fees)
        total_codes += len(cdcp_fees)
        total_matched += len(role_fees)
        for code in sorted(cdcp_fees.keys()):
            pt_prof, pt_lab, pt_combo = resolve_dd_role_values(role_fees.get(code, {}))
            rows.append((province, code, cdcp_fees[code], (pt_prof, pt_lab, pt_combo)))

    build_dd_sheet(wb_new, template_ws, rows)
    return total_codes, total_matched


def main() -> None:
    template_wb = openpyxl.load_workbook(TEMPLATE_WORKBOOK, data_only=False)

    wb_new = openpyxl.Workbook()
    wb_new.remove(wb_new.active)
    copy_claim_lines_sheet(wb_new, template_wb["Claim Lines"])

    overall_codes = overall_matched = 0

    print("=== DH ===")
    codes, matched = process_dh(wb_new, template_wb["DH"])
    overall_codes += codes; overall_matched += matched

    print("\n=== GP ===")
    codes, matched = process_gp(wb_new, template_wb["GP"])
    overall_codes += codes; overall_matched += matched

    print("\n=== QC GP ===")
    codes, matched = process_qc_gp(wb_new, template_wb["QC GP"])
    overall_codes += codes; overall_matched += matched

    print("\n=== SP ===")
    codes, matched = process_sp(wb_new, template_wb["SP"])
    overall_codes += codes; overall_matched += matched

    print("\n=== QC SP ===")
    codes, matched = process_qc_sp(wb_new, template_wb["QC SP"])
    overall_codes += codes; overall_matched += matched

    print("\n=== DD ===")
    codes, matched = process_dd(wb_new, template_wb["DD"])
    overall_codes += codes; overall_matched += matched

    wb_new.save(OUTPUT_WORKBOOK)
    print(f"\nSaved {OUTPUT_WORKBOOK}")
    print(f"Overall PT fee match rate: {overall_matched}/{overall_codes} "
          f"({100 * overall_matched / overall_codes:.1f}%)")


if __name__ == "__main__":
    main()
