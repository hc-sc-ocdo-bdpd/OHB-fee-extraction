"""
Consolidate CDCP price file data and PT (provincial/territorial) association fee
guide data into a "Fee Comparison" workbook, matching the shape of
Data/Output_2026 Fee Comparison/2026 Fee Comparisons v2.xlsx.

Current scope: Dental Hygienist (DH), all provinces/territories that have a
DH sheet in their CDCP price file (AB, BC, MB, NB, NL, NS, ON, PE, QC, SK --
NT, NU, YT have no CDCP DH data at all, matching the template).

For each province, PT fees are resolved from whatever fee-guide file(s)
actually exist for that province/specialty (see scripts/fee_extraction.py):
xlsx/xlsm/xls spreadsheets first, then csv, then docx, then pdf -- trying
each in turn and filling in only the codes still missing at each step, so
multiple partial sources combine. A province with no fee guide file at all
(e.g. SK, NS for DH) gets "N/A" PT fees throughout, same as the template.

Known gaps:
  - None of the input files contain 2025 fees, so '2025 CDCP Fee' and
    '2025 PT Fee' are written as "N/A".
  - CDCP claim-count data (used to compute weighted increases in columns
    I-Y) is not available from these files, so claim-count columns (L, O, P)
    are written as 0. The formulas are still written so the workbook is
    ready to be populated with real claim counts later.

Output:
  - Data/Output_2026 Fee Comparison/2026 Fee Comparisons - generated.xlsx
"""

import copy
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fee_extraction import load_pt_fees, normalize_code

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "Data"

CDCP_DIR = DATA_DIR / "Input_CDCP price files"
PT_GUIDES_DIR = DATA_DIR / "Input_PT association fee guides"
TEMPLATE_WORKBOOK = DATA_DIR / "Output_2026 Fee Comparison" / "2026 Fee Comparisons v2.xlsx"
OUTPUT_WORKBOOK = DATA_DIR / "Output_2026 Fee Comparison" / "2026 Fee Comparisons - generated.xlsx"

SPECIALTY = "DH"

# Provinces/territories in the order they'll appear in the output, matching
# the template's Claim Lines sheet. Not every province has CDCP data for
# every specialty (e.g. no DH data at all for NT/NU/YT); those are skipped
# automatically based on what's actually in each CDCP price file.
ALL_PROVINCES = ["AB", "BC", "MB", "NB", "NL", "NS", "ON", "PE", "QC", "SK", "NT", "NU", "YT"]

# A PT fee more than this many times larger/smaller than the matching CDCP
# fee is flagged for manual review rather than trusted silently -- it's
# usually a sign the extractor grabbed the wrong number from an ambiguous
# source layout.
PLAUSIBILITY_MIN_RATIO = 0.15
PLAUSIBILITY_MAX_RATIO = 6.0


def load_cdcp_fees(province: str, specialty: str) -> dict[str, float] | None:
    """Read 2026 CDCP provider fees per procedure code for one province/specialty.

    Returns None if this province's CDCP price file has no sheet for the
    given specialty (e.g. no DH data at all for NT/NU/YT).
    """
    path = CDCP_DIR / f"2026 - CDCP PRICE FILE - {province}.xlsx"
    if not path.exists():
        return None
    wb = openpyxl.load_workbook(path, data_only=True)
    if specialty not in wb.sheetnames:
        return None
    ws = wb[specialty]
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idx = {name: i for i, name in enumerate(header)}

    fees: dict[str, float] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[idx["Province"]] != province or row[idx["Specialty"]] != specialty:
            continue
        fee = row[idx["Provider Fee"]]
        if fee is None:
            continue
        code = normalize_code(row[idx["Procedure Code"]])
        if code is None:
            continue
        fees[code] = float(fee)
    return fees


def _dh_header_row(template_ws) -> list[str]:
    return [c.value for c in next(template_ws.iter_rows(min_row=1, max_row=1))]


def build_dh_sheet(wb_new: openpyxl.Workbook, template_ws, rows: list[tuple[str, str, dict, dict]]):
    """rows: list of (province, specialty, cdcp_fees, pt_fees) processed in order."""
    ws = wb_new.create_sheet("DH")

    headers = _dh_header_row(template_ws)
    for col, name in enumerate(headers, start=1):
        src = template_ws.cell(1, col)
        dst = ws.cell(1, col, name)
        dst.font = copy.copy(src.font)
        dst.fill = copy.copy(src.fill)
        dst.alignment = copy.copy(src.alignment)
        dst.number_format = src.number_format

    style_row = 2  # representative data row in the template used for styling
    style_cells = {col: template_ws.cell(style_row, col) for col in range(1, len(headers) + 1)}

    r = 2
    for province, specialty, cdcp_fees, pt_fees in rows:
        for code in sorted(cdcp_fees.keys()):
            cdcp_2026 = cdcp_fees.get(code)
            pt_2026 = pt_fees.get(code)

            values = {
                "A": code,
                "B": specialty,
                "C": province,
                "D": f"=A{r}&C{r}&B{r}",
                "E": "N/A",  # 2025 CDCP fee: not available in current input files
                "F": cdcp_2026 if cdcp_2026 is not None else "N/A",
                "G": "N/A",  # 2025 PT fee: not available in current input files
                "H": pt_2026 if pt_2026 is not None else "N/A",
                "I": f'=IFERROR((F{r}-E{r})/E{r},"N/A")',
                "J": f'=IFERROR((H{r}-G{r})/G{r},"N/A")',
                "K": f'=IFERROR(F{r}/H{r},"N/A")',
                "L": 0,  # CDCP claim weight: claim-count data not available yet
                "M": f'=IF(H{r}="N/A",0,L{r})',
                "N": f"=IFERROR(M{r}/VLOOKUP(C{r},'Claim Lines'!A$26:F$38,6,FALSE),\"\")",
                "O": 0,  # Claim count nat'l: claim-count data not available yet
                "P": 0,  # PT weight: derived from claim counts, not available yet
                "Q": f"=IFERROR(P{r}*I{r},0)",
                "R": f"=O{r}/'Claim Lines'!F$25",
                "S": f"=IFERROR(R{r}*I{r},0)",
                "T": f'=IFERROR(K{r}*R{r},"N/A")',
                "U": f'=IFERROR(J{r}*R{r},"N/A")',
                "V": f'=IFERROR(J{r}*P{r},"N/A")',
                "W": f"=L{r}/'Claim Lines'!H$3",
                "X": f"=IFERROR(W{r}*I{r},0)",
                "Y": f'=IFERROR(N{r}*K{r},"")',
            }

            for col_idx, col_letter in enumerate("ABCDEFGHIJKLMNOPQRSTUVWXY", start=1):
                cell = ws.cell(r, col_idx, values[col_letter])
                style_src = style_cells[col_idx]
                cell.font = copy.copy(style_src.font)
                cell.fill = copy.copy(style_src.fill)
                cell.alignment = copy.copy(style_src.alignment)
                cell.number_format = style_src.number_format
            r += 1

    for col, dim in template_ws.column_dimensions.items():
        ws.column_dimensions[col].width = dim.width
    ws.freeze_panes = template_ws.freeze_panes
    if template_ws.row_dimensions[1].height:
        ws.row_dimensions[1].height = template_ws.row_dimensions[1].height


def build_claim_lines_sheet(wb_new: openpyxl.Workbook) -> None:
    """Minimal, self-contained Claim Lines sheet: only the cells the DH sheet's
    formulas actually reference (A26:F38 for VLOOKUP, F25, H3). Claim counts
    (columns L/O on the DH sheet) are currently zero, so these all evaluate to
    zero until real CDCP claim-count data is added."""
    ws = wb_new.create_sheet("Claim Lines")

    ws["A1"] = "For CDCP 2025 to 2026"
    ws["A3"] = "Total"
    ws["B3"] = 0
    ws["C3"] = 0
    ws["D3"] = "=B3+C3"
    ws["F3"] = "=SUM(DH!L:L)"
    ws["H3"] = "=SUM(D3:F3)"

    for i, prov in enumerate(ALL_PROVINCES):
        ws.cell(4 + i, 1, prov)

    ws["A23"] = ("For CDCP/PT Comparison (removing claim lines where the 2026 PT Fee Guide "
                 "or 2026 CDCP benefit grid does not have a fee)")
    ws["A25"] = "Total"
    ws["F25"] = "=SUM(DH!M:M)"

    for i, prov in enumerate(ALL_PROVINCES):
        row = 26 + i
        ws.cell(row, 1, prov)
        ws.cell(row, 6, f'=SUMIF(DH!C:C,"{prov}",DH!M:M)')


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


def main() -> None:
    specialty_dir = PT_GUIDES_DIR / SPECIALTY

    template_wb = openpyxl.load_workbook(TEMPLATE_WORKBOOK, data_only=False)
    template_ws = template_wb["DH"]

    wb_new = openpyxl.Workbook()
    wb_new.remove(wb_new.active)
    build_claim_lines_sheet(wb_new)

    rows = []
    total_codes = 0
    total_matched = 0
    for province in ALL_PROVINCES:
        cdcp_fees = load_cdcp_fees(province, SPECIALTY)
        if cdcp_fees is None or not cdcp_fees:
            print(f"{province}: no CDCP {SPECIALTY} data, skipped.")
            continue

        pt_fees, sources_used, files_found = load_pt_fees(
            specialty_dir, province, set(cdcp_fees.keys()), verbose=False
        )
        matched = len(pt_fees)
        total_codes += len(cdcp_fees)
        total_matched += matched

        source_desc = ", ".join(f"{name} ({n})" for name, n in sources_used) or "no PT fee guide found"
        print(f"{province}: {len(cdcp_fees)} CDCP codes, {matched} PT fees matched -- {source_desc}")
        if not files_found:
            print(f"    Note: no PT fee guide file found for {province}/{SPECIALTY}; PT fees will be 'N/A'.")

        report_plausibility(cdcp_fees, pt_fees)
        rows.append((province, SPECIALTY, cdcp_fees, pt_fees))

    build_dh_sheet(wb_new, template_ws, rows)

    wb_new.save(OUTPUT_WORKBOOK)
    total_rows = sum(len(cdcp_fees) for _, _, cdcp_fees, _ in rows)
    print(f"\nWrote {total_rows} rows across {len(rows)} provinces to {OUTPUT_WORKBOOK}")
    print(f"Overall PT fee match rate: {total_matched}/{total_codes} "
          f"({100 * total_matched / total_codes:.1f}%)")


if __name__ == "__main__":
    main()
