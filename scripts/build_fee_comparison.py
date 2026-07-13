"""
Consolidate CDCP price file data and PT (provincial/territorial) association fee
guide data into a "Fee Comparison" workbook, matching the shape of
Data/Output_2026 Fee Comparison/2026 Fee Comparisons v2.xlsx.

Current scope: Ontario (ON), Dental Hygienist (DH) only.

Inputs:
  - Data/Input_CDCP price files/2026 - CDCP PRICE FILE - ON.xlsx
  - Data/Input_PT association fee guides/DH/ON DH Fee Guide 2026.xlsx  (primary PT fee source)
  - Data/Input_PT association fee guides/DH/ON DH Fee Guide 2026.pdf  (cross-check only)

Notes on known gaps (see conversation / README at top of repo for rationale):
  - Neither input file contains 2025 fees, so '2025 CDCP Fee' and '2025 PT Fee'
    are written as "N/A".
  - CDCP claim-count data (used to compute weighted increases in columns I-Y)
    is not available from these three files, so claim-count columns (L, O, P)
    are written as 0. The formulas are still written so the workbook is ready
    to be populated with real claim counts later.

Output:
  - Data/Output_2026 Fee Comparison/2026 Fee Comparisons - generated.xlsx
"""

import copy
import re
from pathlib import Path

import openpyxl
import pypdf

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "Data"

CDCP_PRICE_FILE = DATA_DIR / "Input_CDCP price files" / "2026 - CDCP PRICE FILE - ON.xlsx"
PT_FEE_GUIDE_XLSX = DATA_DIR / "Input_PT association fee guides" / "DH" / "ON DH Fee Guide 2026.xlsx"
PT_FEE_GUIDE_PDF = DATA_DIR / "Input_PT association fee guides" / "DH" / "ON DH Fee Guide 2026.pdf"
TEMPLATE_WORKBOOK = DATA_DIR / "Output_2026 Fee Comparison" / "2026 Fee Comparisons v2.xlsx"
OUTPUT_WORKBOOK = DATA_DIR / "Output_2026 Fee Comparison" / "2026 Fee Comparisons - generated.xlsx"

PROVINCE = "ON"
SPECIALTY = "DH"

CLAIM_LINES_PROVINCES = ["AB", "BC", "MB", "NB", "NL", "NS", "ON", "PE", "QC", "SK", "NT", "NU", "YT"]

_DOLLAR_RE = re.compile(r"[\d,]+\.?\d*")


def normalize_code(raw) -> str | None:
    """Normalize a procedure code (int or str, possibly missing leading zeros) to 5 digits."""
    if raw is None:
        return None
    try:
        return f"{int(raw):05d}"
    except (ValueError, TypeError):
        s = str(raw).strip()
        return s if s.isdigit() else None


def extract_max_dollar(value) -> float | None:
    """Extract the largest dollar amount from a fee cell.

    Handles plain numbers, ranges ("$44.66 to $89.32"), and suffixed fees
    ("$56.02 + exp"). Non-numeric fees (e.g. "c.s." / client specific) return None.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    matches = _DOLLAR_RE.findall(str(value))
    numbers = [float(m.replace(",", "")) for m in matches if m and m != "."]
    return max(numbers) if numbers else None


def load_cdcp_fees(path: Path, province: str, specialty: str) -> dict[str, float]:
    """Read 2026 CDCP provider fees per procedure code from a CDCP price file."""
    wb = openpyxl.load_workbook(path, data_only=True)
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


def load_pt_fees_xlsx(path: Path, sheet_name: str = "Table 1") -> dict[str, float]:
    """Read 2026 PT association suggested fees per procedure code from the fee guide xlsx."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name]

    fees: dict[str, float] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 3:
            continue
        code = normalize_code(row[1])
        if code is None:
            continue
        fee = extract_max_dollar(row[2])
        if fee is not None:
            fees[code] = fee
    return fees


_PDF_FEE_RE = re.compile(
    r"\b(\d{5})\b\s+"
    r"(\$[\d,]+\.\d{2}(?:\s*(?:to|-)\s*\$?[\d,]+\.\d{2})?(?:\s*\+\s*[A-Za-z]+)*"
    r"|c\.?\s*s\.?\s*\(?client specific\)?\.?"
    r"|c\.?\s*s\.?)",
    re.IGNORECASE,
)


def load_pt_fees_pdf(path: Path) -> dict[str, float]:
    """Best-effort extraction of procedure code -> fee pairs from the PDF fee guide.

    Used only as a cross-check against the xlsx fee guide, since PDF text
    extraction can split a code and its fee across lines when the source
    table wraps a row (this happens for a handful of codes in practice).
    """
    reader = pypdf.PdfReader(str(path))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)

    fees: dict[str, float] = {}
    for match in _PDF_FEE_RE.finditer(text):
        code = match.group(1)
        fee = extract_max_dollar(match.group(2))
        if fee is not None:
            fees[code] = fee
    return fees


def cross_check_pt_fees(primary: dict[str, float], secondary: dict[str, float], codes_of_interest) -> None:
    """Log any disagreements between the xlsx and PDF fee guide for the codes we actually use."""
    mismatches = []
    missing_from_pdf = []
    for code in codes_of_interest:
        xlsx_fee = primary.get(code)
        pdf_fee = secondary.get(code)
        if xlsx_fee is None:
            continue
        if pdf_fee is None:
            missing_from_pdf.append(code)
        elif abs(xlsx_fee - pdf_fee) > 0.005:
            mismatches.append((code, xlsx_fee, pdf_fee))

    if mismatches:
        print(f"  WARNING: {len(mismatches)} code(s) differ between xlsx and PDF fee guides:")
        for code, xf, pf in mismatches:
            print(f"    {code}: xlsx=${xf:.2f} pdf=${pf:.2f}")
    if missing_from_pdf:
        print(f"  Note: {len(missing_from_pdf)} code(s) not confidently extracted from PDF "
              f"(likely a wrapped table row in the source PDF): {', '.join(missing_from_pdf)}")
    if not mismatches and not missing_from_pdf:
        print("  PDF cross-check: all codes agree with the xlsx fee guide.")


def _dh_header_row(template_ws) -> list[str]:
    return [c.value for c in next(template_ws.iter_rows(min_row=1, max_row=1))]


def build_dh_sheet(wb_new: openpyxl.Workbook, template_ws, codes: list[str],
                    cdcp_fees: dict[str, float], pt_fees: dict[str, float],
                    province: str, specialty: str) -> None:
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

    for r, code in enumerate(codes, start=2):
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

        for col_idx, col_letter in enumerate(
            "ABCDEFGHIJKLMNOPQRSTUVWXY", start=1
        ):
            cell = ws.cell(r, col_idx, values[col_letter])
            style_src = style_cells[col_idx]
            cell.font = copy.copy(style_src.font)
            cell.fill = copy.copy(style_src.fill)
            cell.alignment = copy.copy(style_src.alignment)
            cell.number_format = style_src.number_format

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

    for i, prov in enumerate(CLAIM_LINES_PROVINCES):
        row = 4 + i
        ws.cell(row, 1, prov)

    ws["A23"] = ("For CDCP/PT Comparison (removing claim lines where the 2026 PT Fee Guide "
                 "or 2026 CDCP benefit grid does not have a fee)")
    ws["A25"] = "Total"
    ws["F25"] = "=SUM(DH!M:M)"

    for i, prov in enumerate(CLAIM_LINES_PROVINCES):
        row = 26 + i
        ws.cell(row, 1, prov)
        ws.cell(row, 6, f'=SUMIF(DH!C:C,"{prov}",DH!M:M)')


def main() -> None:
    print(f"Loading CDCP {SPECIALTY} fees for {PROVINCE} from {CDCP_PRICE_FILE.name}")
    cdcp_fees = load_cdcp_fees(CDCP_PRICE_FILE, PROVINCE, SPECIALTY)
    print(f"  {len(cdcp_fees)} procedure codes with a 2026 CDCP fee.")

    print(f"Loading PT {SPECIALTY} fees for {PROVINCE} from {PT_FEE_GUIDE_XLSX.name}")
    pt_fees_xlsx = load_pt_fees_xlsx(PT_FEE_GUIDE_XLSX)
    print(f"  {len(pt_fees_xlsx)} procedure codes with a 2026 PT fee.")

    print(f"Cross-checking against {PT_FEE_GUIDE_PDF.name}")
    pt_fees_pdf = load_pt_fees_pdf(PT_FEE_GUIDE_PDF)
    cross_check_pt_fees(pt_fees_xlsx, pt_fees_pdf, cdcp_fees.keys())

    codes = sorted(cdcp_fees.keys())
    matched = sum(1 for c in codes if c in pt_fees_xlsx)
    print(f"Matched PT fee for {matched}/{len(codes)} CDCP procedure codes "
          f"({len(codes) - matched} will show 'N/A' PT fee).")

    template_wb = openpyxl.load_workbook(TEMPLATE_WORKBOOK, data_only=False)
    template_ws = template_wb["DH"]

    wb_new = openpyxl.Workbook()
    wb_new.remove(wb_new.active)

    build_claim_lines_sheet(wb_new)
    build_dh_sheet(wb_new, template_ws, codes, cdcp_fees, pt_fees_xlsx, PROVINCE, SPECIALTY)

    wb_new.save(OUTPUT_WORKBOOK)
    print(f"\nWrote {len(codes)} rows to {OUTPUT_WORKBOOK}")


if __name__ == "__main__":
    main()
