"""
Shared, format-agnostic helpers for extracting procedure-code -> fee mappings
out of PT/association fee guides, whatever format they happen to be in
(xlsx, xlsm, xls, csv, docx, or pdf).

The core idea: fee guide tables differ wildly in column layout from one
province/specialty to the next, but they share one property we can exploit --
each data row has a cell holding a recognizable procedure code and another
cell holding a recognizable fee. So instead of hardcoding column positions,
every loader here reduces its source to rows of raw cell values and a single
generic scanner (`extract_codes_from_rows`) looks for (code, fee) pairs,
using the set of procedure codes we actually care about (from the matching
CDCP price file) to avoid false positives on category headers/page numbers.
"""

import csv
import re
from pathlib import Path

import openpyxl
import pypdf
import xlrd
from docx import Document

_DOLLAR_RE = re.compile(r"\d[\d,]*(?:\.\d+)?")

SPREADSHEET_SUFFIXES = {".xlsx", ".xlsm", ".xls"}
CSV_SUFFIXES = {".csv"}
DOC_SUFFIXES = {".docx"}
PDF_SUFFIXES = {".pdf"}


def normalize_code(raw) -> str | None:
    """Normalize a procedure code (int or str, possibly missing leading zeros) to 5 digits.

    Codes are always whole numbers, so a fractional float (e.g. a $221.75 fee
    cell) is rejected rather than silently truncated -- otherwise a fee could
    coincidentally truncate to a real code (221.75 -> 221 -> "00221") and get
    misread as a code cell in an unrelated row.
    """
    if raw is None:
        return None
    if isinstance(raw, float) and not raw.is_integer():
        return None
    try:
        return f"{int(raw):05d}"
    except (ValueError, TypeError):
        s = str(raw).strip()
        return s if s.isdigit() and len(s) <= 6 else None


def extract_max_dollar(value) -> float | None:
    """Extract the largest dollar amount from a fee cell.

    Handles plain numbers, ranges ("$44.66 to $89.32"), and suffixed fees
    ("$56.02 + exp"). Non-numeric fees (e.g. "c.s." / client specific) return None.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    numbers = []
    for m in _DOLLAR_RE.findall(str(value)):
        try:
            numbers.append(float(m.replace(",", "")))
        except ValueError:
            continue
    return max(numbers) if numbers else None


def _fee_candidates(cell) -> list[float]:
    """Fee candidates for one non-code cell. Real numeric cells (the normal
    case for a spreadsheet fee column) are always trusted. For text cells,
    extract dollar-like numbers but ignore bare 5-digit tokens -- those are
    almost always a *different* procedure code mentioned in a description
    ("...see code 00616 below"), not a fee."""
    if isinstance(cell, (int, float)):
        return [float(cell)]
    if not isinstance(cell, str):
        return []
    values = []
    for token in _DOLLAR_RE.findall(cell):
        if re.fullmatch(r"\d{5}", token):
            continue
        try:
            values.append(float(token.replace(",", "")))
        except ValueError:
            continue
    return values


def extract_codes_from_rows(rows, known_codes: set[str]) -> dict[str, float]:
    """Generic (code, fee) scanner for row-shaped tabular data (list of lists of cell values).

    For each row, any cell that normalizes to a code in `known_codes` is treated
    as a procedure-code cell. The fee is taken from actual numeric-typed cells
    in that row if there are any (the reliable case for real spreadsheets);
    otherwise it falls back to a dollar-like number found in the other cells'
    text (for all-string sources like csv/docx). In both cases we take the
    *last* (rightmost) candidate rather than the largest: fee tables commonly
    have the fee as their rightmost numeric column, sometimes preceded by
    unrelated numeric columns (e.g. a multi-year price-escalation table where
    earlier columns hold smaller prior-year and unrounded intermediate
    values) where the largest number is not necessarily the current fee.
    """
    fees: dict[str, float] = {}
    for row in rows:
        cells = list(row)
        code_cells = []
        for i, cell in enumerate(cells):
            code = normalize_code(cell)
            if code and code in known_codes:
                code_cells.append((i, code))
        if not code_cells:
            continue
        for i, code in code_cells:
            if code in fees:
                continue
            other_cells = [cell for j, cell in enumerate(cells) if j != i]
            numeric_candidates = [float(c) for c in other_cells if isinstance(c, (int, float))]
            if numeric_candidates:
                fees[code] = numeric_candidates[-1]
                continue
            text_candidates = [v for cell in other_cells for v in _fee_candidates(cell)]
            if text_candidates:
                fees[code] = text_candidates[-1]
    return fees


def rows_from_spreadsheet(path: Path):
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        wb = openpyxl.load_workbook(path, data_only=True)
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                yield row
    elif suffix == ".xls":
        wb = xlrd.open_workbook(str(path))
        for sheet in wb.sheets():
            for row_idx in range(sheet.nrows):
                yield sheet.row_values(row_idx)
    else:
        raise ValueError(f"Unsupported spreadsheet suffix: {suffix}")


def rows_from_csv(path: Path):
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        yield from csv.reader(f)


def rows_from_docx_tables(path: Path):
    doc = Document(str(path))
    for table in doc.tables:
        for row in table.rows:
            yield [cell.text.strip() for cell in row.cells]


def docx_paragraph_text(path: Path) -> str:
    doc = Document(str(path))
    return "\n".join(p.text for p in doc.paragraphs)


# Fee-shaped tokens, tried in order of specificity: a dollar amount
# (optionally a range or "+ lab/exp" suffix), a bare two-decimal amount
# (some guides drop the "$"), a "client specific" / "c.s." marker (no fixed
# fee), and finally a bare whole number (some guides, e.g. Quebec's, list
# fees with no decimals at all -- riskiest, so tried last and only within a
# code's own text segment, see extract_codes_from_text).
_FEE_TOKEN_RES = [
    re.compile(r"\$[\d,]+\.\d{2}(?:\s*(?:to|-)\s*\$?[\d,]+\.\d{2})?(?:\s*\+\s*[A-Za-z]+)*"),
    re.compile(r"\b[\d,]+\.\d{2}\b(?:\s*(?:to|-)\s*\$?[\d,]+\.\d{2})?(?:\s*\+\s*[A-Za-z]+)*"),
    re.compile(r"c\.?\s*s\.?\s*\(?client specific\)?\.?", re.IGNORECASE),
    re.compile(r"c\.?\s*s\.?", re.IGNORECASE),
    re.compile(r"\b\d{1,4}\b"),
]

_ANY_CODE_RE = re.compile(r"\b\d{5}\b")

# How far into a code's text segment to look for its fee. Keeps the
# whole-number fallback tier from wandering into an unrelated number many
# sentences later.
_SEGMENT_SEARCH_WINDOW = 700


_YEAR_RE = re.compile(r"^(19|20)\d{2}$")


def _fee_token_in_segment(segment: str) -> float | None:
    """Find the fee within one code's text segment.

    For the specific tiers (dollar amounts, decimals, "c.s." markers) the
    fee is the *first* match -- these guides put the fee right after the
    code ("00111 $63.80..."), or it's the only such token in a descriptive
    paragraph. The bare-whole-number fallback tier is different: guides that
    use it often prefix the fee with a quantity label ("00211 1 image 46"),
    so there the fee is the *last* match instead.
    """
    window = segment[:_SEGMENT_SEARCH_WINDOW]
    for pattern in _FEE_TOKEN_RES[:-1]:
        for match in pattern.finditer(window):
            text = match.group(0)
            if _YEAR_RE.match(text.strip()):
                continue
            return extract_max_dollar(text)

    last_valid = None
    for match in _FEE_TOKEN_RES[-1].finditer(window):
        text = match.group(0)
        if _YEAR_RE.match(text.strip()):
            continue
        last_valid = text
    return extract_max_dollar(last_valid) if last_valid is not None else None


# Words that typically precede a *cross-reference* to another code inside a
# description ("...as per 00100", "...see code 00616", "...use code 00616")
# rather than that code's own entry. Occurrences preceded by one of these are
# ignored entirely -- both as segment boundaries and as extraction points --
# so a mention of another code doesn't truncate the current entry's segment,
# and a mention of *this* code doesn't get mistaken for its own definition.
_CROSS_REF_CUE_WORDS = {
    "per", "code", "codes", "see", "use", "using", "refer",
    "of", "or", "and", "lieu", "than", "instead", "not",
}
_PRECEDING_WORD_RE = re.compile(r"([A-Za-z]+)\s*$")
_FOLLOWED_BY_PERIOD_RE = re.compile(r"^\s?\.")


def _is_cross_reference(text: str, pos: int) -> bool:
    preceding = _PRECEDING_WORD_RE.search(text[max(0, pos - 30):pos])
    if preceding and preceding.group(1).lower() in _CROSS_REF_CUE_WORDS:
        return True
    # A code immediately followed by a sentence-ending period ("...listed in
    # 00100.") reads as a citation closing out someone else's sentence, not
    # the start of this code's own entry (which is normally followed by more
    # descriptive text, not punctuation).
    end = pos + 5
    return bool(_FOLLOWED_BY_PERIOD_RE.match(text[end:end + 2]))


def extract_codes_from_text(text: str, known_codes: set[str]) -> dict[str, float]:
    """Segment `text` by occurrences of *any* 5-digit code (not just ones we
    care about), then look for a fee token within each known code's segment
    (the text up to the next code of any kind). Bounding on any code --
    rather than only known ones -- keeps an unrelated nearby entry's numbers
    (for a code outside our CDCP list) from bleeding into the segment.
    Cross-reference mentions of a code within another entry's description are
    excluded from consideration entirely (see _is_cross_reference)."""
    all_matches = [
        m for m in _ANY_CODE_RE.finditer(text) if not _is_cross_reference(text, m.start())
    ]

    fees: dict[str, float] = {}
    for i, m in enumerate(all_matches):
        code = m.group(0)
        if code not in known_codes or code in fees:
            continue
        next_start = all_matches[i + 1].start() if i + 1 < len(all_matches) else len(text)
        segment = text[m.end():next_start]
        fee = _fee_token_in_segment(segment)
        if fee is not None:
            fees[code] = fee
    return fees


def load_fees_from_pdf(path: Path, known_codes: set[str]) -> dict[str, float]:
    reader = pypdf.PdfReader(str(path))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    return extract_codes_from_text(text, known_codes)


def load_fees_from_docx(path: Path, known_codes: set[str]) -> dict[str, float]:
    fees = extract_codes_from_rows(rows_from_docx_tables(path), known_codes)
    missing = known_codes - fees.keys()
    if missing:
        # Some docx fee guides are prose/paragraphs rather than tables.
        fees.update({
            code: fee
            for code, fee in extract_codes_from_text(docx_paragraph_text(path), known_codes).items()
            if code not in fees
        })
    return fees


def load_fees_from_spreadsheet(path: Path, known_codes: set[str]) -> dict[str, float]:
    return extract_codes_from_rows(rows_from_spreadsheet(path), known_codes)


def load_fees_from_csv(path: Path, known_codes: set[str]) -> dict[str, float]:
    return extract_codes_from_rows(rows_from_csv(path), known_codes)


PROVINCE_ALIASES: dict[str, list[str]] = {
    "PE": ["PE", "PEI"],
    "YK": ["YK", "YT", "YU", "YUKON"],
    "NT": ["NT", "NWT"],
}


def discover_pt_files(specialty_dir: Path, province: str) -> list[Path]:
    """Find every file under `specialty_dir` (recursively) whose name starts
    with the given province's abbreviation (or a known alias)."""
    aliases = PROVINCE_ALIASES.get(province, [province])
    patterns = [re.compile(rf"^{re.escape(a)}\b", re.IGNORECASE) for a in aliases]

    matches = []
    for path in specialty_dir.rglob("*"):
        if not path.is_file():
            continue
        if any(p.match(path.name) for p in patterns):
            matches.append(path)
    return matches


def _is_english(path: Path) -> bool:
    name = path.stem.upper()
    return "FR" not in name.split() and "FRENCH" not in name and "- FR" not in name.upper()


def load_pt_fees(specialty_dir: Path, province: str, known_codes: set[str], verbose: bool = True):
    """Resolve PT fees for a province/specialty from whatever files are available.

    Tries, in priority order (filling in only the codes still missing at each
    step, so multiple partial sources combine): spreadsheets (xlsx/xlsm/xls),
    csv, docx, then pdf (English files before French, if language is discernible).
    Returns (fees dict, list of (source_description, codes_found) used).
    """
    files = discover_pt_files(specialty_dir, province)
    fees: dict[str, float] = {}
    sources_used: list[tuple[str, int]] = []

    def _apply(label: str, new_fees: dict[str, float]):
        added = {c: f for c, f in new_fees.items() if c not in fees}
        fees.update(added)
        if added:
            sources_used.append((label, len(added)))

    spreadsheets = sorted(
        (f for f in files if f.suffix.lower() in SPREADSHEET_SUFFIXES),
        key=lambda f: f.suffix.lower() != ".xlsx",  # prefer .xlsx over .xlsm/.xls
    )
    for f in spreadsheets:
        try:
            _apply(f.name, load_fees_from_spreadsheet(f, known_codes))
        except Exception as e:
            if verbose:
                print(f"    WARNING: failed to read {f.name}: {e}")

    for f in (f for f in files if f.suffix.lower() in CSV_SUFFIXES):
        if known_codes - fees.keys():
            try:
                _apply(f.name, load_fees_from_csv(f, known_codes))
            except Exception as e:
                if verbose:
                    print(f"    WARNING: failed to read {f.name}: {e}")

    for f in (f for f in files if f.suffix.lower() in DOC_SUFFIXES):
        if known_codes - fees.keys():
            try:
                _apply(f.name, load_fees_from_docx(f, known_codes))
            except Exception as e:
                if verbose:
                    print(f"    WARNING: failed to read {f.name}: {e}")

    pdfs = sorted(
        (f for f in files if f.suffix.lower() in PDF_SUFFIXES),
        key=lambda f: not _is_english(f),  # English first
    )
    for f in pdfs:
        if known_codes - fees.keys():
            try:
                _apply(f.name, load_fees_from_pdf(f, known_codes))
            except Exception as e:
                if verbose:
                    print(f"    WARNING: failed to read {f.name}: {e}")

    return fees, sources_used, files
