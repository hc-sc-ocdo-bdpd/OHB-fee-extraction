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
from docx.oxml.ns import qn
from docx.table import Table as _DocxTable
from docx.text.paragraph import Paragraph as _DocxParagraph

_DOLLAR_RE = re.compile(r"\d[\d,]*(?:\.\d+)?")

# French-Canadian formatted amount within a spreadsheet text cell (space as
# thousands separator, comma as decimal separator, trailing "$" -- e.g.
# "1 509 $ + L" in QC's SP guide). Matched and removed from the cell *before*
# the plain-digit _DOLLAR_RE pass in _fee_candidates below, which would
# otherwise split "1 509" into separate "1" and "509" tokens (fragmenting
# the thousands digit and, worse, silently discarding it as looking like a
# quantity) rather than reading the number as 1509.
_FRENCH_GROUPED_SPACE_RE = re.compile(r"(?<!\d)\d{1,3}(?:[\s ]\d{3})+(?:[.,]\d{2})?\s*\$")

# Same French-Canadian space-grouped thousands, but without a trailing "$" --
# seen in QC's GP guide, where a cell just reads "1 261" with no dollar sign
# at all. Without this, "1 261" falls all the way through to the plain-digit
# _DOLLAR_RE pass below, which has no notion of a space as a thousands
# separator and so splits it into two unrelated tokens, "1" and "261" -- and
# since callers take the *last* candidate as the fee, the leading digit(s)
# get silently dropped (a real case: code 27200's true fee of 1261 was
# extracted as 261). Guarded on both sides with a "not adjacent to another
# digit" check so it can't be carved out of an unrelated longer digit run --
# e.g. two 5-digit procedure codes separated only by a space ("01120 01130")
# must NOT be misread as a single grouped number.
_FRENCH_GROUPED_NO_DOLLAR_RE = re.compile(r"(?<!\d)\d{1,3}(?:[\s ]\d{3})+(?:[.,]\d{2})?(?!\d)")

# "S.C." ("Service Charge"/"Independent Charge", per NB's DD guide's own
# abbreviations legend) is yet another no-fixed-fee marker, like "I.C." and
# "c.s." -- but with the two letters in the opposite order, so it does NOT
# match the "c.s." alternative above despite meaning the same thing. Without
# its own alternative here, a cell/segment whose only content is "S.C."
# isn't recognized as a marker at all, so has_no_fee_marker (below, and in
# _FEE_TOKEN_TIERS) never fires for it and it just resolves to nothing.
# "B.R." ("By Report") is AB's DD guide's own version of the same idea --
# the fee is individually assessed and reported, not fixed.
_NO_FEE_MARKER_RE = re.compile(
    r"^\s*(?:I\.?\s*C\.?|c\.?\s*s\.?|s\.?\s*c\.?|(?:actual\s+)?lab(?:\s+fee)?|B\.?\s*R\.?)\s*\.?\s*$",
    re.IGNORECASE,
)

SPREADSHEET_SUFFIXES = {".xlsx", ".xlsm", ".xls"}
CSV_SUFFIXES = {".csv"}
DOC_SUFFIXES = {".docx"}
PDF_SUFFIXES = {".pdf"}


_ALPHA_CODE_RE = re.compile(r"^[A-Z]\d{3,5}$")


def normalize_code(raw) -> str | None:
    """Normalize a procedure code (int or str, possibly missing leading zeros) to 5 digits.

    Codes are always whole numbers, so a fractional float (e.g. a $221.75 fee
    cell) is rejected rather than silently truncated -- otherwise a fee could
    coincidentally truncate to a real code (221.75 -> 221 -> "00221") and get
    misread as a code cell in an unrelated row.

    A handful of GP/SP codes are alphanumeric (e.g. "P0500") rather than
    purely numeric -- one letter followed by 3-5 digits is also accepted.

    A leading "*" is stripped before matching -- BC's DH guide flags a code
    with "*00616" at its real, priced entry (the asterisk is a footnote
    marker referencing a relocation note elsewhere), then repeats the bare
    code "00616" again later at a stub cross-reference row ("* moved after
    00611 for clarity") with no real fee, just placeholder zeros. Without
    stripping the "*", the starred row's code is never recognized at all, so
    the scanner falls through to the stub row instead and reports its
    placeholder 0 as the fee.
    """
    if raw is None:
        return None
    if isinstance(raw, float) and not raw.is_integer():
        return None
    try:
        return f"{int(raw):05d}"
    except (ValueError, TypeError):
        s = str(raw).strip().upper().lstrip("*").strip()
        if s.isdigit() and len(s) <= 6:
            return s
        if _ALPHA_CODE_RE.match(s):
            return s
        return None


def _marker_text(value: str) -> str:
    """Identity parser for the no-fixed-fee marker tiers ("I.C.", "c.s.",
    "S.C.") -- returns the exact matched text as-is (whitespace-trimmed only),
    rather than resolving it to a number or dropping it. A code with no
    fixed fee is real, meaningful information (the source is telling us the
    fee is individually costed / client-specific / a service charge), so the
    reference sheet should show that literal text instead of a generic
    "N/A", which would look the same as "the extractor found nothing at
    all"."""
    return value.strip()


def _is_zero_padded_numeric_code(raw) -> bool:
    """True if `raw` is a *number* that normalize_code would have to pad with
    a leading zero to reach 5 digits (e.g. 2116 -> "02116"). Such a cell is
    far more likely a fee than a procedure code: a code stored as text keeps
    its leading zero ("02116"), so only genuinely 5-digit-valued numbers
    (71209) read as codes without padding. See extract_codes_from_rows for
    how this disambiguates a row that appears to hold two different codes."""
    if isinstance(raw, bool) or not isinstance(raw, (int, float)):
        return False
    if isinstance(raw, float) and not raw.is_integer():
        return False
    return 0 <= int(raw) < 10000


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
    ("...see code 00616 below"), not a fee -- and ignore small whole numbers
    with no decimal point, since those are usually a size/quantity/duration
    mentioned in a description ("Scar Tissue - 1 - 2 cm", "1 unit of time")
    rather than a fee. This matters most for all-text sources like csv/docx,
    where every cell is a string and there's no numeric type to fall back on
    to distinguish a real fee cell from a description cell."""
    if isinstance(cell, (int, float)):
        return [float(cell)]
    if not isinstance(cell, str):
        return []
    # Collected as (position, value) and sorted at the end so a cell with
    # both kinds of amount in it (e.g. a French range "949 $ - 1 369 $")
    # comes back in left-to-right reading order -- callers take the *last*
    # candidate as the fee (see extract_codes_from_rows), so losing that
    # order would silently swap which end of a range wins.
    matches: list[tuple[int, float]] = []
    masked = cell
    for m in _FRENCH_GROUPED_SPACE_RE.finditer(cell):
        parsed = _parse_french_amount(m.group(0))
        if parsed is not None:
            matches.append((m.start(), parsed))
        # Blank out the matched span (same length, so positions of
        # everything else are unaffected) so the plain-digit pass below
        # doesn't also re-match the digits inside it.
        masked = masked[:m.start()] + " " * (m.end() - m.start()) + masked[m.end():]
    for m in _FRENCH_GROUPED_NO_DOLLAR_RE.finditer(masked):
        parsed = _parse_french_amount(m.group(0))
        if parsed is not None:
            matches.append((m.start(), parsed))
        masked = masked[:m.start()] + " " * (m.end() - m.start()) + masked[m.end():]
    for m in _DOLLAR_RE.finditer(masked):
        token = m.group(0)
        digits = token.replace(",", "")
        # Check the 5-digit-code exclusion against the comma-stripped form,
        # not the raw token: _DOLLAR_RE's character class treats a comma as
        # part of the number, so a code-like reference immediately followed
        # by a comma in prose ("...classes 30000, 40000, and 70000.") comes
        # out as "30000," -- which doesn't literally fullmatch \d{5} (it has
        # 6 characters) even though it's the exact same 5-digit number this
        # exclusion exists to catch.
        if re.fullmatch(r"\d{5}", digits):
            continue
        try:
            value = float(digits)
        except ValueError:
            continue
        if "." not in token and value < 10:
            continue
        matches.append((m.start(), value))
    matches.sort(key=lambda t: t[0])
    return [v for _, v in matches]


# Column headers containing one of these words are treated as a "fee"
# column when looking for which cell in a row holds the fee -- this avoids
# being misled by an unrelated numeric column elsewhere in the row (e.g. a
# sequential "ID"/row-number column) when the *real* fee is stored as
# formatted text ("$89.80") rather than a native number, which otherwise
# defeats the "prefer real numeric cells" heuristic below.
_FEE_HEADER_RE = re.compile(r"fee|amount|price|tarif|montant|\bfrom\b|\brate\b|\bcost\b|\bcharge\b", re.IGNORECASE)

# Column headers containing "code" are treated as *the* code column when
# present. Without this, a fee value that happens to zero-pad to a real
# procedure code (e.g. a $1,802 fee coincidentally matching code "01802")
# gets misread as a reference to that other code -- restricting code
# detection to a labeled code column, when one is identifiable, avoids that
# numeric coincidence entirely.
_CODE_HEADER_RE = re.compile(r"\bcode\b", re.IGNORECASE)


def _looks_like_header(row) -> bool:
    """True if `row` looks like a header row: several short, densely-packed
    text labels, no cell that could plausibly be a procedure code (a real
    header row won't have a 5-digit-normalizable cell -- code columns are
    headed by text like "Code"). Requires at least 2 populated cells and
    short cell text, so a single-cell title/banner row (common as row 1 in
    these fee guides, e.g. "NLDHA ... 2026 Fee Guide ...") isn't mistaken
    for a real header just because it's textual and happens to contain a
    word like "Fee" -- that would misdirect the fee-column search entirely.
    """
    cells = [c for c in row if c is not None]
    if len(cells) < 2:
        return False
    if any(normalize_code(c) for c in cells):
        return False
    if any(isinstance(c, str) and len(c) > 60 for c in cells):
        return False
    if sum(isinstance(c, str) for c in cells) / len(cells) <= 0.5:
        return False
    return len(cells) / len(row) >= 0.3


_EXACT_FEE_HEADER_NAMES = {"fee", "price", "amount", "tarif", "montant", "rate", "cost", "from"}


def find_fee_column_indices(header) -> list[int]:
    """Prefer a column whose header is *exactly* a fee-like word (e.g. "Fee")
    over one that merely contains one as a substring (e.g. "UpperFee",
    "InternalLabFee"). Some guides have both a base fee column and a
    secondary upper-bound-of-range column ("Fee" / "UpperFee") -- the base
    "Fee" column is what the reference file treats as the canonical fee, so
    an exact match should win outright rather than being merged in with (and
    potentially outranked by, since the rightmost match wins) the range
    column.

    A range column pair headed exactly "From"/"To" is the opposite case,
    though (seen in one ON GP sheet): the reference treats the ceiling
    ("To") as the fee, not the floor ("From"). "To" isn't in
    _EXACT_FEE_HEADER_NAMES on its own -- alone it's too generic a word and
    wrongly outranks a real fuzzy-matched fee column in other sheets (one ON
    OS sheet has "Suggested Fee " / "To", where "To" is just an unused
    vestigial column and the real data lives in "Suggested Fee", which would
    get wrongly excluded if "To" won as an exact match by itself) -- so it's
    only added here, and only when "From" is also an exact match in the same
    header, confirming a genuine paired range rather than a stray "To".

    A "Fee" / "UpperFee" pair is the SAME "take the ceiling" case, not the
    opposite one this docstring used to claim: confirmed against AB's guide,
    where a fracture-reduction code's "Fee" is $1,357.70 and "UpperFee" is
    $1,697.10, and $1,697.10 is the correct fee -- consistent with how a
    range is read everywhere else in this project (extract_max_dollar always
    takes the largest end). "UpperFee" is appended after the exact "Fee"
    match (only when "Fee" itself is an exact match, narrowly targeting the
    word "upper" so it doesn't pull in some unrelated column that merely
    contains "fee") so it still wins as the rightmost candidate even though
    "Fee" alone would otherwise be returned outright."""
    exact = [i for i, h in enumerate(header) if h and str(h).strip().lower() in _EXACT_FEE_HEADER_NAMES]
    if exact and any(header[i] and str(header[i]).strip().lower() == "from" for i in exact):
        exact += [i for i, h in enumerate(header) if h and str(h).strip().lower() == "to"]
    elif exact and any(header[i] and str(header[i]).strip().lower() == "fee" for i in exact):
        # Require the header to be essentially just "upperfee"/"upper fee",
        # not merely *contain* "upper" somewhere -- a broader substring
        # match here turned out to sweep in unrelated columns (e.g. some
        # other guide's "Upper Age Limit" note) that aren't a fee range at
        # all, causing several other provinces' GP sheets to regress once
        # this rule went in.
        exact += [i for i, h in enumerate(header)
                  if h and i not in exact and re.sub(r"\s+", "", str(h).strip().lower()) == "upperfee"]
    fee_cols = exact if exact else [i for i, h in enumerate(header) if h and _FEE_HEADER_RE.search(str(h))]

    # A fee-like column immediately followed by a column with *no* header
    # text at all, or literally headed "To", is very likely an unlabeled or
    # lightly-labeled ceiling-of-range companion -- seen in several ON SP
    # guides: "Fee" / <blank> in the EN and PE guides (confirmed against
    # real codes, e.g. 01204, where the blank column held the reference's
    # actual expected fee while "Fee" held a lower, superseded value), and
    # "Suggested Fee " / "To" in the OS guide, which an earlier version of
    # this function dismissed as "just an unused vestigial column" -- that
    # was wrong for at least those same codes. Only the column immediately
    # adjacent is considered, not a blank column found anywhere else in the
    # header, so an unrelated blank formatting column elsewhere in the
    # sheet isn't swept in.
    for i in list(fee_cols):
        if i + 1 < len(header) and i + 1 not in fee_cols:
            companion = header[i + 1]
            is_blank = companion is None or (isinstance(companion, str) and not companion.strip())
            is_to = isinstance(companion, str) and companion.strip().lower() == "to"
            if is_blank or is_to:
                fee_cols.append(i + 1)
    return fee_cols


def find_code_column_indices(header) -> list[int]:
    return [i for i, h in enumerate(header) if h and _CODE_HEADER_RE.search(str(h))]


def extract_codes_from_rows(
    tables, known_codes: set[str], target_specialty: str | None = None
) -> dict[str, float | str]:
    """Generic (code, fee) scanner for row-shaped tabular data.

    `tables` is an iterable of (header, data_rows) pairs -- one per
    sheet/table, header may be None if none was detected. For each data row,
    any cell that normalizes to a code in `known_codes` is treated as a
    procedure-code cell. If the table has an identifiable "fee" column (by
    header name), only that column is considered -- this is what keeps an
    unrelated numeric column (e.g. a sequential ID) from being mistaken for
    the fee when the real fee is stored as text. Otherwise, the fee is taken
    from actual numeric-typed cells in the row if there are any (the
    reliable case for real spreadsheets), else from a dollar-like number
    found in the other cells' text (for all-string sources like csv/docx).
    In all cases we take the *last* (rightmost) candidate rather than the
    largest: fee tables commonly have the fee as their rightmost relevant
    column, sometimes preceded by unrelated numeric columns (e.g. a
    multi-year price-escalation table where earlier columns hold smaller
    prior-year and unrounded intermediate values) where the largest number
    is not necessarily the current fee.

    `target_specialty`, if given (a CDCP sub-specialty code like "EN"), asks
    the scanner to prefer a row whose own specialty-labeled column (see
    find_row_specialty_column) matches that sub-specialty when the same code
    appears more than once in the source under different specialties (e.g.
    PE's combined GP+SP guide lists code 25781 once under "GP" at $101 and
    again under "END" at $139.30) -- without it, the first row encountered
    always wins regardless of which specialty it actually belongs to.
    """
    fees: dict[str, float | str] = {}
    # code -> [(specialty_label_or_None, fee), ...], only populated when
    # target_specialty is given, since the single-pass "first row wins" path
    # above is enough (and already validated) for every other caller.
    labeled_candidates: dict[str, list[tuple[str | None, float]]] = {}

    for header, rows in tables:
        fee_col_indices = find_fee_column_indices(header) if header else []
        code_col_indices = find_code_column_indices(header) if header else []
        candidate_code_col_indices = code_col_indices or None  # None = scan every column
        specialty_col_idx = find_row_specialty_column(rows) if target_specialty is not None else None
        for row in rows:
            cells = list(row)
            code_cells = []
            for i, cell in enumerate(cells):
                if candidate_code_col_indices is not None and i not in candidate_code_col_indices:
                    continue
                code = normalize_code(cell)
                if code and code in known_codes:
                    code_cells.append((i, code))
            if not code_cells:
                continue
            # A numeric cell only reaching 5 digits via zero-padding (raw
            # value under 10000, e.g. the number 2116 -> "02116") is
            # ambiguous: it's usually this row's *fee*, coincidentally
            # resembling some unrelated procedure code. When the row also
            # holds a code that needed no padding, that unpadded one is the
            # row's real code and the padded lookalikes are dropped --
            # otherwise the row gets harvested a second time under the
            # bogus code, with some other column standing in as its "fee".
            # Confirmed against SK's SP guide, whose row for code 75303
            # ("...2116") was being recorded as code 02116 priced at $25 --
            # 25 being the row's numeric *specialty* column, the only other
            # number left once the real code and the fee were excluded.
            # Rows where EVERY code-shaped cell is padded are left alone: a
            # source legitimately storing "01011" as the number 1011 in its
            # code column has no unpadded alternative to prefer. The
            # unpadded alternative is looked for across the whole row, not
            # just among code_cells, since the row's real code is often one
            # this particular call isn't even asking about (75303 above is
            # not in `known_codes` when only 02116 was requested).
            if any(not _is_zero_padded_numeric_code(cells[i]) for i, _ in code_cells):
                code_cells = [(i, c) for i, c in code_cells
                              if not _is_zero_padded_numeric_code(cells[i])]
            elif any(
                normalize_code(cell) is not None and not _is_zero_padded_numeric_code(cell)
                for j, cell in enumerate(cells)
                if candidate_code_col_indices is None or j in candidate_code_col_indices
            ):
                code_cells = []
            if not code_cells:
                continue
            # Exclude every cell that restates *this same* code elsewhere in
            # the row (e.g. a truncated numeric id "1011" alongside the full
            # code "01011") -- but NOT cells that happen to match a
            # *different* known code, since that's usually just this row's
            # fee value coincidentally zero-padding to resemble some other,
            # unrelated procedure code (e.g. a $2,116 fee for code 75303
            # looks like code "02116" if that also happens to be a real
            # code) rather than an actual reference to it.
            for i, code in code_cells:
                if target_specialty is None and code in fees:
                    continue
                exclude_indices = {j for j, c in code_cells if c == code} | {i}
                if fee_col_indices:
                    other_cells = [cells[j] for j in fee_col_indices
                                   if j not in exclude_indices and j < len(cells)]
                else:
                    other_cells = [cell for j, cell in enumerate(cells) if j not in exclude_indices]

                fee = None
                numeric_candidates = [float(c) for c in other_cells if isinstance(c, (int, float))]
                # A row whose fee cell is a standalone "I.C."/"c.s."/"S.C."/
                # "lab" marker (no fixed fee, by design) resolves to that
                # exact text rather than a number -- a code the source
                # explicitly marks as individually costed/client-specific is
                # meaningfully different from one where nothing could be
                # found at all, and the reference sheet should show that
                # literal marker instead of a generic "N/A". When no fee
                # column is identifiable (header=None, so every column got
                # scanned), this also has to block the numeric/text
                # fallbacks below entirely -- otherwise some *other*,
                # unrelated real numeric cell in the row (e.g. a page number
                # column) would get mistaken for the fee instead.
                no_fee_markers = [c.strip() for c in other_cells
                                   if isinstance(c, str) and _NO_FEE_MARKER_RE.match(c)]
                if no_fee_markers:
                    fee = no_fee_markers[-1]
                elif numeric_candidates:
                    fee = numeric_candidates[-1]
                else:
                    text_candidates = [v for cell in other_cells for v in _fee_candidates(cell)]
                    if text_candidates:
                        fee = text_candidates[-1]
                if fee is None:
                    continue
                if target_specialty is None:
                    fees[code] = fee
                else:
                    label = (_classify_specialty_cell(cells[specialty_col_idx])
                              if specialty_col_idx is not None and specialty_col_idx < len(cells) else None)
                    labeled_candidates.setdefault(code, []).append((label, fee))

    if target_specialty is not None:
        for code, candidates in labeled_candidates.items():
            exact = [f for label, f in candidates if label == target_specialty]
            neutral = [f for label, f in candidates if label is None]
            # First occurrence wins here, not last: unlike the generic
            # single-fee scanner above (where a later column is usually the
            # more current one), a code repeating multiple times *within*
            # one already specialty-scoped source is typically a summary/
            # total row followed by itemized sub-step breakdowns (seen in
            # QC's Endodontie sheet: code 33115's first row is the overall
            # retreatment fee, followed by four rows breaking it into
            # pulpectomy/cleaning/obturation/removal sub-fees) -- the
            # reference treats that first, overall row as the code's fee.
            if exact:
                fees[code] = exact[0]
            elif neutral:
                fees[code] = neutral[0]
            else:
                # No row for the requested specialty, and none unlabeled.
                # A row this same guide labels "GP" is the right last
                # resort: a combined GP+SP guide prices a procedure once
                # under GP and adds specialist rows only for the
                # specialties that bill it at a premium, so a specialist
                # performing a procedure with no premium row of its own
                # bills the guide's own general rate. Confirmed against
                # PE's combined guide: every flagged case where NO row
                # matched the requested specialty (e.g. 02101 asked as OS
                # or PE, 21221 as EN or PR, 41211 as OM or OP) has the
                # guide's GP fee as its reference value.
                #
                # Deliberately GP specifically, never "any other label":
                # picking whichever row happened to come last is what this
                # branch used to (correctly) refuse to do, and the other
                # labels really are wrong-context -- PE's guide also
                # carries "LTC" (long-term-care premium) rows for several
                # of these same codes (21221 at $240.50 vs GP's $185), and
                # a *different* specialty's premium row is wronger still
                # (41211 was resolving to the "PER" row's $140.40 when
                # asked for OM/OP, against a reference value of $84).
                #
                # Ranked below `neutral` as well, so a guide with genuinely
                # unlabeled rows keeps using those first, exactly as before.
                gp_rows = [f for label, f in candidates if label == "GP"]
                if gp_rows:
                    fees[code] = gp_rows[0]
                # Else: nothing usable in this source -- still left
                # unmatched rather than guessed at, so
                # load_pt_fees_by_subspecialty's own GP-fallback (which
                # reads the province's separate GP guide) can still run.
    return fees


def tables_from_spreadsheet_by_sheet(path: Path):
    """Like tables_from_spreadsheet, but yields (sheet_title, header,
    data_rows) triples -- used when a source splits sub-specialties across
    separate *worksheets* within one workbook, rather than separate files
    (see discover_pt_files) or a per-row specialty column (see
    find_row_specialty_column) -- e.g. QC's combined SP guide, with sheets
    named "Endodontie", "Parodontie", etc."""
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        wb = openpyxl.load_workbook(path, data_only=True)
        for ws in wb.worksheets:
            rows_iter = ws.iter_rows(values_only=True)
            first = next(rows_iter, None)
            if first is None:
                continue
            if _looks_like_header(first):
                yield ws.title, first, list(rows_iter)
            else:
                yield ws.title, None, [first, *rows_iter]
    elif suffix == ".xls":
        wb = xlrd.open_workbook(str(path))
        for sheet in wb.sheets():
            if sheet.nrows == 0:
                continue
            all_rows = [sheet.row_values(i) for i in range(sheet.nrows)]
            if _looks_like_header(all_rows[0]):
                yield sheet.name, all_rows[0], all_rows[1:]
            else:
                yield sheet.name, None, all_rows
    else:
        raise ValueError(f"Unsupported spreadsheet suffix: {suffix}")


# Substrings (rather than classify_file_specialty's whole-word regex) for
# recognizing a CDCP sub-specialty from a worksheet *title* -- these guides
# use French terms not in SUBSPECIALTY_FILE_MARKERS, and one of QC's sheet
# titles ("M?d - Path - Rad Buccale") has a corrupted accented character
# that breaks clean word-boundary matching, so "DIATRIQUE" (surviving
# fragment of "P?diatrique") and the "PATH"+"RAD" combination handled below
# are chosen specifically to survive that corruption.
_SHEET_SPECIALTY_SUBSTRINGS: dict[str, list[str]] = {
    "EN": ["ENDODONTIE"],
    "OS": ["MAXILLO-FACIALE", "MAXILLOFACIAL"],
    # CDCP's "PA"/"PE" codes are the reverse of what their letters suggest:
    # "PA" is Pediatric Dentistry, "PE" is Periodontics (see
    # SUBSPECIALTY_FILE_MARKERS for how this was confirmed). "PARODONTIE" is
    # French for Periodontics -> "PE"; "DIATRIQUE" (surviving fragment of
    # "Pédiatrique", Pediatric) -> "PA".
    "PE": ["PARODONTIE"],
    "PA": ["DIATRIQUE"],
    "PR": ["PROSTHODONTIE"],
    "OR": ["ORTHODONTIE"],
}


def classify_sheet_specialty(title: str) -> set[str]:
    """Which CDCP sub-specialty code(s), if any, a worksheet title
    identifies -- see tables_from_spreadsheet_by_sheet and
    _SHEET_SPECIALTY_SUBSTRINGS. Empty set means the sheet isn't specific to
    one sub-specialty (e.g. QC's "Section Commune" general-services sheet,
    or a "Membres-*" member-directory sheet, which has no fee data at all)."""
    name = title.upper()
    matches = {code for code, substrings in _SHEET_SPECIALTY_SUBSTRINGS.items()
               if any(s in name for s in substrings)}
    # "M?d - Path - Rad Buccale" combines Oral Medicine, Oral Pathology, and
    # Radiology into one sheet -- all three sub-specialties draw their fee
    # from it.
    if "PATH" in name and "RAD" in name:
        matches |= {"OM", "OP", "RA"}
    return matches


def tables_from_spreadsheet(path: Path):
    """Yields (header, data_rows) per worksheet. The first row of each sheet
    is treated as the header if it looks like one (see _looks_like_header);
    otherwise every row in that sheet is treated as data with no header."""
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        wb = openpyxl.load_workbook(path, data_only=True)
        for ws in wb.worksheets:
            rows_iter = ws.iter_rows(values_only=True)
            first = next(rows_iter, None)
            if first is None:
                continue
            if _looks_like_header(first):
                yield first, list(rows_iter)
            else:
                yield None, [first, *rows_iter]
    elif suffix == ".xls":
        wb = xlrd.open_workbook(str(path))
        for sheet in wb.sheets():
            if sheet.nrows == 0:
                continue
            all_rows = [sheet.row_values(i) for i in range(sheet.nrows)]
            if _looks_like_header(all_rows[0]):
                yield all_rows[0], all_rows[1:]
            else:
                yield None, all_rows
    else:
        raise ValueError(f"Unsupported spreadsheet suffix: {suffix}")


def tables_from_csv(path: Path):
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        all_rows = list(csv.reader(f))
    if not all_rows:
        return
    if _looks_like_header(all_rows[0]):
        yield all_rows[0], all_rows[1:]
    else:
        yield None, all_rows


_MC_FALLBACK_TAG = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"


def _in_mc_fallback(element) -> bool:
    """True if `element` is nested inside a <mc:Fallback> branch.

    Word represents each floating text box (or other AlternateContent-eligible
    shape) TWICE in the raw XML: once as a modern DrawingML <mc:Choice> and
    once as a legacy VML <mc:Fallback> -- both branches hold an identical
    copy of the same content (tables, paragraphs). A raw XML search for every
    <w:tbl>/<w:p> that doesn't account for this silently double-counts every
    text box's content. Confirmed against AB's DD guide: document.xml has 178
    <mc:AlternateContent> blocks, each with one <w:tbl> in Choice and one
    equal copy in Fallback -- explains both the doubled 88-vs-44 table count
    and the "same Prof/Lab/Total triplet repeated 2-4 times" pattern seen
    when reading paragraph text. Skipping the Fallback branch keeps exactly
    one copy of each text box's content.
    """
    for ancestor in element.iterancestors():
        if ancestor.tag == _MC_FALLBACK_TAG:
            return True
    return False


def _dedupe_repeated_leading_cell(row: list) -> list:
    """Collapse a run of cells at the START of a row that are all an exact
    duplicate of the first cell. Seen throughout AB's DD guide's text-box
    tables: a horizontally-merged code/description/fee cell (e.g.
    "41611\\tPartial Maxillary...\\t915.00") surfaces from python-docx as
    2-3 separate cells all reporting that SAME merged text, followed by the
    row's remaining genuine cells (Lab, Total). Left alone this misaligns
    every downstream positional read; collapsing the repeat back to one
    cell restores the intended code/description/prof/lab/total shape.
    Deliberately scoped to a repeat of cell[0] specifically (not a blanket
    adjacent-duplicate collapse across the whole row), since Prof and Lab
    can legitimately be equal for a real procedure."""
    if len(row) < 2:
        return row
    first = row[0]
    i = 1
    while i < len(row) and row[i] == first:
        i += 1
    return [first] + row[i:]


def _flatten_tab_cells(row: list) -> list:
    """Split any cell containing literal tab characters into separate
    cells. Seen in AB's DD guide: a merged cell can hold "code\\t
    description\\tfee" (or just "code\\tdescription") as one tab-joined
    string rather than the description/fee living in their own table
    cells -- e.g. "32510\\tComplete Maxillary - Reline...\\t586.00" next to
    separate Lab/Total cells. Splitting it out restores a normal
    code/description/prof/lab/total cell sequence."""
    out = []
    for cell in row:
        if isinstance(cell, str) and "\t" in cell:
            out.extend(part.strip() for part in cell.split("\t"))
        else:
            out.append(cell)
    return out


def _expand_stacked_row(row: list) -> list[list]:
    """A row whose cells each hold 2+ values joined by a blank line
    ("\\n\\n") represents multiple codes' worth of data horizontally
    squeezed into one physical table row -- seen in AB's DD guide, e.g. a
    code cell "31511\\n\\n31520" beside fee cells "732.00\\n\\n732.00" /
    "328.00\\n\\n328.00" / "1060.00\\n\\n1060.00" (two codes sharing one
    fee). Only expands when a candidate cell splits into the SAME count as
    the first (code) cell, so a cell that merely has an incidentally
    blank-line-separated description doesn't force a bogus split -- such a
    cell is instead broadcast unchanged to every expanded row."""
    if not row or not isinstance(row[0], str) or "\n\n" not in row[0]:
        return [row]
    code_parts = [p.strip() for p in row[0].split("\n\n")]
    n = len(code_parts)
    expanded_cells = [code_parts]
    for cell in row[1:]:
        if isinstance(cell, str) and "\n\n" in cell:
            parts = [p.strip() for p in cell.split("\n\n")]
            if len(parts) == n:
                expanded_cells.append(parts)
                continue
        expanded_cells.append([cell] * n)
    return [list(vals) for vals in zip(*expanded_cells)]


def _normalize_docx_rows(raw_rows: list[list]) -> list[list]:
    rows = []
    for row in raw_rows:
        row = _dedupe_repeated_leading_cell(row)
        row = _flatten_tab_cells(row)
        rows.extend(_expand_stacked_row(row))
    return rows


def tables_from_docx(path: Path):
    doc = Document(str(path))
    # Document.tables only returns tables that are *direct* children of the
    # document body -- a table nested inside another table's cell, or (seen
    # in AB's DD guide) inside a floating text box (<w:txbxContent>, which
    # isn't part of python-docx's Cell/Table object model at all), is
    # invisible to it. Searching the raw XML tree for every <w:tbl> element
    # regardless of where it's nested finds all of them: confirmed against
    # AB's guide, whose document.xml has 88 <w:tbl> elements total, only 22
    # of which doc.tables ever surfaced -- the other 66, in text boxes,
    # held real fee rows (e.g. code 10010) that were silently never even
    # considered for extraction. Each raw element is wrapped back into a
    # normal python-docx Table (its rows/cells API needs a parent, but
    # doesn't otherwise care that the element came from a text box) via the
    # same constructor Document.tables itself uses internally.
    for tbl_element in doc.element.body.findall(".//" + qn("w:tbl")):
        if _in_mc_fallback(tbl_element):
            continue
        table = _DocxTable(tbl_element, doc)
        raw_rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        rows = _normalize_docx_rows(raw_rows)
        if not rows:
            continue
        if _looks_like_header(rows[0]):
            yield rows[0], rows[1:]
        else:
            yield None, rows


def docx_paragraph_text(path: Path) -> str:
    doc = Document(str(path))
    # Document.paragraphs, like Document.tables (see tables_from_docx), only
    # returns *direct* children of the document body -- paragraphs inside a
    # floating text box are invisible to it. Confirmed against AB's DD
    # guide: it has 3691 <w:p> elements total but doc.paragraphs only
    # surfaces 878 of them -- the missing ones, in text boxes, hold real
    # fee data (a code and its Prof/Lab/Total figures, each as its own
    # paragraph/"line") that would otherwise silently vanish.
    paragraphs = (
        _DocxParagraph(p, doc)
        for p in doc.element.body.findall(".//" + qn("w:p"))
        if not _in_mc_fallback(p)
    )
    return "\n".join(p.text for p in paragraphs)


def _parse_french_amount(text: str) -> float | None:
    """Parse a French-Canadian formatted amount like "1 026 $" or
    "12 345,67 $" -- a space (or non-breaking space) as the thousands
    separator, a comma as the decimal separator, "$" trailing rather than
    leading. Common throughout Quebec's PDF fee guides. Plain digit-based
    parsing (extract_max_dollar) would otherwise fragment "1 026" into two
    separate numbers ("1" and "026") and grab the wrong one."""
    t = re.sub(r"\s+", "", text).replace("$", "").strip()
    t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


_FRENCH_NUMBER_RE = re.compile(r"\d[\d\s ]*(?:[.,]\d{2})?(?=\s*\$)")


def _parse_french_range_amount(text: str) -> float | None:
    """Parse a French-Canadian amount or range where "$" trails each number
    but there's no thousands-grouping or decimals to distinguish it from a
    plain integer ("97 $", "97 $ - 135 $"). Returns the largest value found,
    consistent with how ranges are handled elsewhere (see extract_max_dollar)."""
    values = [_parse_french_amount(m) for m in _FRENCH_NUMBER_RE.findall(text)]
    values = [v for v in values if v is not None]
    return max(values) if values else None


# Fee-shaped tokens, tried in order of specificity, each paired with the
# parser that turns its matched text into a float:
#   - French-Canadian amounts (space-grouped thousands, comma decimals,
#     trailing "$") -- must come first, before the plain-digit tiers below
#     would otherwise fragment a grouped number like "1 026" in two.
#   - a dollar amount with cents (optionally a range or "+ lab/exp" suffix)
#   - a bare two-decimal amount (some guides drop the "$")
#   - a dollar amount with NO cents ("$109") -- some guides (e.g. Yukon's)
#     quote whole-dollar fees with no decimal point. Without this tier these
#     fall all the way through to the bare-number fallback below, where a
#     later, unrelated number in the description ("...up to the age of 3
#     years...") can outrank the real fee.
#   - a "client specific" / "c.s." marker (no fixed fee -- extract_max_dollar
#     finds no digits in it and returns None, which is the correct result)
#   - a French-grouped whole number with no "$" and no decimals at all
#     ("1 261") -- some Quebec guides print 4+-digit fees this way. Must
#     come before the bare-number fallback below, which has no notion of
#     digit-grouping and would otherwise split "1 261" on the space and
#     take just the last group ("261") as if it were the whole fee.
#   - finally a bare whole number with no "$" at all (some guides, e.g.
#     Quebec's, list fees with no dollar sign or decimals -- riskiest, so
#     tried last and only within a code's own text segment, see
#     extract_codes_from_text).
_FEE_TOKEN_TIERS = [
    (_FRENCH_GROUPED_SPACE_RE, _parse_french_amount),
    (re.compile(r"\d+[.,]\d{2}\s*\$"), _parse_french_amount),
    # Bare digits with trailing "$" and no thousands-grouping or decimals
    # ("97 $", "97 $ - 135 $") -- the two tiers above require grouping or a
    # decimal to distinguish a French-formatted amount from an unrelated
    # bare number, but not every French guide's fees are large/precise
    # enough to have either.
    (re.compile(r"\d+\s*\$(?:\s*(?:to|-)\s*\d+\s*\$)?"), _parse_french_range_amount),
    (re.compile(r"\$[\d,]+\.\d{2}(?:\s*(?:to|-)\s*\$?[\d,]+\.\d{2})?(?:\s*\+\s*[A-Za-z]+)*"), extract_max_dollar),
    (re.compile(r"\b[\d,]+\.\d{2}\b(?:\s*(?:to|-)\s*\$?[\d,]+\.\d{2})?(?:\s*\+\s*[A-Za-z]+)*"), extract_max_dollar),
    (re.compile(r"\$[\d,]+(?:\s*(?:to|-)\s*\$?[\d,]+)?(?:\s*\+\s*[A-Za-z]+)*"), extract_max_dollar),
    # These three no-fixed-fee markers resolve to the *exact matched text*
    # (see _marker_text), not a number -- a code whose source says "I.C." is
    # meaningfully different from one where no fee could be found at all,
    # and the reference sheet should show that literal marker rather than a
    # generic "N/A" that looks the same as an extraction failure. The
    # internal "\s{0,2}" (rather than an unbounded "\s*") caps how much
    # whitespace can separate the two letters: layout-mode PDF text pads
    # rows out to their on-page column positions, so an unbounded "\s*"
    # could span clear across that padding and glue together two entirely
    # separate "S.C." occurrences many characters apart (seen in NB's DD
    # guide) into one garbled, meaningless matched string -- capping it
    # keeps the match to one genuine, tightly-printed occurrence of the
    # marker instead. Harmless for plain-mode/spreadsheet text, where a real
    # marker is never printed with more than a single separating space
    # anyway.
    (re.compile(r"c\.?\s{0,2}s\.?\s{0,2}\(?client specific\)?\.?", re.IGNORECASE), _marker_text),
    (re.compile(r"c\.?\s{0,2}s\.?", re.IGNORECASE), _marker_text),
    # "I.C." ("Individually Costed") -- another no-fixed-fee marker, same
    # idea as "c.s." above. Without recognizing it, a code's segment like
    # "74112 1 - 2 cm I.C." falls through every tier here to the risky
    # bare-whole-number fallback below, which then misreads the "2" from
    # the size range ("1 - 2 cm") in the description as if it were a fee.
    (re.compile(r"\bI\.?\s{0,2}C\.?\b", re.IGNORECASE), _marker_text),
    # "S.C." ("Service Charge"/"Independent Charge", per NB's DD guide) --
    # the same "no fixed fee" idea as "c.s." above, but with the two letters
    # reversed, so the "c.s." tier never matches it. Seen 136 times in NB's
    # DD guide alone; without this tier, every one of those rows' segments
    # (no digits at all -- just the word "S.C." itself) fails to match
    # anything all the way down to the bare-number tier too, so the code
    # never resolves at all, rather than correctly resolving to "no fee".
    (re.compile(r"\bS\.?\s{0,2}C\.?\b", re.IGNORECASE), _marker_text),
    (_FRENCH_GROUPED_NO_DOLLAR_RE, _parse_french_amount),
]
_BARE_NUMBER_TIER = (re.compile(r"\b\d{1,4}\b"), extract_max_dollar)

_ANY_CODE_RE = re.compile(r"\b\d{5}\b")

# How far into a code's text segment to look for its fee. Keeps the
# whole-number fallback tier from wandering into an unrelated number many
# sentences later.
_SEGMENT_SEARCH_WINDOW = 700

# pypdf's "layout" extraction mode (see load_fees_from_pdf) pads every line
# out to match the page's on-screen column positions, so the same code-to-
# fee distance that's a few characters in plain-mode text can be over a
# thousand characters in layout-mode text purely from that padding -- one
# QC DH guide table measured 1364 characters between one code and the next.
# _SEGMENT_SEARCH_WINDOW's 700 silently truncated before ever reaching the
# fee on some of those wide rows, so a code's segment then bled into the
# next code's own row (whose fee was still within the truncated leftover of
# the previous, already-bounded segment), misattributing a fee to the wrong
# code entirely. This is only used for layout-mode text, not plain-mode's
# (prose sources still want the tighter 700 -- widening it there raises the
# risk of the risky bare-whole-number fallback tier wandering into an
# unrelated number many sentences later in an ordinary paragraph).
_LAYOUT_SEGMENT_SEARCH_WINDOW = 3000


_YEAR_RE = re.compile(r"^(19|20)\d{2}$")

# A page-number token ("PAGE 9") immediately preceding a bare number in a
# segment -- seen when a code is the last one on its page, so its segment
# runs on into the next page's header/footer noise (a recurring "Table of
# contents\nPAGE 9\n<GUIDE TITLE> 2026" block, in QC's DH guide). Without
# this, the risky bare-whole-number fallback tier (which takes the *last*
# match in the window) picks up the page number instead of the real fee
# that came right before it -- e.g. "...Minimum of 14 images 150\n...PAGE
# 9..." resolved to 9 instead of 150.
_PRECEDED_BY_PAGE_RE = re.compile(r"PAGE\s*$", re.IGNORECASE)


def _fee_token_in_segment(segment: str, window_size: int = _SEGMENT_SEARCH_WINDOW) -> tuple[float | str | None, bool]:
    """Find the fee within one code's text segment.

    For the specific tiers (dollar amounts, decimals, "c.s."/"I.C." markers)
    the fee is the *first* match -- these guides put the fee right after the
    code ("00111 $63.80..."), or it's the only such token in a descriptive
    paragraph. The bare-whole-number fallback tier is different: guides that
    use it often prefix the fee with a quantity label ("00211 1 image 46"),
    so there the fee is the *last* match instead.

    Returns (fee, found_something) -- `found_something` is True whenever any
    tier matched at all, even a marker tier that parses to no numeric fee
    (e.g. "I.C."/"c.s."), and False only when nothing resembling a fee token
    was found anywhere in the window. The caller (extract_codes_from_text)
    uses this to decide whether this occurrence of the code is authoritative
    -- a marker match means "yes, this is the code's real entry, and it
    genuinely has no fixed fee," which is different from a segment that's
    just descriptive prose mentioning the code in passing (e.g. an intro
    paragraph referencing it before its actual price-table entry appears
    later) and shouldn't block a later, real entry from being found.
    """
    window = segment[:window_size]
    for pattern, parser in _FEE_TOKEN_TIERS:
        last_valid = None
        for match in pattern.finditer(window):
            text = match.group(0)
            if _YEAR_RE.match(text.strip()):
                continue
            last_valid = text
        if last_valid is not None:
            # If a tier matches more than once in the window (e.g. a
            # "CODE PROF LAB TOTAL" row rendered as plain tab-separated text
            # "31112  1118.00  746.00  1864.00"), the last match is the
            # running total, which is what the reference treats as the
            # code's fee -- and it's harmless when there's only one match,
            # since first and last are then the same token.
            return parser(last_valid), True

    bare_pattern, bare_parser = _BARE_NUMBER_TIER
    last_valid = None
    for match in bare_pattern.finditer(window):
        text = match.group(0)
        if _YEAR_RE.match(text.strip()):
            continue
        if _PRECEDED_BY_PAGE_RE.search(window[:match.start()]):
            continue
        last_valid = text
    if last_valid is not None:
        return bare_parser(last_valid), True
    return None, False


# Words that typically precede a *cross-reference* to another code inside a
# description ("...as per 00100", "...see code 00616", "...use code 00616")
# rather than that code's own entry. Occurrences preceded by one of these are
# ignored entirely -- both as segment boundaries and as extraction points --
# so a mention of another code doesn't truncate the current entry's segment,
# and a mention of *this* code doesn't get mistaken for its own definition.
_CROSS_REF_CUE_WORDS = {
    "per", "code", "codes", "see", "use", "using", "refer",
    "of", "or", "and", "lieu", "than", "instead", "not",
    "et",  # French "and" -- QC guides list codes together in French prose
}
# Deliberately "[ \t]*", not "\\s*": the cue word only counts when it sits on
# the SAME line as the code. A real prose cross-reference reads inline ("as
# per 00100", "see code 00616"), whereas a code that starts its own line is
# beginning its own entry -- and a fee table's wrapped description lines
# routinely end on one of these cue words purely by accident (confirmed in
# SK's SP guide, where 71201's description wraps to "...Sectioning of Tooth
# for Removal of" immediately above 71211's real entry). Letting "\\s*" reach
# back across the newline excluded that genuine entry as a "cross-reference",
# which in turn removed it as a segment boundary -- so the PREVIOUS code's
# segment ran on through it and picked up its fee (code 71209 resolving to
# 71211's $417 instead of its own $343).
# Anchored with "\\Z", not "$": "$" also matches immediately BEFORE a final
# newline, which would let the cue word be found across exactly the line
# break this is meant to stop at.
_PRECEDING_WORD_RE = re.compile(r"([A-Za-z]+)[ \t]*\Z")
_FOLLOWING_WORD_RE = re.compile(r"^\s*([A-Za-z]+)")
_FOLLOWED_BY_PERIOD_RE = re.compile(r"^\s?\.")
# "/" plus a digit is the other code-range shorthand these guides use, seen
# in SK's SP guide listing a section's member codes as "(To include 73111,
# 73141/42, 73151/54, 73161, 73171/72, 73181/84)" -- a pure cross-reference
# list, but one where the "/84" tail also reads as a bare number, so the
# codes in it were resolving to fees like $84 from that mention instead of
# from their own real entries further down. Requires the digit so an
# entry legitimately followed by a slash-delimited description isn't caught.
_FOLLOWED_BY_RANGE_RE = re.compile(r"^\s*(?:-|/\d)")
_FOLLOWED_BY_LIST_PUNCTUATION_RE = re.compile(r"^\s*[,)]")

# Marks the start of a back-of-guide numeric index (code -> page number),
# seen in QC's GP guide as several pages headed "INDEX / CODE NUMÉRIQUE
# NUMERICAL CODE" listing every code, followed by a separately-printed
# block of page numbers. In plain-mode (draw-order) text this is harmless
# on its own -- a code's segment there is just the next code with nothing
# in between, so _fee_token_in_segment finds nothing and doesn't lock it
# in. But in *layout* mode, which reconstructs the index's two columns
# (code, page number) side by side, each code ends up immediately followed
# by its page number -- and the bare-number fallback tier then misreads
# that page number as the code's fee (e.g. code 11300 wrongly resolving to
# "454", a page number, instead of its real fee). Since an index is purely
# a cross-reference to *where* a code's real entry lives, never a source of
# fee data itself, every code occurrence from this marker to the end of the
# text is excluded from consideration entirely -- same treatment as a
# prose cross-reference (see _is_cross_reference), just spanning a whole
# section instead of a few words.
#
# Deliberately matched on just the "CODE NUMÉRIQUE ... NUMERICAL CODE"
# bilingual column-header phrase, not the "INDEX" page heading that
# precedes it in plain-mode text: layout mode's column reconstruction
# scatters "INDEX" and the page number to a different position relative to
# this phrase (confirmed against the actual guide), so anchoring on
# "INDEX" first would silently fail to match layout-mode text and let this
# exact bug back in through that mode. The bilingual header phrase itself
# is positioned consistently in both modes.
_BACK_MATTER_INDEX_RE = re.compile(r"CODE\s+NUM\S*\s+NUMERICAL\s+CODE", re.IGNORECASE)


def _is_cross_reference(text: str, pos: int) -> bool:
    preceding = _PRECEDING_WORD_RE.search(text[max(0, pos - 30):pos])
    if preceding and preceding.group(1).lower() in _CROSS_REF_CUE_WORDS:
        return True
    end = pos + 5
    # A code immediately followed by a cue word ("...01200 et 01250)") is
    # sitting in the *middle* of a list of codes mentioned together in
    # prose -- not preceded by a cue word itself (that catches the last
    # item, "01250"), and not followed by list punctuation either (that
    # catches earlier items like "01120,"), so it needs its own check.
    following_word = _FOLLOWING_WORD_RE.match(text[end:end + 10])
    if following_word and following_word.group(1).lower() in _CROSS_REF_CUE_WORDS:
        return True
    following = text[end:end + 2]
    # A code immediately followed by a sentence-ending period ("...listed in
    # 00100.") reads as a citation closing out someone else's sentence, not
    # the start of this code's own entry (which is normally followed by more
    # descriptive text, not punctuation).
    if _FOLLOWED_BY_PERIOD_RE.match(following):
        return True
    # A code immediately followed by "-<digit>" ("04401-02") is a code-range
    # reference, typically from an alphabetical index/appendix section
    # ("Dental Legal Letters, 93121-23") rather than the code's own fee
    # entry -- a real entry is followed by a space then descriptive text.
    if _FOLLOWED_BY_RANGE_RE.match(following):
        return True
    # A code immediately followed by "," or ")" is part of a comma-separated
    # or parenthetical list of codes mentioned together in prose (e.g. a
    # French explanatory note "01120, 01130, 01200 et 01250"), not its own
    # entry.
    return bool(_FOLLOWED_BY_LIST_PUNCTUATION_RE.match(following))


def extract_codes_from_text(
    text: str, known_codes: set[str], window_size: int = _SEGMENT_SEARCH_WINDOW
) -> dict[str, float | str]:
    """Segment `text` by occurrences of *any* 5-digit code (not just ones we
    care about), then look for a fee token within each known code's segment
    (the text up to the next code of any kind). Bounding on any code --
    rather than only known ones -- keeps an unrelated nearby entry's numbers
    (for a code outside our CDCP list) from bleeding into the segment.
    Cross-reference mentions of a code within another entry's description are
    excluded from consideration entirely (see _is_cross_reference).

    The *first* legitimate (non-cross-reference) occurrence of a code wins
    once it actually resolves something -- either a numeric fee, or an
    explicit "no fixed fee" marker like "I.C."/"c.s." (see
    _fee_token_in_segment) -- even when that something is "no fixed fee":
    some combined guides (e.g. a "GP SP Fee Guide" covering both general and
    specialist rates in one PDF) list the same code a second time much later
    under a different, specialist-rate section with a real number, and that
    later occurrence must not be mistaken for this code's definition just
    because the first one's fee is genuinely variable. But a segment with no
    recognizable fee token *at all* (e.g. an intro paragraph mentioning the
    code by name well before its actual price-table entry, with no price
    nearby) isn't treated as authoritative, so a later, real entry still
    gets a chance -- otherwise a code mentioned in passing before its own
    definition would never resolve.

    Text from the start of a back-of-guide numeric index onward (see
    _BACK_MATTER_INDEX_RE) is excluded entirely first, before any of the
    above -- an index's code->page-number pairing must never be mistaken
    for a code->fee pairing.
    """
    index_start = _BACK_MATTER_INDEX_RE.search(text)
    if index_start is not None:
        text = text[:index_start.start()]

    all_matches = [
        m for m in _ANY_CODE_RE.finditer(text) if not _is_cross_reference(text, m.start())
    ]

    fees: dict[str, float | str] = {}
    seen: set[str] = set()
    for i, m in enumerate(all_matches):
        code = m.group(0)
        if code not in known_codes or code in seen:
            continue
        next_start = all_matches[i + 1].start() if i + 1 < len(all_matches) else len(text)
        segment = text[m.end():next_start]
        fee, found_something = _fee_token_in_segment(segment, window_size)
        if found_something:
            seen.add(code)
        if fee is not None:
            fees[code] = fee
    return fees


def load_fees_from_abbreviated_pdf(path: Path, known_codes: set[str]) -> dict[str, float]:
    """Positional-pairing reader for "abbreviated"/condensed fee guides
    (e.g. QC's "GUIDE ABRÉGÉ") laid out as two side-by-side columns per
    page -- a dense list of (code, description) entries on the left, and
    just the corresponding fees on the right, with no other text mixed in.

    pypdf's plain-mode (draw-order) extraction serializes this exactly
    column-major, same problem as NB's DD guide: every code+description
    entry for the whole page comes first, then every fee for the whole
    page afterward, as two separate un-paired blocks. Unlike NB's guide,
    though, *layout* mode doesn't reliably fix it here either -- confirmed
    against the real file, layout mode still misattributes fees (a code's
    search window ends up running past several other codes' worth of
    content because some of them fail to re-parse as clean 5-digit runs in
    the reconstructed layout, the same kind of split-digit artifact seen
    elsewhere in QC's guides).

    Since both blocks are internally *in the same relative order* (each
    fee is the Nth fee for the Nth code), the reliable fix is positional:
    read off every code in first-occurrence order, then every fee-shaped
    line in the trailing block in order, and pair them up index-for-index
    -- not a text-proximity heuristic at all, which is what makes this
    immune to both the column-major and split-digit problems above.

    This only actually pairs anything on a page where the two lists come
    out the *same length* -- if they don't, something about that page
    doesn't match this guide's assumed structure (or this isn't this kind
    of guide at all), and guessing a pairing anyway risks silently
    mismatching every code after the first discrepancy, which is worse
    than resolving nothing. That length check is also what makes this
    reader self-gating: run against an ordinary (non-abbreviated) guide,
    essentially no page will happen to have equal-length code and
    fee-shaped-line lists, so it naturally contributes nothing there
    instead of needing to be turned on per-province.

    That self-gating length check is NOT enough on its own, though: a
    back-of-guide numeric index page (see _BACK_MATTER_INDEX_RE) is
    *also* a list of codes followed by a same-length list of numbers --
    just page numbers, not fees -- so it can coincidentally pass the same
    check and get its page numbers paired in as if they were real fees
    (confirmed against QC's actual guide). Every page from the first
    index marker onward is skipped entirely for exactly that reason.
    """
    reader = pypdf.PdfReader(str(path))
    fees: dict[str, float] = {}
    past_index_start = False
    for page in reader.pages:
        text = page.extract_text() or ""
        if not past_index_start and _BACK_MATTER_INDEX_RE.search(text):
            past_index_start = True
        if past_index_start:
            continue
        code_matches = list(_ANY_CODE_RE.finditer(text))
        if not code_matches:
            continue
        codes_in_order: list[str] = []
        seen_on_page: set[str] = set()
        for m in code_matches:
            code = m.group(0)
            if code not in seen_on_page:
                seen_on_page.add(code)
                codes_in_order.append(code)

        tail = text[code_matches[-1].end():]
        fee_lines = [
            line for line in (raw.strip() for raw in tail.split("\n"))
            # A genuine fee-block line always starts with a digit; this
            # also naturally excludes leftover wrapped description text
            # from the page's last code (never starts with a digit) and a
            # stray running-header/footer year label like "2026" (starts
            # with a digit but is exactly a bare year, same exclusion
            # _fee_token_in_segment already applies elsewhere).
            if line and re.match(r"^\d", line) and not _YEAR_RE.match(line)
        ]
        if len(fee_lines) != len(codes_in_order):
            continue

        for code, fee_text in zip(codes_in_order, fee_lines):
            if code not in known_codes or code in fees:
                continue
            candidates = _fee_candidates(fee_text)
            if candidates:
                fees[code] = candidates[-1]
    return fees


def load_fees_from_pdf(path: Path, known_codes: set[str]) -> dict[str, float | str]:
    reader = pypdf.PdfReader(str(path))
    plain_text = "\n".join(page.extract_text() or "" for page in reader.pages)
    plain_fees = extract_codes_from_text(plain_text, known_codes)
    if len(plain_fees) >= len(known_codes):
        return plain_fees
    # Some guides' fee tables are laid out in columns that pypdf's default
    # (draw-order) extraction serializes column-major -- every description,
    # then every code, then every fee, each in its own block far from the
    # others -- which breaks the code-adjacent-to-its-own-fee assumption
    # extract_codes_from_text relies on entirely (seen in NB's DD guide: only
    # 24/398 codes resolved this way, most to the wrong fee). "layout" mode
    # instead reconstructs each row using the text's actual on-page position,
    # so a row's code and fee end up next to each other again (331/398 for
    # the same file). Only tried as a fallback, and only kept if it actually
    # resolves more codes than the plain pass did -- for a source where plain
    # mode already works (the common case), layout's heavier padding and
    # different line-wrapping isn't worth risking a regression on.
    layout_text = "\n".join(
        page.extract_text(extraction_mode="layout") or "" for page in reader.pages
    )
    layout_fees = extract_codes_from_text(layout_text, known_codes, window_size=_LAYOUT_SEGMENT_SEARCH_WINDOW)
    return layout_fees if len(layout_fees) > len(plain_fees) else plain_fees


def load_fees_from_docx(path: Path, known_codes: set[str], target_specialty: str | None = None) -> dict[str, float | str]:
    fees = extract_codes_from_rows(tables_from_docx(path), known_codes, target_specialty)
    missing = known_codes - fees.keys()
    if missing:
        # Some docx fee guides are prose/paragraphs rather than tables.
        fees.update({
            code: fee
            for code, fee in extract_codes_from_text(docx_paragraph_text(path), known_codes).items()
            if code not in fees
        })
    return fees


def load_fees_from_spreadsheet(path: Path, known_codes: set[str], target_specialty: str | None = None) -> dict[str, float | str]:
    if target_specialty is not None:
        # If any worksheet's title identifies it as specific to the
        # requested sub-specialty (e.g. QC's combined SP guide splits
        # sub-specialties across sheets like "Endodontie", "Parodontie"),
        # prefer those sheets over the rest of the workbook -- otherwise
        # the same code appearing on multiple specialty sheets (a common
        # diagnostic/radiograph code, say) resolves to whichever sheet
        # happens to come first, regardless of which specialty was asked
        # for.
        titled_tables = list(tables_from_spreadsheet_by_sheet(path))
        if any(classify_sheet_specialty(title) for title, _, _ in titled_tables):
            specific = [(h, r) for title, h, r in titled_tables
                        if target_specialty in classify_sheet_specialty(title)]
            fees = extract_codes_from_rows(specific, known_codes, target_specialty)
            missing = known_codes - fees.keys()
            if missing:
                general = [(h, r) for title, h, r in titled_tables if not classify_sheet_specialty(title)]
                fees.update({c: f for c, f in extract_codes_from_rows(general, missing, target_specialty).items()
                             if c not in fees})
            return fees
    return extract_codes_from_rows(tables_from_spreadsheet(path), known_codes, target_specialty)


def load_fees_from_csv(path: Path, known_codes: set[str], target_specialty: str | None = None) -> dict[str, float | str]:
    return extract_codes_from_rows(tables_from_csv(path), known_codes, target_specialty)


PROVINCE_ALIASES: dict[str, list[str]] = {
    "PE": ["PE", "PEI"],
    "YT": ["YK", "YT", "YU", "YUKON"],
    "YK": ["YK", "YT", "YU", "YUKON"],
    "NT": ["NT", "NWT"],
}


def discover_pt_files(specialty_dir: Path, province: str) -> list[Path]:
    """Find every file relevant to one province under `specialty_dir`.

    Two ways a file can qualify:
    1. Its name starts with the province's abbreviation (or a known alias),
       wherever it lives (covers most cases, including files inside a
       same-named subfolder like SP/BC/BC PA Fee Guide 2026.xlsx).
    2. It lives inside a per-province subfolder that doesn't itself follow
       the naming convention (e.g. SP/MB/MDA 2026 Endo....xlsx -- Manitoba's
       sub-specialty fee guides are split into several files named after the
       vendor, not the province).
    """
    aliases = PROVINCE_ALIASES.get(province, [province])
    # A trailing \b wouldn't reliably match here: \b only draws a boundary
    # where a "word" character (regex \w, which includes "_") meets a
    # non-word one, so "^ON\b" never matches "ON_PR_Fee_Guide_2026.xlsx" --
    # this project's actual naming convention -- since "_" doesn't create
    # one (see classify_file_specialty for the same underlying mistake,
    # found via this exact symptom on ON's SP guides). A lookahead for "the
    # next character isn't a letter/digit" (or end of string) reliably
    # covers "_", "-", " ", "." alike.
    patterns = [re.compile(rf"^{re.escape(a)}(?=[^A-Za-z0-9]|$)", re.IGNORECASE) for a in aliases]

    province_subdirs = [
        specialty_dir / alias for alias in aliases if (specialty_dir / alias).is_dir()
    ]

    matches = []
    for path in specialty_dir.rglob("*"):
        if not path.is_file():
            continue
        if any(p.match(path.name) for p in patterns):
            matches.append(path)
        elif any(sub in path.parents for sub in province_subdirs):
            matches.append(path)
    return matches


def _is_english(path: Path) -> bool:
    # str.split() only splits on whitespace -- against this project's
    # actual underscore-separated filenames (e.g. "QC_DH_Fee_Guide_FR"),
    # "FR" is never its own whitespace-delimited token, so the check below
    # would silently never fire without normalizing "_"/"-" to spaces
    # first (see classify_file_specialty for the same underlying issue).
    name = path.stem.upper().replace("_", " ").replace("-", " ")
    return "FR" not in name.split() and "FRENCH" not in name


def load_pt_fees_from_files(
    files: list[Path], known_codes: set[str], verbose: bool = True, target_specialty: str | None = None
):
    """Resolve PT fees for `known_codes` from a specific set of files, in
    priority order (filling in only the codes still missing at each step, so
    multiple partial sources combine): spreadsheets (xlsx/xlsm/xls), csv,
    docx, then pdf (English files before French, if language is discernible).

    `target_specialty`, if given, is passed down to the row-based readers
    (spreadsheet/csv/docx) so a combined guide with its own per-row
    specialty column (see find_row_specialty_column) can prefer the row
    actually labeled for that sub-specialty over whichever row happens to
    come first -- not applied to the pdf reader, which uses a different,
    text-segment-based scanner (see extract_codes_from_text).

    Returns (fees dict, list of (source_description, codes_found) used).
    """
    fees: dict[str, float | str] = {}
    sources_used: list[tuple[str, int]] = []

    def _apply(label: str, new_fees: dict[str, float]):
        added = {c: f for c, f in new_fees.items() if c not in fees}
        fees.update(added)
        if added:
            sources_used.append((label, len(added)))

    def _is_regional_variant(f: Path) -> bool:
        # Some provinces publish a separate fee guide for remote/northern
        # regions (e.g. "MB GP 2026 Fee Guide - NORTHERN.xlsx" alongside the
        # regular "MB GP 2026 Fee Guide.xlsx") with a premium over standard
        # rates. The reference file treats the standard-rate guide as
        # canonical, so it should be tried before, not after, a regional
        # variant -- otherwise the variant's higher rate wins for any code
        # both files list.
        return "NORTHERN" in f.stem.upper()

    spreadsheets = sorted(
        (f for f in files if f.suffix.lower() in SPREADSHEET_SUFFIXES),
        key=lambda f: (f.suffix.lower() != ".xlsx", _is_regional_variant(f)),  # prefer .xlsx, then standard-rate files
    )
    for f in spreadsheets:
        try:
            _apply(f.name, load_fees_from_spreadsheet(f, known_codes, target_specialty))
        except Exception as e:
            if verbose:
                print(f"    WARNING: failed to read {f.name}: {e}")

    for f in (f for f in files if f.suffix.lower() in CSV_SUFFIXES):
        if known_codes - fees.keys():
            try:
                _apply(f.name, load_fees_from_csv(f, known_codes, target_specialty))
            except Exception as e:
                if verbose:
                    print(f"    WARNING: failed to read {f.name}: {e}")

    for f in (f for f in files if f.suffix.lower() in DOC_SUFFIXES):
        if known_codes - fees.keys():
            try:
                _apply(f.name, load_fees_from_docx(f, known_codes, target_specialty))
            except Exception as e:
                if verbose:
                    print(f"    WARNING: failed to read {f.name}: {e}")

    pdfs = sorted(
        (f for f in files if f.suffix.lower() in PDF_SUFFIXES),
        key=lambda f: not _is_english(f),  # English first
    )
    # Tried before the generic PDF scanner below: a condensed/"abbreviated"
    # companion guide (see load_fees_from_abbreviated_pdf) that lays a
    # province's *entire* fee schedule out compactly, in a shape the
    # generic scanner can't reliably read even with its layout-mode
    # fallback. Safe to try unconditionally first across every source's
    # PDFs -- it only ever pairs anything on a page whose code count and
    # fee-line count match exactly, so it contributes nothing (not wrong
    # guesses, just nothing) for a PDF that isn't actually this shape.
    for f in pdfs:
        if known_codes - fees.keys():
            try:
                _apply(f"{f.name} (abbreviated)", load_fees_from_abbreviated_pdf(f, known_codes))
            except Exception as e:
                if verbose:
                    print(f"    WARNING: failed to read {f.name} as an abbreviated guide: {e}")
    for f in pdfs:
        if known_codes - fees.keys():
            try:
                _apply(f.name, load_fees_from_pdf(f, known_codes))
            except Exception as e:
                if verbose:
                    print(f"    WARNING: failed to read {f.name}: {e}")

    # Last resort, only for whatever spreadsheet/csv/docx/pdf-text all
    # failed to resolve: some PDFs (e.g. NL's DD guide) have no extractable
    # text at all -- their text was flattened to vector curves on export --
    # so nothing above can ever find anything in them no matter how the
    # content stream is read. OCR-ing the rendered page image is the only
    # way to recover data from a file like that. Deliberately tried last and
    # only for the remaining gap (never re-tried on a pdf that already
    # resolved fine above) since it's slow and occasionally misreads a
    # digit, and only imported here so a machine without the OCR
    # dependencies installed (see ocr_pdf_fees.py) just skips this tier
    # instead of failing the whole extraction run.
    if pdfs and (known_codes - fees.keys()):
        try:
            from ocr_pdf_fees import load_fees_from_pdf_via_ocr
        except ImportError as e:
            if verbose:
                print(f"    WARNING: OCR fallback unavailable ({e}); skipping")
        else:
            for f in pdfs:
                remaining = known_codes - fees.keys()
                if not remaining:
                    break
                try:
                    _apply(f"{f.name} (OCR)", load_fees_from_pdf_via_ocr(f, remaining))
                except Exception as e:
                    if verbose:
                        print(f"    WARNING: OCR failed for {f.name}: {e}")

    return fees, sources_used


def load_pt_fees(specialty_dir: Path, province: str, known_codes: set[str], verbose: bool = True):
    """Resolve PT fees for a province/specialty from whatever files are available.
    Returns (fees dict, list of (source_description, codes_found) used, files found)."""
    files = discover_pt_files(specialty_dir, province)
    fees, sources_used = load_pt_fees_from_files(files, known_codes, verbose)
    return fees, sources_used, files


# DD (denturist) fee guides commonly break a procedure's fee into separate
# Professional / Lab / Total columns (e.g. SK: "Prof Fee"/"Lab Fee"/"Total
# Fee", QC (French): "Honoraires"/"Frais de lab."/"Total", BC/MB/NT/NU/ON:
# "PROF"/"LAB"/"TOTAL" headers repeated before every section of the price
# table rather than once at the top of the sheet). The generic single-fee
# scanner above can't represent three distinct values per code, so DD gets
# its own role-aware column classifier and row scanner.
_DD_PROF_HEADER_RE = re.compile(r"\bprof(essional)?\b|honoraires", re.IGNORECASE)
_DD_LAB_HEADER_RE = re.compile(r"\blab\b|laboratoire|frais\s*de\s*lab", re.IGNORECASE)
_DD_TOTAL_HEADER_RE = re.compile(r"\btotal\b", re.IGNORECASE)
_YEAR_TOKEN_RE = re.compile(r"\b(19\d{2}|20\d{2})\b")
# This project's fee guides are all for the 2026 rate year -- a header
# naming *that* year ("Total Fee 2026") is still the current column, but one
# naming an earlier year ("2025 Prof Fee") is a prior-year column to ignore.
_CURRENT_GUIDE_YEAR = 2026


def _dd_column_role(cell) -> str | None:
    """Classify one header cell as 'prof', 'lab', 'total', or None. A cell
    naming a *prior* year (e.g. "2025 Prof Fee") is excluded from all three
    -- these guides list last year's Prof/Lab/Total fee alongside the
    current one under the same style of label, and only the current-year (or
    undated) column is this year's actual fee."""
    if not isinstance(cell, str):
        return None
    text = cell.strip()
    if not text:
        return None
    if any(int(m.group(0)) < _CURRENT_GUIDE_YEAR for m in _YEAR_TOKEN_RE.finditer(text)):
        return None
    if _DD_TOTAL_HEADER_RE.search(text):
        return "total"
    if _DD_LAB_HEADER_RE.search(text):
        return "lab"
    if _DD_PROF_HEADER_RE.search(text):
        return "prof"
    return None


def _find_dd_role_column_candidates(row) -> dict[str, list[int]]:
    """Every column matching each of Prof/Lab/Total, by header text, in
    left-to-right order -- a role label can legitimately appear more than
    once in the same header row (e.g. one QC sheet lists an unlabeled
    prior-year "Honoraires" column alongside the current one under the same
    text; one NS sheet lists a rounded current-year "TOTAL" alongside an
    unrounded prior-year "TOTAL"). Which occurrence is the real current-year
    one isn't consistently the first or the last across sources, so all
    candidates are kept and extract_dd_codes_from_rows picks between them
    using the actual data (whichever Total most closely equals Prof + Lab)
    rather than guessing from position."""
    candidates: dict[str, list[int]] = {}
    for i, cell in enumerate(row):
        role = _dd_column_role(cell)
        if role:
            candidates.setdefault(role, []).append(i)
    return candidates


def _looks_like_dd_role_header(row) -> bool:
    """True if `row` is itself a Prof/Lab/Total-labeled header row, as
    opposed to a data row. Some DD guides (e.g. BC's, MB's, NT's, NU's,
    ON's) repeat this header before every section of the price table rather
    than listing it once at the top of the sheet -- requiring 2+ identified
    roles avoids a data row that merely mentions "lab" in a description
    cell being mistaken for one."""
    return len(_find_dd_role_column_candidates(row)) >= 2


# A Lab cell containing only "+L" (optionally without the "+") -- ON's DD
# guide's own no-fixed-lab-fee marker, distinct from the "prof+L" suffix
# pattern PE's guide uses (see _DD_VARIABLE_LAB_RE/_DD_VARIABLE_LAB_TRIPLE_RE):
# here the whole cell is nothing but the marker, with no number attached at
# all. Canonicalized to "L" (dropping the "+") to match the reference sheet's
# own convention for this marker.
_LAB_VARIABLE_CELL_RE = re.compile(r"^\s*\+?\s*L\s*$", re.IGNORECASE)


def extract_dd_codes_from_rows(tables, known_codes: set[str]) -> dict[str, dict[str, float]]:
    """DD-specific (code -> {'prof': .., 'lab': .., 'total': ..}) scanner.
    Like extract_codes_from_rows, but for sources with identifiable
    Prof/Lab/Total columns (see _dd_column_role) instead of one generic fee
    column. The active column mapping is re-detected from any row that
    looks like a role header, not just the table's own detected header,
    since several guides repeat it before every section rather than listing
    it once. It also carries forward across tables that don't have (or
    repeat) a header of their own -- some docx guides (e.g. AB's) are a long
    series of small per-section tables sharing one consistent column layout,
    where only a handful of sections actually repeat the "DAC CODE /
    PROFESSIONAL FEE / LAB FEE / TOTAL FEE" header. Deliberately NOT
    backfilled the other direction (using a mapping to reinterpret tables
    that came *before* it was found): tried that and it broke NB's sheet,
    where a Prof/Lab/Total-labeled section buried in the middle of the file
    (digital denture services) got treated as the layout for unrelated
    earlier sections, producing wrong-but-plausible-looking $0 fees instead
    of correctly leaving them for the single-fee fallback. A code with no
    active role mapping yet (no header seen so far in this file) or no fee
    found in any active role column is simply skipped -- the caller falls
    back to the generic single-fee scanner for those.
    """
    results: dict[str, dict[str, float]] = {}
    active_candidates: dict[str, list[int]] = {}
    code_col_indices: list[int] = []
    for header, rows in tables:
        if header:
            header_candidates = _find_dd_role_column_candidates(header)
            if header_candidates:
                active_candidates = header_candidates
            header_code_cols = find_code_column_indices(header)
            if header_code_cols:
                code_col_indices = header_code_cols
        for row in rows:
            cells = list(row)
            if _looks_like_dd_role_header(cells):
                active_candidates = _find_dd_role_column_candidates(cells)
                detected_code_cols = find_code_column_indices(cells)
                if detected_code_cols:
                    code_col_indices = detected_code_cols
                continue
            if not active_candidates:
                continue
            candidate_code_col_indices = code_col_indices or None
            code = None
            for i, cell in enumerate(cells):
                if candidate_code_col_indices is not None and i not in candidate_code_col_indices:
                    continue
                c = normalize_code(cell)
                if c and c in known_codes:
                    code = c
                    break
            if code is None or code in results:
                continue

            def _role_cells(role: str) -> list[tuple[object, float | None]]:
                out = []
                for idx in active_candidates.get(role, []):
                    if idx >= len(cells):
                        continue
                    cell = cells[idx]
                    out.append((cell, extract_max_dollar(cell)))
                return out

            def _is_blank(cell) -> bool:
                return cell is None or (isinstance(cell, str) and not cell.strip())

            prof_cells = _role_cells("prof")
            lab_cells = _role_cells("lab")
            total_candidates = [f for _, f in _role_cells("total") if f is not None]
            prof_candidates = [f for _, f in prof_cells if f is not None]
            lab_candidates = [f for _, f in lab_cells if f is not None]
            # A Lab cell holding just "+L" (no attached number -- ON's DD
            # guide's own way of flagging "this procedure's lab component is
            # billed separately/variably") is real information, not a blank:
            # kept as literal marker text "L", the same convention as
            # I.C./c.s./S.C./B.R. elsewhere, rather than left unresolved.
            lab_marker = None
            if not lab_candidates:
                for cell, _ in lab_cells:
                    if isinstance(cell, str) and _LAB_VARIABLE_CELL_RE.match(cell):
                        lab_marker = "L"
                        break
            # A genuinely blank Prof or Lab cell (as opposed to a
            # non-numeric marker like "+L"/"SC" for a variable/
            # client-specific charge) means this procedure simply has no
            # component on that side -- 0, not unknown -- matching the
            # reference's convention (and load_cdcp_dd_fees's `or 0` for the
            # CDCP side). Only default one side to 0 this way when the
            # *other* side has a real value, so a row where both Prof and
            # Lab columns are blank (e.g. a flat, undivided fee) isn't
            # forced to a fake $0 + $0 instead of being left for Total alone.
            if not prof_candidates and lab_candidates and prof_cells and _is_blank(prof_cells[-1][0]):
                prof_candidates = [0.0]
            if not lab_candidates and prof_candidates and lab_cells and _is_blank(lab_cells[-1][0]):
                lab_candidates = [0.0]
            # Any of the three roles can end up with more than one matching
            # column: a source can repeat the same "Prof Fee"/"Lab Fee"
            # header for its unlabeled prior-year figure alongside the
            # current one (SK's denture sections), or repeat "Total" for a
            # rounded current-year figure alongside an unrounded prior-year
            # one (NS) -- in either left-right order, so position alone
            # can't tell current from prior. Instead, pick whichever
            # combination of candidates actually satisfies this row's own
            # Prof + Lab = Total, rather than guessing by position.
            prof = prof_candidates[-1] if prof_candidates else None
            lab = lab_candidates[-1] if lab_candidates else (lab_marker if lab_marker else None)
            total = total_candidates[-1] if total_candidates else None
            if len(prof_candidates) > 1 or len(lab_candidates) > 1 or len(total_candidates) > 1:
                best = None
                best_diff = None
                # Iterate right-to-left so that on an exact tie (seen in one
                # QC sheet, where a uniform escalation factor means BOTH the
                # prior-year and current-year triples satisfy Prof+Lab=Total
                # exactly), the rightmost/current-year combination is found
                # first and a later, equally-good match doesn't overwrite it
                # -- consistent with this file's "rightmost = current"
                # convention (see extract_codes_from_rows, _find_dd_role_
                # column_candidates).
                for p in reversed(prof_candidates or [None]):
                    for l in reversed(lab_candidates or [None]):
                        for t in reversed(total_candidates or [None]):
                            if p is None or l is None or t is None:
                                continue
                            diff = abs((p + l) - t)
                            if best is None or diff < best_diff:
                                best, best_diff = (p, l, t), diff
                if best is not None:
                    prof, lab, total = best

            values: dict[str, float] = {}
            if prof is not None:
                values["prof"] = prof
            if lab is not None:
                values["lab"] = lab
            if total is not None:
                values["total"] = total
            if values:
                results[code] = values
    return results


_DD_PDF_LINE_CODE_RE = re.compile(r"\b(\d{5})\b")
_DD_PDF_NUMBER_RE = re.compile(r"\d[\d,]*\.\d{2}")

# A Prof fee immediately followed by a Lab fee marked "+L" -- a variable,
# unspecified additional lab charge, not a fixed number (seen in PE's DD
# guide, e.g. "1604.00   789+L   2393.00+L": Prof $1604, Lab "$789 plus an
# unspecified lab surcharge", Total likewise "+L"). "789" and "2393.00" here
# have no decimal point-free digit run before them respectively that would
# make them safe to read as a second/third fixed number the way the
# Prof/Lab/Total tier above does -- Lab and Total are genuinely variable
# here, not just unresolved, so only Prof (the one unambiguous, fixed
# number on the line) is claimed.
_DD_VARIABLE_LAB_RE = re.compile(r"(\d[\d,]*\.\d{2})\s*[\t ]*\d[\d,]*\s*\+\s*L\b")

# The full Prof/Lab/Total triple when Lab and Total both carry the "+L"
# marker -- confirmed against PE's DD guide across 18 such lines (e.g.
# "1531.00   755+L   2286.00+L") that "+L" is a footnote annotation ("this
# procedure may incur an additional lab charge"), not a sign the printed
# figures themselves are unknown: every one of the 18 satisfies
# Prof + Lab == Total exactly once "+L" is stripped (1531 + 755 = 2286).
# Only claimed when that arithmetic holds; otherwise falls through to
# _DD_VARIABLE_LAB_RE's Prof-only claim below, same as before.
_DD_VARIABLE_LAB_TRIPLE_RE = re.compile(
    r"(\d[\d,]*\.\d{2})\s*[\t ]*(\d[\d,]*(?:\.\d{2})?)\s*\+\s*L\b\s*(\d[\d,]*\.\d{2})\s*\+\s*L\b"
)

# Same no-fixed-fee markers as _NO_FEE_MARKER_RE, but usable with findall
# against a whole text line instead of only an exact, whole-cell match --
# _NO_FEE_MARKER_RE is anchored (^...$) for that reason and can't find a
# marker embedded partway through a longer line of description text. "lab
# fee" is deliberately left out of this alternation (unlike
# _NO_FEE_MARKER_RE's): free-flowing description text can genuinely contain
# the word "Lab" (e.g. "Lab Processed") without meaning the marker.
_DD_LINE_MARKER_RE = re.compile(r"\b(?:I\.?\s*C\.?|c\.?\s*s\.?|s\.?\s*c\.?|B\.?\s*R\.?)\.?\b", re.IGNORECASE)


def extract_dd_codes_from_lines(text: str, known_codes: set[str]) -> dict[str, dict[str, float]]:
    """DD-specific line scanner for guides that print a code's Prof/Lab/Total
    fees together on the same line, in that left-to-right order -- e.g.
    "Diagnostic Model - Maxillary   10120   122.00   184.00   306.00"
    (confirmed left-to-right by Total == Prof + Lab in every such line).
    Format-agnostic despite the name's origin (it was built for pdf "layout"
    extraction-mode text, see load_fees_from_pdf -- plain-mode text
    serializes multi-column tables in draw order, scattering a row's numbers
    away from its code entirely) -- it works equally well on docx paragraph
    text (see docx_paragraph_text), confirmed against PE's DD guide, whose
    fee table isn't a real Word table at all, just tab-separated paragraphs
    ("31112\\t\\t1118.00\\t\\t746.00\\t\\t1864.00").

    The generic single-fee scanner (load_fees_from_pdf / load_fees_from_docx)
    can only ever take one number per code (the rightmost, i.e. Total),
    silently losing the Prof/Lab breakdown for any such row. This recovers
    it directly from the numbers on the same line as the code, without
    needing a recognizable column header -- NB's own header ("CODE /
    CLINICAL FEE / TOTAL FEE") doesn't even name a "Lab" column at all; the
    breakdown only shows up as a third number on the rows that have one.

    Only the two HIGH-CONFIDENCE shapes are claimed here:
    - Exactly two IDENTICAL numbers, common for procedures with no lab
      component at all (Clinical Fee == Total Fee); recorded as Total with
      Lab forced to 0.0, not left unknown, matching this project's
      established "a genuinely blank side is 0, not unknown" DD convention
      (see extract_dd_codes_from_rows).
    - Three or more numbers read as (prof, lab, total), kept only if it
      actually satisfies prof + lab == total (within a cent).

    Every other shape (exactly one number, two unequal numbers, or 3+
    numbers that don't satisfy the arithmetic check) is left unclaimed
    rather than guessed at, so those codes fall through to the generic
    single-fee fallback (load_pt_fees_from_files) in
    load_pt_dd_fees_from_files instead. Guessing in those low-confidence
    cases used to work fine for NB (the guide this scanner was built and
    tested against), but wrongly intercepted codes on other provinces' DD
    PDFs that the older, more robust generic fallback already resolved
    correctly.
    """
    results: dict[str, dict[str, float]] = {}
    for line in text.split("\n"):
        code_match = _DD_PDF_LINE_CODE_RE.search(line)
        if not code_match:
            continue
        code = code_match.group(1)
        if code not in known_codes or code in results:
            continue
        segment = line[code_match.end():]
        numbers = [float(n.replace(",", "")) for n in _DD_PDF_NUMBER_RE.findall(segment)]
        if not numbers:
            # No numbers at all -- but a no-fixed-fee marker ("S.C.", "I.C.",
            # "B.R.", ...) can appear on the line instead, printed twice the
            # same way a shared fee is (see the 2-identical-numbers case
            # below): NB's guide prints "S.C." once for Clinical/Total each,
            # e.g. "73008 ... S.C. S.C.". Only claimed when both occurrences
            # match (after normalizing spacing/case) -- one lone marker
            # occurrence is left for the generic fallback, same as before.
            marker_matches = _DD_LINE_MARKER_RE.findall(segment)
            if len(marker_matches) == 2:
                norm = [re.sub(r"\s+", "", m).upper() for m in marker_matches]
                if norm[0] == norm[1]:
                    marker_text = marker_matches[0].strip()
                    results[code] = {"prof": marker_text, "lab": marker_text, "total": marker_text}
            continue
        if len(numbers) == 2 and numbers[0] == numbers[1]:
            # Two identical numbers with no third (Total) token on the line
            # -- NB's guide's own "CLINICAL FEE / LABORATORY / TOTAL FEE"
            # header names 3 roles, but for a procedure with no separate lab
            # step it only ever prints the shared figure once more (Clinical
            # and Total), not three times. Confirmed against NB's own
            # ground-truth reference: Internal Lab Fee for these codes
            # mirrors Prof/Total (e.g. code 70150 = 73/73/73), not 0 -- see
            # resolve_dd_role_values for why mirroring Lab into Prof here
            # doesn't double-count Total.
            results[code] = {"prof": numbers[0], "lab": numbers[0], "total": numbers[0]}
        elif len(numbers) >= 3:
            prof, lab, total = numbers[0], numbers[1], numbers[-1]
            # A small (<= $1) mismatch is a source-side rounding slip, not a
            # sign this triple is misread -- confirmed against PE's DD
            # guide: of 6 lines whose 3 numbers don't satisfy prof+lab=total
            # exactly, 5 are off by precisely $1.00 (e.g. 953.00 + 468.00 =
            # 1421.00 but the guide's own Total column reads 1420.00) while
            # the 1 genuinely-wrong line is off by $131 -- comfortably far
            # outside this tolerance, so it's still correctly left unclaimed.
            if abs((prof + lab) - total) <= 1.00:
                results[code] = {"prof": prof, "lab": lab, "total": total}
            # else: numbers on the line don't satisfy prof+lab=total even
            # loosely, so this isn't confidently a Prof/Lab/Total triple --
            # leave the code unclaimed rather than guessing, so it falls
            # through to the generic single-fee fallback
            # (load_pt_fees_from_files) in load_pt_dd_fees_from_files
            # instead.
        else:
            # Neither of the two shapes above matched (e.g. two *unequal*
            # numbers -- a decimal-matching Prof and a decimal-matching
            # Total, with an in-between Lab value marked "+L"). Try the full
            # Prof/Lab/Total triple first (see _DD_VARIABLE_LAB_TRIPLE_RE):
            # "+L" turns out to be a footnote annotation on real, fixed
            # figures, not a sign they're unknown, but only trusted once
            # Prof + Lab == Total confirms the numbers were read correctly.
            # Falls back to claiming just Prof (the one unambiguous fixed
            # number right after the code) when the triple doesn't parse or
            # doesn't satisfy that check.
            var_triple_match = _DD_VARIABLE_LAB_TRIPLE_RE.match(segment.lstrip())
            if var_triple_match:
                prof = float(var_triple_match.group(1).replace(",", ""))
                lab = float(var_triple_match.group(2).replace(",", ""))
                total = float(var_triple_match.group(3).replace(",", ""))
                if abs((prof + lab) - total) < 0.01:
                    results[code] = {"prof": prof, "lab": lab, "total": total}
                else:
                    results[code] = {"prof": prof}
            else:
                var_lab_match = _DD_VARIABLE_LAB_RE.match(segment.lstrip())
                if var_lab_match:
                    results[code] = {"prof": float(var_lab_match.group(1))}
            # else: not confident enough to claim here -- same fallthrough
            # to the generic fallback. These low-confidence branches used
            # to record a bare {"total": ...} guess directly, which was
            # fine for NB (the guide this scanner was built and tested
            # against) but wrongly intercepted codes on other provinces' DD
            # PDFs (e.g. PE) that the older, more robust generic fallback
            # already resolved correctly -- causing a regression when this
            # tier was made to run on every DD PDF, not just NB's.
    return results


def extract_dd_codes_from_headerless_docx_tables(tables, known_codes: set[str]) -> dict[str, dict[str, float]]:
    """DD-specific scanner for docx tables with *no* header row at all (seen
    throughout AB's DD guide: it's a long series of small per-section
    tables, most un-headered, each consistently shaped
    code / description / Prof / Lab / Total -- e.g. ['31310', 'Complete
    Maxillary - Standard', '1032.00', '673.00', '1705.00']). Without a
    header, extract_dd_codes_from_rows' role-column detection has nothing to
    go on and skips these tables entirely.

    Since the shape is consistent within a table -- code, then description,
    then purely positional Prof/Lab/Total-shaped values -- this reads the
    cells after the description the same way extract_dd_codes_from_lines
    reads numbers off a text line: two equal values -> Total with Lab
    forced to 0; three values satisfying Prof + Lab == Total -> the full
    triple; anything else falls back to the single highest-confidence value
    (the first real number found) recorded as Total alone, letting
    resolve_dd_role_values's "one known value stands for both Prof and
    Total" convention apply -- a genuine table row (not free-flowing prose)
    is inherently higher-confidence than a text-line match, so this is
    deliberately less conservative than extract_dd_codes_from_lines about
    claiming the single- or mismatched-value cases rather than leaving them
    for a later, worse fallback.

    A row whose candidate cells are no-fixed-fee markers ("B.R." for AB,
    same idea as "I.C."/"c.s." elsewhere -- see _NO_FEE_MARKER_RE) with no
    real number at all resolves to that marker text instead of a number,
    same convention as everywhere else in this project (see _marker_text).
    """
    results: dict[str, dict[str, float | str]] = {}
    for header, rows in tables:
        # Skip only a *genuine* Prof/Lab/Total role header -- tables_from_docx's
        # _looks_like_header is a generic heuristic (several short,
        # densely-packed cells) that can misfire on a stray, mostly-blank
        # continuation row from the previous table's layout (seen in AB's
        # guide: a header of ['', 'of one clasps', '', '', '']) and yield it
        # as this table's "header" -- skipping every such table here would
        # wrongly hand it to extract_dd_codes_from_rows instead, which would
        # then use whatever role mapping happened to carry forward from an
        # earlier, unrelated table rather than reading this row positionally.
        if header and _looks_like_dd_role_header(header):
            continue
        for row in rows:
            cells = list(row)
            if len(cells) < 3:
                continue
            code = normalize_code(cells[0])
            if code is None or code not in known_codes or code in results:
                continue
            candidates = cells[2:]
            values: list[float | None] = []
            marker_text = None
            for cell in candidates:
                text = str(cell).strip() if cell is not None else ""
                if _NO_FEE_MARKER_RE.match(text):
                    marker_text = marker_text or text
                    values.append(None)
                    continue
                values.append(extract_max_dollar(cell))
            real = [v for v in values if v is not None]
            if not real:
                if marker_text is not None:
                    results[code] = {"prof": marker_text}
                continue
            if len(real) == 2 and real[0] == real[1]:
                # Two identical figures with no distinguishable 3rd (Total)
                # cell -- same shape and same fix as
                # extract_dd_codes_from_lines' 2-identical-numbers case (see
                # there, and resolve_dd_role_values for why mirroring into
                # Lab instead of forcing 0 doesn't double-count Total):
                # confirmed against AB's own "CLINICAL FEE / LABORATORY /
                # TOTAL FEE" table rows for codes with no separate lab step
                # (e.g. 71010 = 85/85/85, not 85/0/85).
                results[code] = {"prof": real[0], "lab": real[0], "total": real[0]}
            elif len(values) >= 3 and all(v is not None for v in values[:3]):
                prof, lab, total = values[0], values[1], values[2]
                if abs((prof + lab) - total) < 0.01:
                    results[code] = {"prof": prof, "lab": lab, "total": total}
                else:
                    # The row's own printed 3rd (Total) number doesn't
                    # satisfy Prof + Lab -- confirmed against AB's DD guide
                    # to be a source-side typo in that 3rd number alone
                    # (e.g. code 41711: printed Total 613.00, but Prof
                    # 420.00 + Lab 211.00 = 631.00, matching the reference
                    # exactly), not a sign Prof/Lab themselves are
                    # misaligned. Positions 0/1 are still claimed as
                    # Prof/Lab; the printed Total is dropped rather than
                    # trusted, letting resolve_dd_role_values recompute it
                    # as Prof + Lab instead.
                    results[code] = {"prof": prof, "lab": lab}
            else:
                results[code] = {"total": real[0]}
    return results


def load_pt_dd_fees_from_files(files: list[Path], known_codes: set[str], verbose: bool = True):
    """DD-specific: resolves separate Prof/Lab/Total fees per code (see
    extract_dd_codes_from_rows) from spreadsheet sources whose columns are
    labeled that way, before falling back to the generic single-fee scan
    (load_pt_fees_from_files) for whatever's still missing -- covering
    sources with no distinguishable Prof/Lab/Total columns (plain text/PDF
    guides, docx) as well as spreadsheets without them. A fallback match is
    recorded as {'total': fee}: the generic scanner takes the rightmost
    numeric candidate in a row, which in every DD source seen so far is the
    Total column when one exists, so that's the safest single role to
    attribute it to.
    Returns (dict[code, {'prof': .., 'lab': .., 'total': ..}], sources_used).
    """
    role_fees: dict[str, dict[str, float]] = {}
    sources_used: list[tuple[str, int]] = []

    def _is_regional_variant(f: Path) -> bool:
        return "NORTHERN" in f.stem.upper()

    spreadsheets = sorted(
        (f for f in files if f.suffix.lower() in SPREADSHEET_SUFFIXES),
        key=lambda f: (f.suffix.lower() != ".xlsx", _is_regional_variant(f)),
    )
    for f in spreadsheets:
        remaining = known_codes - role_fees.keys()
        if not remaining:
            break
        try:
            new_fees = extract_dd_codes_from_rows(tables_from_spreadsheet(f), remaining)
        except Exception as e:
            if verbose:
                print(f"    WARNING: failed to read {f.name}: {e}")
            continue
        if new_fees:
            role_fees.update(new_fees)
            sources_used.append((f.name, len(new_fees)))

    # AB's docx (and likely others of the same shape) is a long series of
    # small per-section tables, most with no header row of their own. Tried
    # first, ahead of the header-based pass below: extract_dd_codes_from_rows
    # deliberately *carries forward* the last real role header it saw across
    # tables that don't repeat one of their own (some docx guides genuinely
    # are one consistent layout split into many small un-headered tables --
    # see that function's docstring), which is exactly wrong once a docx has
    # dozens of un-headered tables that DON'T all share one layout (seen
    # once tables_from_docx started finding tables nested in text boxes,
    # AB's guide went from 22 top-level tables to 88 total) -- a stale
    # carried-over column mapping from an unrelated earlier table then
    # misreads a later table's own Code column as if it were Prof. This
    # tier never looks past the row it's on, so it can't be confused that
    # way; only tables it can't confidently resolve on their own are left
    # for the header-based/carry-forward pass afterward.
    for f in (f for f in files if f.suffix.lower() in DOC_SUFFIXES):
        remaining = known_codes - role_fees.keys()
        if not remaining:
            break
        try:
            new_fees = extract_dd_codes_from_headerless_docx_tables(tables_from_docx(f), remaining)
        except Exception as e:
            if verbose:
                print(f"    WARNING: failed to read {f.name} as headerless tables: {e}")
            continue
        if new_fees:
            role_fees.update(new_fees)
            sources_used.append((f"{f.name} (headerless tables)", len(new_fees)))

    # Whatever's left, including sections that genuinely do repeat a
    # "DAC CODE / PROFESSIONAL FEE / LAB FEE / TOTAL FEE" header (or share a
    # consistent layout with one, per the carry-forward behavior described
    # above), resolved the same way as a spreadsheet's labeled columns.
    for f in (f for f in files if f.suffix.lower() in DOC_SUFFIXES):
        remaining = known_codes - role_fees.keys()
        if not remaining:
            break
        try:
            new_fees = extract_dd_codes_from_rows(tables_from_docx(f), remaining)
        except Exception as e:
            if verbose:
                print(f"    WARNING: failed to read {f.name}: {e}")
            continue
        if new_fees:
            role_fees.update(new_fees)
            sources_used.append((f.name, len(new_fees)))

    # Some docx guides (e.g. PE's) have no real Word table for their fee
    # data at all -- it's plain paragraphs with tab-separated values, which
    # the table-based pass above (tables_from_docx) can't see. Tried on
    # paragraph text the same way as PDF layout text (see
    # extract_dd_codes_from_lines) so those guides don't lose their
    # Prof/Lab breakdown down to Total-only (or worse, have their generic
    # single-fee fallback's Total value wrongly stand in as Prof too --
    # see load_pt_fees_from_files below) just because there's no table.
    for f in (f for f in files if f.suffix.lower() in DOC_SUFFIXES):
        remaining = known_codes - role_fees.keys()
        if not remaining:
            break
        try:
            new_fees = extract_dd_codes_from_lines(docx_paragraph_text(f), remaining)
        except Exception as e:
            if verbose:
                print(f"    WARNING: failed to read {f.name} as paragraph text: {e}")
            continue
        if new_fees:
            role_fees.update(new_fees)
            sources_used.append((f"{f.name} (paragraph text)", len(new_fees)))

    # PDFs whose rows carry Prof/Lab/Total on the same line as the code
    # (see extract_dd_codes_from_lines) -- tried before the generic
    # single-fee fallback below so a PDF source doesn't lose its Prof/Lab
    # breakdown down to Total-only just because it isn't a spreadsheet/docx.
    for f in (f for f in files if f.suffix.lower() in PDF_SUFFIXES):
        remaining = known_codes - role_fees.keys()
        if not remaining:
            break
        try:
            reader = pypdf.PdfReader(str(f))
            layout_text = "\n".join(
                page.extract_text(extraction_mode="layout") or "" for page in reader.pages
            )
            new_fees = extract_dd_codes_from_lines(layout_text, remaining)
        except Exception as e:
            if verbose:
                print(f"    WARNING: failed to read {f.name}: {e}")
            continue
        if new_fees:
            role_fees.update(new_fees)
            sources_used.append((f"{f.name} (pdf rows)", len(new_fees)))

    # The final generic single-fee scanner (load_pt_fees_from_files) is
    # deliberately NOT pointed at docx sources here: every DD-aware docx
    # tier above (headerless tables, header-based rows, paragraph-line
    # pairing) already had its shot at the SAME document, each requiring a
    # code and its fee to be on the same row/line -- a much higher
    # confidence bar than the generic scanner's. A code neither could
    # resolve means the docx simply doesn't say, in any readable position,
    # what that code's fee is (confirmed against AB's DD guide: several
    # codes' fee numbers live in a *different* floating text-box shape than
    # their code, with no reliable reading-order link between the two).
    # Letting the generic scanner take one more, looser pass at that same
    # document was pulling in a nearby-but-wrong number often enough to be
    # worse than just leaving the code unresolved (N/A) -- confirmed
    # against several codes users flagged as *wrong*, not merely missing.
    # PDF/spreadsheet sources aren't affected: those never got a DD-aware
    # docx tier's more careful attempt in the first place.
    non_docx_files = [f for f in files if f.suffix.lower() not in DOC_SUFFIXES]
    missing = known_codes - role_fees.keys()
    if missing and non_docx_files:
        single_fees, single_sources = load_pt_fees_from_files(non_docx_files, missing, verbose)
        for code, fee in single_fees.items():
            role_fees[code] = {"total": fee}
        sources_used.extend(single_sources)

    return role_fees, sources_used


def load_pt_dd_fees(specialty_dir: Path, province: str, known_codes: set[str], verbose: bool = True):
    """Resolve DD PT fees for a province from whatever files are available.
    Returns (dict[code, {'prof': .., 'lab': .., 'total': ..}], sources_used, files found)."""
    files = discover_pt_files(specialty_dir, province)
    fees, sources_used = load_pt_dd_fees_from_files(files, known_codes, verbose)
    return fees, sources_used, files


def resolve_dd_role_values(values: dict[str, float]) -> tuple[float | None, float | None, float | None]:
    """Reduce a per-code {'prof', 'lab', 'total'} dict (see
    load_pt_dd_fees) to the (prof, lab, total) triple written to the DD
    sheet. When both Prof and Lab were found, Total is *recomputed* as
    their sum rather than trusted as extracted, even if a labeled Total
    column was also found: some guides' own Total column includes a small
    escalation/adjustment (seen in ON's and NS's sources, off by a couple
    percent) that the reference doesn't carry through, while every
    confirmed-correct source's labeled Total already equals Prof + Lab
    exactly anyway -- so recomputing it is a no-op for those and a fix for
    the rest. Total falls back to whatever was extracted only when Prof or
    Lab is missing, and -- for a source that doesn't distinguish Prof from
    Lab at all (just a single fallback 'total' value) -- that value is
    treated as the Prof fee too, matching this project's original
    single-value DD behavior for sources without labeled columns.

    The one exception to the recompute: when Prof and Lab arrive already
    EQUAL, they're not two independent additive amounts -- that shape only
    ever comes from extract_dd_codes_from_lines's own "exactly one value,
    printed once, mirrored into both roles" case (see there), where the
    source line structurally never had a separate (2x) Total token to
    begin with. Summing them there would silently double an already-correct
    Total (confirmed against NB's DD guide, e.g. "70150 ... 73.00 73.00"
    with no third number -- the real Total is 73, not 146).
    """
    prof = values.get("prof")
    lab = values.get("lab")
    total = values.get("total")
    both_numeric = isinstance(prof, (int, float)) and isinstance(lab, (int, float))
    if both_numeric and prof != lab:
        total = prof + lab
    elif total is None:
        if prof is not None:
            total = prof
        elif lab is not None:
            total = lab
    if prof is None and lab is None and total is not None:
        prof = total
    return prof, lab, total


# Maps a CDCP SP sub-specialty code to name fragments that identify a PT fee
# guide file as specific to that sub-specialty (e.g. "ON PE Fee Guide.xlsx",
# "MDA 2026 Periodontics....xlsx" both indicate Periodontics -> "PE"). Used
# to prefer a sub-specialty-specific guide's fee over a general/all-specialty
# guide's fee for the same code, since the same procedure code commonly has
# a genuinely different fee depending on which specialty bills it.
# CDCP's own two-letter specialty codes are counter-intuitive for these two
# in particular -- "PA" is Pediatric Dentistry and "PE" is Periodontics, the
# reverse of what the letters would suggest. Confirmed directly against
# PE's own CDCP price file: codes 01501/01502/01503, which PE's own PT fee
# guide explicitly labels "PER" (Periodontal exam codes), are assigned CDCP
# specialty "PE" -- and separately, codes 23411-23512 ("Primary Anterior"/
# "Primary Posterior" tooth-coloured restorations -- baby-tooth fillings,
# squarely Pediatric Dentistry's domain, not Periodontics') are assigned
# CDCP specialty "PA". (This was previously backwards here, which is what
# caused several provinces' SP sheets' Pediatric/Periodontics rows to
# apparently "swap" against the ground truth -- the ground truth was right,
# this mapping was wrong.)
SUBSPECIALTY_FILE_MARKERS: dict[str, list[str]] = {
    "EN": ["EN", "END", "ENDO", "ENDODONTIC", "ENDODONTICS"],
    "OS": ["OS", "OMS", "ORAL SURGERY", "ORAL AND MAXILLOFACIAL SURGERY", "MAXILLOFACIAL"],
    "PA": ["PA", "PED", "PEDIATRIC", "PEDIATRICS", "PAEDIATRIC", "PAEDIATRICS"],
    "PE": ["PE", "PER", "PERIODONTIC", "PERIODONTICS", "PERIODONTOLOGY"],
    "PR": ["PR", "PROSTHODONTIC", "PROSTHODONTICS"],
    "OM": ["OM", "ORAL MEDICINE"],
    "OP": ["OP", "ORAL PATHOLOGY"],
    "OR": ["OR", "ORT", "ORTHODONTIC", "ORTHODONTICS"],
    "RA": ["RA", "RADIOLOGY"],
    "AN": ["AN", "ANESTHESIA", "ANESTHESIOLOGY"],
}


# Reverse lookup from a specialty *label* (as it might appear in a source's
# own per-row specialty column, e.g. "END", "OMS", "PER") to the CDCP
# sub-specialty code it means -- reuses the same marker vocabulary as
# filename classification (see classify_file_specialty), since combined
# guides tend to abbreviate specialties the same way whether in a filename
# or a column value. "GP" and "LTC" aren't CDCP sub-specialties but share
# this column in combined guides, so they get their own pseudo-entries --
# purely so a GP/LTC row is recognized as *labeled* (and therefore excluded
# from matching any real sub-specialty, and from the "unlabeled" fallback
# bucket -- an LTC-context rate isn't the general rate for whichever
# sub-specialty the code also happens to belong to) rather than being
# mistaken for an unlabeled row (see find_row_specialty_column).
_ALL_SPECIALTY_MARKERS: dict[str, str] = {"GP": "GP", "LTC": "LTC"}
for _code, _markers in SUBSPECIALTY_FILE_MARKERS.items():
    for _marker in _markers:
        _ALL_SPECIALTY_MARKERS.setdefault(_marker, _code)

# Some SP guides label each row's specialty with a small internal
# reference *number* instead of a letter abbreviation -- confirmed in SK's
# Specialist Fee Guide, whose own numeric column uses 21-30 as
# section/category numbers, cross-checked against that guide's own section
# headings (e.g. 24 = "PERIODONTICS, ...", 29 = "ENDODONTIC SERVICES").
# Only mapped where a number corresponds to exactly one CDCP sub-specialty
# -- SK's "28" section merges Oral Medicine and Oral Pathology together
# under one number with no way to tell them apart from the number alone,
# and 21/22/30 are generic categories (Diagnostic, Radiology, Adjunctive)
# that aren't any one sub-specialty at all -- so those are deliberately
# left unmapped rather than guessed at (see _classify_specialty_cell).
_NUMERIC_SPECIALTY_MARKERS: dict[int, str] = {
    23: "PA",  # SK: "PEDIATRIC, ..." -- CDCP's "PA" is Pediatric Dentistry
    24: "PE",  # SK: "PERIODONTICS, ..." -- CDCP's "PE" is Periodontics
    25: "OS",  # SK: "ORAL&MAXILLOFACIAL SURG"
    26: "PR",  # SK: "PROSTHO SERV..."
    29: "EN",  # SK: "ENDODONTIC..."
}


def _classify_specialty_cell(cell) -> str | None:
    if isinstance(cell, (int, float)) and float(cell).is_integer():
        return _NUMERIC_SPECIALTY_MARKERS.get(int(cell))
    if not isinstance(cell, str):
        return None
    return _ALL_SPECIALTY_MARKERS.get(cell.strip().upper())


def find_row_specialty_column(rows) -> int | None:
    """Find a column whose values are recognizable CDCP sub-specialty labels
    (e.g. "END", "OMS", "PER", "GP") rather than a specialty-specific
    *file*. Some combined SP guides (e.g. PE's) list every sub-specialty's
    codes together in one file with a per-row specialty column instead of
    splitting into separate files/sections -- without identifying that
    column, the same code appearing under two different specialties (e.g.
    "25781" priced differently under "GP" and under "END") can't be told
    apart, and whichever row happens to come first wins regardless of which
    specialty was actually asked for.

    Returns the column index with the most matches among a sample of rows,
    if at least a meaningful fraction of that sample matches -- a low
    fraction means this source likely doesn't have a real specialty column
    at all (a stray "OR" or "PA" abbreviation elsewhere shouldn't count).
    """
    sample = rows[:500]
    if not sample:
        return None
    width = max((len(r) for r in sample), default=0)
    best_idx, best_count = None, 0
    for i in range(width):
        count = sum(1 for r in sample if i < len(r) and _classify_specialty_cell(r[i]) is not None)
        if count > best_count:
            best_idx, best_count = i, count
    if best_idx is not None and best_count >= len(sample) * 0.3:
        return best_idx
    return None


def classify_file_specialty(path: Path, province: str | None = None) -> set[str]:
    """Which CDCP sub-specialty code(s), if any, a PT fee guide filename
    identifies (e.g. "ON PE Fee Guide 2026.xlsx" -> {"PE"}). Empty set means
    the file isn't specific to one sub-specialty (e.g. a general/combined
    guide like "ON DA Fee Guide" or "BC LTC Fee Guide").

    If `province` is given, its aliases (see PROVINCE_ALIASES) are stripped
    from the start of the name before matching -- otherwise a filename like
    "PE GP SP LTC Fee Guide" (Prince Edward Island's combined guide) gets
    misread as Periodontics-specific, since "PE" is coincidentally both the
    province's abbreviation and the Periodontics specialty marker.

    Underscores and hyphens are normalized to spaces before any \\b-anchored
    matching below: \\b only draws a boundary where a "word" character
    (regex \\w, which includes "_") meets a non-word one, so "_" never
    creates one on its own -- \\bPR\\b silently never matches inside
    "ON_PR_FEE_GUIDE_2026", the actual naming convention this project's own
    PT fee guide files use (confirmed: every ON SP guide -- PA, PE, PR, OS,
    EN -- classified as unspecific/general under the un-normalized version,
    which fed them all into the same "no dedicated file" fallback pool
    regardless of which specialty they actually named, corrupting SP
    fee resolution for any code shared across specialties).
    """
    name = path.stem.upper().replace("_", " ").replace("-", " ")
    if province:
        for alias in PROVINCE_ALIASES.get(province, [province]):
            m = re.match(rf"^{re.escape(alias)}\b\s*", name, re.IGNORECASE)
            if m:
                name = name[m.end():]
                break
    matches = set()
    for code, markers in SUBSPECIALTY_FILE_MARKERS.items():
        for marker in markers:
            if re.search(rf"\b{re.escape(marker)}\b", name):
                matches.add(code)
                break
    return matches


# Filename markers for guides that are scoped to a specific care *setting*
# (e.g. Long Term Care) rather than a specific CDCP sub-specialty. These
# aren't a real CDCP specialty code, so a file scoped *only* to one of these
# settings must not be treated as a general/blanket fallback for every
# sub-specialty that lacks its own guide (a code's LTC-context rate isn't
# its regular-context rate). But some provinces publish one combined guide
# covering GP, SP, *and* LTC together (e.g. "PE GP SP LTC Fee Guide") --
# that file is a legitimate general SP source despite mentioning LTC, since
# it isn't LTC-exclusive.
_CONTEXT_RESTRICTED_MARKERS = ["LTC"]
_COMBINED_GUIDE_MARKERS = ["GP", "SP"]


def _is_context_restricted(path: Path) -> bool:
    # See classify_file_specialty for why "_"/"-" must be normalized to
    # spaces before \b-anchored matching against this project's actual
    # underscore-separated filenames.
    name = path.stem.upper().replace("_", " ").replace("-", " ")
    has_restricted_marker = any(re.search(rf"\b{marker}\b", name) for marker in _CONTEXT_RESTRICTED_MARKERS)
    if not has_restricted_marker:
        return False
    is_combined_guide = any(re.search(rf"\b{marker}\b", name) for marker in _COMBINED_GUIDE_MARKERS)
    return not is_combined_guide


# Some guides (so far only PE's combined GP+SP PDF) define a specialist's
# fee for an entire numeric code range as a flat percentage markup over the
# general practitioner's fee, rather than listing individual specialist
# fees at all -- e.g. Appendix F: "SERVICES PROVIDED BY A PROSTHODONTIST /
# SECTION 50000 - 59999 / FEES FOR ALL CODES 20% HIGHER THAN FOR GENERAL
# PRACTITIONER'S SUGGESTED FEE." Used as a per-code fallback in
# load_pt_fees_by_subspecialty for codes in that range with no
# specialist-specific rate elsewhere.
_SPECIALIST_NOUN_TO_SUBSPECIALTY: dict[str, str] = {
    "PROSTHODONTIST": "PR",
    "PERIODONTIST": "PE",
    "ENDODONTIST": "EN",
    "ORTHODONTIST": "OR",
    "ORAL AND MAXILLO-FACIAL SURGEON": "OS",
    "ORAL SURGEON": "OS",
    "PAEDIATRIC DENTIST": "PA",
    "PEDIATRIC DENTIST": "PA",
}
_MULTIPLIER_APPENDIX_RE = re.compile(
    r"SERVICES PROVIDED BY (?:A |AN |CERTIFIED )*([A-Za-z][A-Za-z \-]*?)\s*\n"
    r"SECTION\s*(\d+)\s*-\s*(\d+)\s*\n"
    r"FEES FOR ALL CODES\s*(\d+)\s*%\s*HIGHER THAN FOR GENERAL PRACTITIONER",
    re.IGNORECASE,
)


def find_specialist_multiplier_ranges(text: str) -> list[tuple[str, int, int, float]]:
    """Returns a list of (sub_specialty_code, range_low, range_high,
    multiplier) parsed from "FEES FOR ALL CODES N% HIGHER..." appendix
    blocks in `text` (see comment above)."""
    results = []
    for m in _MULTIPLIER_APPENDIX_RE.finditer(text):
        noun = re.sub(r"\s+", " ", m.group(1)).strip().upper()
        code = _SPECIALIST_NOUN_TO_SUBSPECIALTY.get(noun)
        if code is None:
            continue
        low, high, pct = int(m.group(2)), int(m.group(3)), int(m.group(4))
        results.append((code, low, high, 1 + pct / 100))
    return results


def _multiplier_for_code(code: str, ranges: list[tuple[int, int, float]]) -> float | None:
    try:
        n = int(code)
    except ValueError:
        return None
    for low, high, mult in ranges:
        if low <= n <= high:
            return mult
    return None


def load_pt_fees_by_subspecialty(
    specialty_dir: Path,
    province: str,
    codes_by_subspecialty: dict[str, set[str]],
    gp_specialty_dir: Path | None = None,
    verbose: bool = True,
):
    """SP-style resolution: like load_pt_fees, but aware that the same code
    can have a different fee under different sub-specialties. Files whose
    name identifies a specific sub-specialty (see classify_file_specialty)
    are tried first for that sub-specialty's codes, before falling back to
    general/unspecific files (which is all load_pt_fees does on its own).

    If `gp_specialty_dir` is given and a sub-specialty gets *no* matches at
    all from SP sources (i.e. no PT guide covers that specialty for this
    province -- e.g. BC has no EN/OM/OP/OS/PR/RA-specific guide), its codes
    fall back to the province's GP fee guide instead -- the reference file's
    own documented convention ("For any SP code without SP specific fee, GP
    fee is assumed"). This is deliberately an all-or-nothing trigger per
    sub-specialty, not a per-missing-code one: a sub-specialty with mostly
    good SP-specific coverage and a few individually-unmatched codes is more
    likely suffering an extraction gap in its own guide than a genuine
    absence of specialty-specific pricing, and guessing the GP fee for those
    stray gaps does more harm (contaminating otherwise-correct data) than
    leaving them "N/A".

    Before that blanket fallback, though, any codes covered by a documented
    percentage-markup rule (see find_specialist_multiplier_ranges -- so far
    only PE's guide) get GP fee x that specialty's stated multiplier
    instead of the plain GP fee, since that's a known, precise rule rather
    than a guess.

    Returns (fees dict keyed by (code, sub_specialty), sources_used, files found).
    """
    files = discover_pt_files(specialty_dir, province)
    general_files = [f for f in files
                      if not classify_file_specialty(f, province) and not _is_context_restricted(f)]

    multiplier_ranges: dict[str, list[tuple[int, int, float]]] = {}
    for f in files:
        if f.suffix.lower() != ".pdf":
            continue
        try:
            reader = pypdf.PdfReader(str(f))
            text = "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception:
            continue
        for code, low, high, mult in find_specialist_multiplier_ranges(text):
            multiplier_ranges.setdefault(code, []).append((low, high, mult))

    fees: dict[tuple[str, str], float] = {}
    sources_used: list[tuple[str, int]] = []
    for sub_specialty, codes in codes_by_subspecialty.items():
        specific_files = [f for f in files if sub_specialty in classify_file_specialty(f, province)]
        candidate_files = specific_files + general_files if specific_files else general_files
        sub_fees, sub_sources = load_pt_fees_from_files(
            candidate_files, codes, verbose, target_specialty=sub_specialty
        )
        for code, fee in sub_fees.items():
            fees[(code, sub_specialty)] = fee
        sources_used.extend(sub_sources)

        ranges = multiplier_ranges.get(sub_specialty)
        if ranges and gp_specialty_dir is not None:
            # Applied to every code in range, not just ones sub_fees missed
            # -- a documented multiplier rule overrides even a value
            # sub_fees *did* find, since that match is usually the generic
            # PDF reader (not specialty-aware, unlike the spreadsheet/csv/
            # docx readers -- see extract_codes_from_rows) grabbing the
            # wrong, non-specialist section of the same combined document
            # rather than genuinely finding this specialty's own rate.
            in_range = {c for c in codes if _multiplier_for_code(c, ranges) is not None}
            if in_range:
                gp_fees, _, _ = load_pt_fees(gp_specialty_dir, province, in_range, verbose=False)
                applied = 0
                for code in in_range:
                    gp_fee = gp_fees.get(code)
                    mult = _multiplier_for_code(code, ranges)
                    if gp_fee is None or mult is None:
                        continue
                    # gp_fee can now be a no-fixed-fee marker string (e.g.
                    # "I.C.", see _marker_text) instead of a number -- there's
                    # no numeric base to apply the specialist markup to, so
                    # carry the marker text through unchanged rather than
                    # crash trying to multiply it.
                    fees[(code, sub_specialty)] = (
                        gp_fee * mult if isinstance(gp_fee, (int, float)) else gp_fee
                    )
                    applied += 1
                if applied:
                    sources_used.append((f"GP fee x specialist markup ({applied})", applied))

        if not sub_fees and gp_specialty_dir is not None:
            gp_fees, gp_sources, _ = load_pt_fees(gp_specialty_dir, province, codes, verbose=False)
            for code, fee in gp_fees.items():
                if (code, sub_specialty) not in fees:
                    fees[(code, sub_specialty)] = fee
            if gp_fees:
                sources_used.extend((f"{label} (as GP fallback)", n) for label, n in gp_sources)

    return fees, sources_used, files
