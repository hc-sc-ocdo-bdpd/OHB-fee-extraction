"""
Loaders for the CDCP price files, one per specialty "shape":

- GP and DH: the 'Specialty' column literally equals the specialty code
  ("GP", "DH"), so one code maps to one fee.
- SP: the 'Specialty' column holds a *sub-specialty* code (PA, EN, OS, ...)
  instead of the literal "SP" -- and the same procedure code commonly
  repeats under several different sub-specialties with different fees. So SP
  rows are returned as a flat list of (code, sub_specialty, fee) rather than
  a code -> fee dict.
- DD: has two separate fee columns (Provider Fee = professional fee, and
  Internal Lab Fee), which the output sheet keeps as separate Prof/Lab/Combo
  columns.
"""

import re
from functools import lru_cache
from pathlib import Path
from typing import NamedTuple

import openpyxl

import config
from fee_extraction import normalize_code

# Yukon's CDCP price file is named/coded "YK" internally (both the filename
# and the 'Province' column value), while everywhere else in this project
# (province lists, PT fee guide folder names, the template's output labels)
# uses "YT". Without this, looking up "YT" finds no file at all.
CDCP_PROVINCE_ALIASES: dict[str, str] = {"YT": "YK"}

# Some CDCP price files are internally inconsistent about their own
# 'Province' column: QC's file uses "QC" for every row except a block of
# P-series GP codes (P0500-P1700) near the end of the sheet, which use "PQ"
# ("Province de Québec") instead -- confirmed by inspecting the file
# directly (row 342 is "QC"/99111, row 343 onward is "PQ"/P0500...). Without
# this, those 12 rows' Province never equals the "QC" we're filtering for,
# so load_cdcp_simple_fees silently drops them -- not an extraction bug,
# just an inconsistent label within one file that needs to be tolerated.
# The canonical `province` argument (not the file's own "PQ" label) is what
# ends up written to the output sheet's PT column, since these loaders only
# return code->fee data -- build_fee_comparison.py supplies the label --
# so accepting "PQ" here doesn't leak an inconsistent label into the output.
CDCP_ROW_PROVINCE_ALIASES: dict[str, set[str]] = {
    "QC": {"QC", "PQ"},
}


def _row_matches_province(row_province, province: str, data_province: str) -> bool:
    return row_province in CDCP_ROW_PROVINCE_ALIASES.get(province, {data_province})


@lru_cache(maxsize=None)
def _find_cdcp_price_file(cdcp_dir_str: str, file_province: str) -> Path | None:
    """Locate one province's CDCP price file, tolerating year-to-year changes
    in how the filename is spelled (see config.cdcp_price_file_names). Falls
    back to a year+province-anchored glob so a naming variant we haven't seen
    still resolves instead of silently yielding no CDCP data at all."""
    cdcp_dir = Path(cdcp_dir_str)
    for name in config.cdcp_price_file_names(file_province):
        candidate = cdcp_dir / name
        if candidate.exists():
            return candidate
    matches = sorted(p for p in cdcp_dir.glob(config.cdcp_price_file_glob(file_province))
                     if not p.name.startswith("~$"))
    return matches[0] if matches else None


@lru_cache(maxsize=None)
def _cached_workbook(path_str: str):
    """One CDCP price file is read once per specialty sheet (GP/DH/SP/DD) for
    the same province -- four full workbook loads of the same file. Cached so
    it's opened once per run."""
    return openpyxl.load_workbook(path_str, data_only=True)


# Accepted spellings for each column the loaders need, in preference order.
# Matched case-insensitively and ignoring surrounding/repeated whitespace (see
# _normalize_header), so "provider fee", "Provider  Fee" and "PROVIDER FEE"
# all resolve. A year-qualified variant ("2025 Provider Fee") is accepted too,
# since some price files label the fee column with its rate year.
COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "Province": ("province", "pt", "prov"),
    "Specialty": ("specialty", "speciality", "spec"),
    "Procedure Code": ("procedure code", "code", "proc code", "dac code"),
    "Provider Fee": (
        "provider fee", "providerfee", "fee", "cdcp fee", "cdcp provider fee",
        "provider fees", "amount",
    ),
    "Internal Lab Fee": (
        "internal lab fee", "internallabfee", "lab fee", "internal lab",
        "commercial lab fee",
    ),
}


class SheetMeta(NamedTuple):
    """Where the header was found and what it said -- carried alongside the
    column index purely so an error can name the offending file/sheet."""
    header_row: int
    header: list
    ctx: str


def _normalize_header(value) -> str:
    return re.sub(r"\s+", " ", str(value)).strip().lower() if value is not None else ""


def _build_column_index(header: list) -> dict[str, int]:
    """Map each logical column name to its position in `header`.

    Resolved by alias rather than exact string so a price file that spells a
    column differently still loads. Year-qualified labels ("2025 Provider
    Fee") match their unqualified alias with the year stripped."""
    normalized = [_normalize_header(h) for h in header]
    stripped = [re.sub(r"\b(19|20)\d{2}\b", " ", n).strip() for n in normalized]
    stripped = [re.sub(r"\s+", " ", n) for n in stripped]

    idx: dict[str, int] = {}
    for logical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            for i, name in enumerate(normalized):
                if name == alias and i not in idx.values():
                    idx[logical] = i
                    break
            if logical in idx:
                break
            for i, name in enumerate(stripped):
                if name == alias and i not in idx.values():
                    idx[logical] = i
                    break
            if logical in idx:
                break
    return idx


def require_columns(idx: dict[str, int], needed: tuple[str, ...], ctx: str, header: list) -> None:
    """Raise a message that says which file/sheet is wrong and what it
    actually contains -- a bare KeyError('Provider Fee') from deep in a row
    loop says nothing about which of 13 provinces' files is at fault."""
    missing = [n for n in needed if n not in idx]
    if not missing:
        return
    found = [str(h) for h in header if h is not None]
    raise KeyError(
        f"{ctx}: could not find column(s) {', '.join(missing)}. "
        f"Columns present: {found}. "
        f"Add the spelling used here to COLUMN_ALIASES in cdcp_loader.py."
    )


def _find_header_row(ws, max_scan: int = 5):
    """Row index (1-based) of the header, and its values.

    Usually row 1, but a price file can carry a title/blank row above it. The
    header is the first row within `max_scan` that resolves at least the
    Province/Specialty/Procedure Code triple."""
    best = (1, [])
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan), start=1):
        header = [c.value for c in row]
        if r == 1:
            best = (1, header)
        idx = _build_column_index(header)
        if all(k in idx for k in ("Province", "Specialty", "Procedure Code")):
            return r, header
        if "Procedure Code" in idx and "Provider Fee" in idx:
            return r, header
    return best


def _load_sheet(cdcp_dir: Path, province: str, sheet_name: str):
    file_province = CDCP_PROVINCE_ALIASES.get(province, province)
    path = _find_cdcp_price_file(str(cdcp_dir), file_province)
    if path is None:
        return None
    wb = _cached_workbook(str(path))
    if sheet_name not in wb.sheetnames:
        return None
    ws = wb[sheet_name]
    header_row, header = _find_header_row(ws)
    idx = _build_column_index(header)
    meta = SheetMeta(
        header_row=header_row,
        header=header,
        ctx=f"{path.name} [{sheet_name}]",
    )
    return ws, idx, file_province, meta


def load_cdcp_simple_fees(
    cdcp_dir: Path, province: str, sheet_name: str, specialty: str, require_fee: bool = True
) -> dict[str, float | None]:
    """GP/DH-style: one 2026 fee per procedure code.

    `require_fee` controls what happens to codes with no Provider Fee in the
    CDCP file (e.g. an "I.C." / individually-costed service). DH's sheet in
    the template excludes such codes entirely; GP's sheet instead keeps them
    as a row with an "N/A" fee. Pass require_fee=False to match GP's
    convention.
    """
    result = _load_sheet(cdcp_dir, province, sheet_name)
    if result is None:
        return {}
    ws, idx, data_province, meta = result
    require_columns(idx, ("Province", "Specialty", "Procedure Code", "Provider Fee"),
                    meta.ctx, meta.header)

    fees: dict[str, float | None] = {}
    for row in ws.iter_rows(min_row=meta.header_row + 1, values_only=True):
        if not _row_matches_province(row[idx["Province"]], province, data_province) or row[idx["Specialty"]] != specialty:
            continue
        fee = row[idx["Provider Fee"]]
        if fee is None and require_fee:
            continue
        code = normalize_code(row[idx["Procedure Code"]])
        if code is None:
            continue
        fees[code] = float(fee) if fee is not None else None
    return fees


def load_cdcp_sp_rows(cdcp_dir: Path, province: str, require_fee: bool = True) -> list[tuple[str, str, float | None]]:
    """SP-style: every (code, sub-specialty) row for the province, since the
    same code can have a different fee under different sub-specialties.
    See load_cdcp_simple_fees for `require_fee`."""
    result = _load_sheet(cdcp_dir, province, "SP")
    if result is None:
        return []
    ws, idx, data_province, meta = result
    require_columns(idx, ("Province", "Specialty", "Procedure Code", "Provider Fee"),
                    meta.ctx, meta.header)

    rows = []
    for row in ws.iter_rows(min_row=meta.header_row + 1, values_only=True):
        if not _row_matches_province(row[idx["Province"]], province, data_province):
            continue
        fee = row[idx["Provider Fee"]]
        if fee is None and require_fee:
            continue
        raw_code = row[idx["Procedure Code"]]
        code = normalize_code(raw_code)
        if code is None:
            # A handful of QC SP rows use an unusual composite code instead
            # of a plain 5-digit/alphanumeric one (e.g. "PS407/12250 age
            # 0-11", a pediatric age-based surcharge code) -- normalize_code
            # rejects that shape entirely, silently dropping the row. Since
            # there's no simpler canonical form to reduce it to, fall back
            # to the raw string as-is (matching this project's reference
            # output, which does the same) rather than losing the row.
            code = str(raw_code).strip() if raw_code else None
        sub_specialty = row[idx["Specialty"]]
        if not code or not sub_specialty:
            continue
        rows.append((code, sub_specialty, float(fee) if fee is not None else None))
    return rows


def load_cdcp_dd_fees(cdcp_dir: Path, province: str) -> dict[str, tuple[float, float]]:
    """DD-style: (professional fee, internal lab fee) per procedure code."""
    result = _load_sheet(cdcp_dir, province, "DD")
    if result is None:
        return {}
    ws, idx, data_province, meta = result
    require_columns(idx, ("Province", "Specialty", "Procedure Code", "Provider Fee",
                          "Internal Lab Fee"), meta.ctx, meta.header)

    fees: dict[str, tuple[float, float]] = {}
    for row in ws.iter_rows(min_row=meta.header_row + 1, values_only=True):
        if not _row_matches_province(row[idx["Province"]], province, data_province) or row[idx["Specialty"]] != "DD":
            continue
        prof_fee = row[idx["Provider Fee"]]
        if prof_fee is None:
            continue
        code = normalize_code(row[idx["Procedure Code"]])
        if code is None:
            continue
        lab_fee = row[idx["Internal Lab Fee"]] or 0
        fees[code] = (float(prof_fee), float(lab_fee))
    return fees
