"""
Load rate-year fee data, for one sheet, as plain dictionaries.

This is the seam between "where the numbers come from" and "how the workbook
is laid out". Everything here answers a single question -- *what are year Y's
CDCP and PT fees for this specialty?* -- and answers it the same way whichever
year is asked for. build_fee_comparison then merges the two years into one row
per code.

Nothing in here knows about columns, styles or formulas, and nothing in
sheet_builders knows where a fee came from. That split is what makes the pair
of years a parameter rather than a rewrite: to build a different pair, change
config.PAST_YEAR / config.CURRENT_YEAR and this module loads those instead.

Extraction itself is deliberately untouched. fee_extraction and cdcp_loader
are used exactly as they are; this module only decides which year's folders
they are pointed at (via year_context) and which codes it asks them about.

WHY LOADING IS TWO PHASES
    A row exists if EITHER year's CDCP file lists that code -- so a code the
    CDCP only started pricing in the newer year is still a row, showing the
    older year's fee for comparison. But the PT fee guides are only asked
    about codes we name, so if each year asked only about its OWN CDCP codes,
    that row's older-year PT column would come back "N/A" even when the older
    guide prices it perfectly well. (ON DD 31511 was exactly this: priced at
    804 in the 2025 guide, absent from the 2025 CDCP file.)

    So: phase 1 reads both years' CDCP files and unions the codes; phase 2
    asks each year's guides about that union. CDCP fees stay strictly
    per-year -- only the *question put to the guides* is widened. Each guide
    remains the authority on whether it prices a code; a year that doesn't
    still yields "N/A".

Each loader returns a SheetData whose `cdcp` and `pt` dicts share one key
shape per sheet, so the two years can be merged by key:

    DH / GP   (province, code)            fee or None
    QC GP     (code,)                     fee or None
    SP        (province, sub, code)       fee or None
    QC SP     (sub, code)                 fee or None
    DD        (province, code)            CDCP: (prof, lab)
                                          PT:   (prof, lab, combo)
"""

from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path

import config
import fee_extraction
from cdcp_loader import load_cdcp_dd_fees, load_cdcp_simple_fees, load_cdcp_sp_rows
from fee_extraction import (
    load_pt_dd_fees, load_pt_fees, load_pt_fees_by_subspecialty,
    resolve_dd_role_values,
)

# Provinces/territories in the order they appear in the output, matching the
# template's Claim Lines sheet. A province with no CDCP data for a specialty
# in either year is skipped automatically, based on what the price files
# actually contain.
ALL_PROVINCES = ["AB", "BC", "MB", "NB", "NL", "NS", "ON", "PE", "QC", "SK", "NT", "NU", "YT"]
NON_QC_PROVINCES = [p for p in ALL_PROVINCES if p != "QC"]

# The sheets a build produces, in output order. QC is split out of GP and SP
# because its CDCP and PT data use different code sets, exactly as the
# template has it.
SHEETS = ("DH", "GP", "QC GP", "SP", "QC SP", "DD")

# Which PT fee-guide subfolder each sheet reads from.
_GUIDE_FOLDER = {"DH": "DH", "GP": "GP", "QC GP": "GP",
                 "SP": "SP", "QC SP": "SP", "DD": "DD"}

# The one-fee-per-code sheets: CDCP specialty, which provinces, whether a code
# with no Provider Fee is still a row, and whether the province is part of the
# key. (QC GP's key is the code alone -- its sheet is single-province.)
_SIMPLE_SHEETS = {
    "DH":    ("DH", ALL_PROVINCES,     True,  True),
    "GP":    ("GP", NON_QC_PROVINCES,  False, True),
    "QC GP": ("GP", ["QC"],            False, False),
}

# The sub-specialty sheets: which provinces, and whether the province is part
# of the key.
_SP_SHEETS = {
    "SP":    (NON_QC_PROVINCES, True),
    "QC SP": (["QC"],           False),
}


@contextmanager
def year_context(year: int):
    """Point every year-derived lookup at `year` for the duration of a load.

    cdcp_loader builds its price-file names from config, and fee_extraction
    decides which DD columns are "prior year" from config -- both ask for the
    year without being told it. Rather than thread a year parameter through
    two modules that are correct and tested, this sets the year they see,
    for exactly as long as one year's load takes, and restores it after.

    Nested and exception-safe: whatever was set before is put back.
    """
    previous_config_year = config.set_active_year(year)
    previous_guide_year = fee_extraction._CURRENT_GUIDE_YEAR
    fee_extraction._CURRENT_GUIDE_YEAR = year
    try:
        yield year
    finally:
        config.set_active_year(previous_config_year)
        fee_extraction._CURRENT_GUIDE_YEAR = previous_guide_year


@dataclass
class SheetData:
    """One year's fees for one sheet, plus what it took to get them."""
    sheet: str
    year: int
    cdcp: dict = field(default_factory=dict)
    pt: dict = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)

    @property
    def keys(self) -> set:
        """Every identity this year knows about, from either source."""
        return set(self.cdcp) | set(self.pt)

    def summary(self) -> str:
        return (f"{self.year} {self.sheet}: {len(self.cdcp)} CDCP codes, "
                f"{len(self.pt)} PT fees")


def _guides_dir(sheet: str, year: int) -> Path:
    return config.pt_guides_dir(year) / _GUIDE_FOLDER[sheet]


def _in_province_order(scope: dict) -> list[str]:
    """Scope provinces in the output's own order, so the provenance notes a
    build prints read the same way every run regardless of which year first
    contributed a province."""
    return [p for p in ALL_PROVINCES if p in scope]


def _note_pt(notes: list[str], province: str, label: str, sources, files_found) -> None:
    described = ", ".join(f"{name} ({n})" for name, n in sources) or "no PT fee guide found"
    notes.append(f"  {province} {label}: {described}")
    if not files_found:
        notes.append(f"    Note: no PT fee guide file for {province}/{label}; PT fees will be 'N/A'.")


# ---------------------------------------------------------------------------
# Phase 1: what the CDCP price files say, per year
# ---------------------------------------------------------------------------

def load_cdcp(sheet: str, year: int) -> dict:
    """One year's CDCP data for one sheet, grouped by province.

    A province absent from that year's price file is simply not a key, which
    is what lets the caller tell "no data for this province this year" from
    "data, but no fee".

    Shapes:  DH/GP/QC GP  {province: {code: fee}}
             SP/QC SP     {province: [(code, sub, fee), ...]}
             DD           {province: {code: (prof, lab)}}
    """
    with year_context(year):
        root = config.cdcp_dir(year)
        by_province: dict[str, object] = {}

        if sheet in _SIMPLE_SHEETS:
            specialty, provinces, require_fee, _ = _SIMPLE_SHEETS[sheet]
            for province in provinces:
                fees = load_cdcp_simple_fees(root, province, specialty, specialty,
                                             require_fee=require_fee)
                if fees:
                    by_province[province] = fees
        elif sheet in _SP_SHEETS:
            for province in _SP_SHEETS[sheet][0]:
                rows = load_cdcp_sp_rows(root, province, require_fee=False)
                if rows:
                    by_province[province] = rows
        else:
            for province in ALL_PROVINCES:
                fees = load_cdcp_dd_fees(root, province)
                if fees:
                    by_province[province] = fees
        return by_province


def _code_scope(sheet: str, cdcp_by_year: dict[int, dict]) -> dict:
    """Every code any year's CDCP file prices, per province -- the set of
    codes the guides are asked about, for every year.

    Union rather than per-year, because the output row set is a union: a code
    only the newer CDCP prices is still a row, and its older-year PT fee is
    the whole point of the comparison.

    {province: {codes}} normally, {province: {sub: {codes}}} for SP.
    """
    scope: dict = {}
    for by_province in cdcp_by_year.values():
        for province, data in by_province.items():
            if sheet in _SP_SHEETS:
                per_sub = scope.setdefault(province, {})
                for code, sub, _fee in data:
                    per_sub.setdefault(sub, set()).add(code)
            else:
                scope.setdefault(province, set()).update(data)
    return scope


# ---------------------------------------------------------------------------
# Phase 2: what the PT fee guides say, for that scope
# ---------------------------------------------------------------------------

def _build_simple(sheet: str, year: int, cdcp: dict, scope: dict) -> SheetData:
    """DH / GP / QC GP: one fee per (province, code)."""
    data = SheetData(sheet=sheet, year=year)
    keyed_by_province = _SIMPLE_SHEETS[sheet][3]
    guides = _guides_dir(sheet, year)

    def key(province, code):
        return (province, code) if keyed_by_province else (code,)

    for province in _in_province_order(scope):
        codes = scope[province]
        pt_fees, sources, files_found = load_pt_fees(guides, province, set(codes), verbose=False)
        _note_pt(data.notes, province, sheet, sources, files_found)

        for code, fee in cdcp.get(province, {}).items():
            data.cdcp[key(province, code)] = fee
        for code, fee in pt_fees.items():
            data.pt[key(province, code)] = fee
    return data


def _build_sp(sheet: str, year: int, cdcp: dict, scope: dict) -> SheetData:
    """SP / QC SP: a fee per (province, sub-specialty, code) -- the same code
    is priced differently under different sub-specialties."""
    data = SheetData(sheet=sheet, year=year)
    keyed_by_province = _SP_SHEETS[sheet][1]
    guides = _guides_dir(sheet, year)
    gp_guides = config.pt_guides_dir(year) / "GP"

    def key(province, sub, code):
        return (province, sub, code) if keyed_by_province else (sub, code)

    for province in _in_province_order(scope):
        codes_by_sub = scope[province]
        pt_fees, sources, files_found = load_pt_fees_by_subspecialty(
            guides, province, codes_by_sub, gp_specialty_dir=gp_guides, verbose=False)
        _note_pt(data.notes, province, sheet, sources, files_found)

        for code, sub, fee in cdcp.get(province, []):
            data.cdcp[key(province, sub, code)] = fee
        for (code, sub), fee in pt_fees.items():
            data.pt[key(province, sub, code)] = fee
    return data


def _build_dd(year: int, cdcp: dict, scope: dict) -> SheetData:
    """DD: professional and lab fees are separate columns, so each side is a
    tuple rather than a single number."""
    data = SheetData(sheet="DD", year=year)
    guides = _guides_dir("DD", year)

    for province in _in_province_order(scope):
        codes = scope[province]
        role_fees, sources, files_found = load_pt_dd_fees(
            guides, province, set(codes), verbose=False)
        _note_pt(data.notes, province, "DD", sources, files_found)

        for code, (prof, lab) in cdcp.get(province, {}).items():
            data.cdcp[(province, code)] = (prof, lab)
        for code in codes:
            triple = resolve_dd_role_values(role_fees.get(code, {}))
            if triple != (None, None, None):
                data.pt[(province, code)] = triple
    return data


def load_pt(sheet: str, year: int, cdcp: dict, scope: dict) -> SheetData:
    """One year's SheetData, given that year's CDCP data and the code scope."""
    with year_context(year):
        if sheet in _SIMPLE_SHEETS:
            return _build_simple(sheet, year, cdcp, scope)
        if sheet in _SP_SHEETS:
            return _build_sp(sheet, year, cdcp, scope)
        return _build_dd(year, cdcp, scope)


# ---------------------------------------------------------------------------
# Entry points
# ---------------------------------------------------------------------------

def load_sheet(sheet: str, year: int) -> SheetData:
    """One sheet's fees for one year, on its own.

    The scope is that year's own CDCP codes -- there's no other year to widen
    it with. Use load_both_years for a build.
    """
    _check_sheet(sheet)
    cdcp = load_cdcp(sheet, year)
    return load_pt(sheet, year, cdcp, _code_scope(sheet, {year: cdcp}))


def load_both_years(sheet: str) -> dict[int, SheetData]:
    """The same sheet for both configured years, keyed by year.

    Both years' CDCP files are read first so each year's guides can be asked
    about every code either year prices -- see WHY LOADING IS TWO PHASES.

    A year whose data folders aren't there yet still returns a SheetData --
    an empty one -- so the build writes "N/A" in that year's columns and
    carries on, rather than failing outright.
    """
    _check_sheet(sheet)
    cdcp_by_year = {year: load_cdcp(sheet, year) for year in config.years()}
    scope = _code_scope(sheet, cdcp_by_year)
    return {year: load_pt(sheet, year, cdcp_by_year[year], scope)
            for year in config.years()}


def _check_sheet(sheet: str) -> None:
    if sheet not in SHEETS:
        raise ValueError(f"unknown sheet {sheet!r}; expected one of {', '.join(SHEETS)}")


def merged_keys(by_year: dict[int, SheetData]) -> list:
    """Every identity either year knows about, in a stable order.

    Union rather than intersection on purpose: a code priced in only one of
    the two years is still a real row -- it shows as a value in that year and
    "N/A" in the other, which is exactly the comparison the workbook is for.
    """
    keys = set()
    for data in by_year.values():
        keys |= data.keys
    return sorted(keys, key=lambda k: tuple(str(part) for part in k))
