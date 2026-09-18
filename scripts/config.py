"""
Single place to configure which rate year the whole pipeline runs against.

Change `YEAR` below and nothing else -- every folder path, input filename
pattern, output filename, and the prior-year column filter in
fee_extraction all derive from it.

Expected layout (the convention actually in use):

    <DATA_DIR>/
        2025/
            2025_CDCP Fees/
            2025_Fee Guides/
            2025_Output/              <- output (created on demand)
        2026/
            2026_CDCP Fees/
            2026_Fee Guides/
            2026_Output/

Each directory is resolved by trying a list of candidates in order and
taking the first that exists, so the older flat layout
("Input_CDCP price files" / "Input_PT association fee guides" directly
under Data/) still works too. To support a new convention, add it to the
relevant candidate list rather than editing anything downstream.
"""

import os
import re
from pathlib import Path

# ---------------------------------------------------------------------------
# THE TWO KNOBS: the pair of rate years the output carries, side by side.
#
# Both are populated from their own year's data folders -- PAST_YEAR fills the
# "<year> CDCP Fee"/"<year> PT Fee" columns on the left, CURRENT_YEAR the ones
# on the right. To build a different pair next time, change these two numbers
# and nothing else; every folder, filename, header label and output path is
# derived from them (see year_dir, cdcp_dir, pt_guides_dir, output_dir).
#
# They do not have to be consecutive, and a pair may be built before one
# year's guides have arrived -- a year with no data folder simply comes out
# "N/A" in its columns rather than failing the build.
# ---------------------------------------------------------------------------
PAST_YEAR = 2025
CURRENT_YEAR = 2026
# ---------------------------------------------------------------------------

# Which of the two years a load is currently working on. Set by
# fee_dataset.year_context() for the duration of one year's extraction, and
# read by _y() below so that year-derived lookups -- CDCP price filenames in
# particular -- resolve to the year being read rather than to a fixed one.
# This is what lets cdcp_loader and fee_extraction stay exactly as they are
# while the build walks two years: they ask config for "the" year, and config
# answers with whichever year is being loaded right now.
_ACTIVE_YEAR: int | None = None

# The template workbook supplies the header rows, styling, column layout and
# the Claim Lines sheet that the generated output is built on. It is the SAME
# file for every year -- only the year labels in its headers get rewritten
# (see build_fee_comparison.build_year_label_map) -- so it can live outside
# any per-year folder. Set this to an explicit path, or leave None to search
# the usual locations (see template_workbook).
TEMPLATE_FILE = None

# Sheets a workbook must contain to be usable as the template. Used to reject
# an unrelated workbook that merely happens to match the filename search --
# without this check, picking e.g. "2025 DD Fee Comparisons - CDCP vs PT.xlsx"
# failed later and much less clearly, with KeyError: 'Worksheet Claim Lines
# does not exist.'
REQUIRED_TEMPLATE_SHEETS = ("Claim Lines", "DH", "GP", "QC GP", "SP", "QC SP", "DD")


def past_year(year: int = None) -> int:
    return PAST_YEAR


def current_year(year: int = None) -> int:
    return CURRENT_YEAR


def years() -> tuple[int, int]:
    """The two years this build populates, oldest first."""
    return (PAST_YEAR, CURRENT_YEAR)


def set_active_year(year: int | None) -> int | None:
    """Point year-derived lookups at `year`; returns the previous setting so
    a caller can restore it. Prefer fee_dataset.year_context()."""
    global _ACTIVE_YEAR
    previous, _ACTIVE_YEAR = _ACTIVE_YEAR, year
    return previous

# Root that holds the per-year folders. Defaults to "Data" beside this repo
# (i.e. <repo parent>/Data). Override without editing this file by setting the
# OHB_DATA_DIR environment variable -- useful when the data lives somewhere
# else entirely, e.g. a OneDrive folder:
#   Windows : set OHB_DATA_DIR=C:\Users\...\Desktop\OHB\Data
#   bash    : export OHB_DATA_DIR="/mnt/c/Users/.../Desktop/OHB/Data"
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.environ.get("OHB_DATA_DIR") or (BASE_DIR / "Data"))


def _y(year: int | None) -> int:
    """Resolve an optional year argument.

    An explicit year always wins. Otherwise it's the year currently being
    loaded (see _ACTIVE_YEAR), and failing that the current year -- so a
    caller that never mentions a year still behaves sensibly."""
    if year is not None:
        return year
    return _ACTIVE_YEAR if _ACTIVE_YEAR is not None else CURRENT_YEAR


# Backwards-compatible alias. Modules that predate the two-year build ask for
# config.YEAR; it now means "the year being loaded, else the current one".
def __getattr__(name):
    if name == "YEAR":
        return _y(None)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def year_dir(year: int = None) -> Path:
    """The per-year parent folder, e.g. <DATA_DIR>/2026.

    Falls back to DATA_DIR itself for the older flat layout, where the
    year-named folders sat directly under Data/ with no year folder between.
    """
    y = _y(year)
    nested = DATA_DIR / str(y)
    return nested if nested.exists() else DATA_DIR


def _first_existing(candidates: list[Path]) -> Path:
    """First candidate that exists, else the first candidate.

    Falling back to candidates[0] rather than raising keeps a partially
    present year usable -- the caller reports a clear "not found" (see the
    startup banner in build_fee_comparison) instead of failing at import.
    """
    for path in candidates:
        if path.exists():
            return path
    return candidates[0]


def cdcp_dir(year: int = None) -> Path:
    """Directory holding that year's CDCP price files."""
    y = _y(year)
    yd = year_dir(y)
    return _first_existing([
        yd / f"{y}_CDCP Fees",
        yd / f"{y}_CDCP price files",
        yd / "Input_CDCP price files",
        DATA_DIR / f"{y}_CDCP Fees",
        DATA_DIR / "Input_CDCP price files",
    ])


def pt_guides_dir(year: int = None) -> Path:
    """Directory holding that year's PT association fee guides, which in turn
    contains the per-specialty subfolders (GP / DH / SP / DD)."""
    y = _y(year)
    yd = year_dir(y)
    return _first_existing([
        yd / f"{y}_Fee Guides",
        yd / f"{y}_PT association fee guides",
        yd / "Input_PT association fee guides",
        DATA_DIR / f"{y}_Fee Guides",
        DATA_DIR / "Input_PT association fee guides",
    ])


def output_dir(year: int = None) -> Path:
    """Directory the generated workbook, the ground-truth workbook and the
    mismatch reports all live in. Unlike the input directories this one is
    created if missing (see build_fee_comparison) -- a year being built for
    the first time won't have it yet."""
    y = _y(year)
    yd = year_dir(y)
    return _first_existing([
        yd / f"{y}_Output",
        yd / f"{y}_Fee Comparison",
        yd / f"Output_{y} Fee Comparison",
        DATA_DIR / f"Output_{y} Fee Comparison",
        DATA_DIR / f"{y}_Fee Comparison",
    ])


def cdcp_price_file_names(province: str, year: int = None) -> list[str]:
    """Candidate filenames for one province's CDCP price file, most likely
    first. Tried in order by cdcp_loader, which falls back to a glob if none
    match -- the exact spelling has varied ("2026 - CDCP PRICE FILE - ON.xlsx"
    vs "2026__CDCP_PRICE_FILE__ON.xlsx")."""
    y = _y(year)
    return [
        f"{y} - CDCP PRICE FILE - {province}.xlsx",
        f"{y}_CDCP_PRICE_FILE_{province}.xlsx",
        f"{y}__CDCP_PRICE_FILE__{province}.xlsx",
        f"{y} CDCP PRICE FILE {province}.xlsx",
        f"{y} - CDCP Price File - {province}.xlsx",
    ]


def cdcp_price_file_glob(province: str, year: int = None) -> str:
    """Last-resort pattern for a province's CDCP price file, used when none of
    cdcp_price_file_names matched. Anchored on both the year and the province
    so it can't pick up a different province's file."""
    return f"*{_y(year)}*{province}*.xlsx"


def _has_required_sheets(path: Path) -> bool:
    """True if `path` is a workbook containing every sheet the build needs.

    Opened read-only so this costs almost nothing -- only the sheet directory
    is parsed, not the cell data."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
        try:
            names = set(wb.sheetnames)
        finally:
            wb.close()
    except Exception:
        return False
    return all(s in names for s in REQUIRED_TEMPLATE_SHEETS)


def available_years() -> list[int]:
    """Year folders present under DATA_DIR, newest first."""
    if not DATA_DIR.exists():
        return []
    years = []
    for p in DATA_DIR.iterdir():
        if p.is_dir() and p.name.isdigit() and len(p.name) == 4:
            years.append(int(p.name))
    return sorted(years, reverse=True)


def templates_dir() -> Path:
    """Folder holding the workbook template, outside any per-year folder.

    The template is the same file for every year -- only its header year
    labels are rewritten -- so it doesn't belong to 2025 or 2026 and is
    better kept somewhere deliberate than picked out of a year's output
    folder alongside generated workbooks. This folder is searched FIRST, so
    a template put here always wins over one found anywhere else.
    """
    return _first_existing([
        DATA_DIR / "templates",
        DATA_DIR / "Templates",
        DATA_DIR / "_templates",
    ])


def _template_names(y: int) -> list[str]:
    return [
        f"{y} Fee Comparisons v2.xlsx",
        f"{y} Fee Comparisons.xlsx",
        f"{y} Fee Comparison v2.xlsx",
        f"{y} Fee Comparison.xlsx",
        f"{y} Fee Comparisons v2_updated.xlsx",
    ]


_SHARED_TEMPLATE_NAMES = (
    "Fee Comparisons Template.xlsx",
    "Fee Comparison Template.xlsx",
    "Template.xlsx",
)


def _plausible_template_files(d: Path, already: set[Path]) -> list[Path]:
    """Other .xlsx files in `d` that could be a combined workbook, as a last
    resort. Per-specialty extracts ("2025 DD Fee Comparisons.xlsx") are
    excluded by name -- each holds one specialty's sheets, never the full
    set -- so the search doesn't waste time opening them. Anything that
    survives is still sheet-validated before being accepted."""
    if not d.exists():
        return []
    per_specialty = re.compile(r"\b(GP|DH|SP|DD|QC)\b", re.IGNORECASE)
    out = []
    for p in sorted(d.glob("*.xlsx")):
        if p in already or p.name.startswith("~$"):
            continue
        stem = p.stem.lower()
        if "generated" in stem or "mismatch" in stem:
            continue
        if per_specialty.search(p.stem):
            continue
        out.append(p)
    return out


def template_candidates(year: int = None) -> list[Path]:
    """Every place the template might be, most specific first.

    The template is the SAME file for every year -- only its year labels are
    rewritten (see build_fee_comparison.build_year_label_map) -- so when the
    year being built has no combined workbook of its own, other years'
    output folders are searched too, newest first. That's the normal case
    for a year whose own folder holds only per-specialty extracts
    ("2025 DD Fee Comparisons.xlsx", "2025 GP Fee Comparisons.xlsx", ...)
    rather than a full combined workbook.
    """
    y = _y(year)
    d = output_dir(y)
    td = templates_dir()
    candidates: list[Path] = []

    # 0. A template kept deliberately in <Data>/templates wins over anything
    #    else, under ANY filename: putting a workbook there is the whole
    #    instruction. Conventional names first purely so a folder holding
    #    several files picks the obvious one; every candidate is still
    #    sheet-validated before it is accepted.
    candidates += [td / n for n in _SHARED_TEMPLATE_NAMES]
    candidates += [td / n for n in _template_names(y)]
    candidates += _plausible_template_files(td, set(candidates))

    # 1. This year's own combined workbook, by conventional name.
    candidates += [d / n for n in _template_names(y)]
    # 2. A year-less shared template, next to this year's output or at the
    #    Data root.
    candidates += [d / n for n in _SHARED_TEMPLATE_NAMES]
    candidates += [DATA_DIR / n for n in _SHARED_TEMPLATE_NAMES]
    # 3. Other years' combined workbooks, newest year first -- this is what
    #    lets a 2025 build borrow 2026's template.
    for other in available_years():
        if other == y:
            continue
        od = output_dir(other)
        candidates += [od / n for n in _template_names(other)]
        candidates += [od / n for n in _SHARED_TEMPLATE_NAMES]
    # 4. Last resort: any other plausible workbook -- this year's output
    #    folder, then the other years'. (The templates folder is already
    #    fully covered by step 0.)
    seen = set(candidates)
    extras = _plausible_template_files(d, seen)
    seen.update(extras)
    for other in available_years():
        if other == y:
            continue
        more = _plausible_template_files(output_dir(other), seen)
        seen.update(more)
        extras += more
    return candidates + extras


def template_workbook(year: int = None) -> Path:
    """The reference workbook whose header rows, styling, column layout and
    Claim Lines sheet the generated output is built from.

    Every candidate is checked for the sheets the build actually needs (see
    REQUIRED_TEMPLATE_SHEETS) rather than trusted on filename alone --
    picking a workbook that merely matched the name pattern is what produced
    "KeyError: 'Worksheet Claim Lines does not exist.'" much later in the
    run, with nothing pointing at the real cause.

    Set config.TEMPLATE_FILE (or the OHB_TEMPLATE environment variable) to
    bypass the search entirely."""
    y = _y(year)
    explicit = os.environ.get("OHB_TEMPLATE") or TEMPLATE_FILE
    if explicit:
        return Path(explicit)
    for candidate in template_candidates(y):
        if candidate.exists() and _has_required_sheets(candidate):
            return candidate
    # Nothing usable: return the conventional name so the caller reports a
    # clear "not found / unusable" message (see build_fee_comparison.main)
    # instead of failing deep inside the build.
    return output_dir(y) / f"{y} Fee Comparisons v2.xlsx"


def output_workbook(year: int = None) -> Path:
    y = _y(year)
    return output_dir(y) / f"{y} Fee Comparisons - generated.xlsx"


GROUND_TRUTH_FILE = None


def ground_truth_workbook(year: int = None) -> Path:
    """The hand-maintained reference workbook that compare_fee_files.py checks
    the generated output against.

    The reference workbook carries BOTH years of fee columns ("2025 CDCP Fee"
    and "2026 CDCP Fee", "2025 PT Fee" and "2026 PT Fee"), both populated --
    so the SAME workbook is the ground truth for either year, and comparing a
    given year just means reading that year's pair of columns (which is what
    fee_column_labels already selects). A year with no combined workbook of
    its own therefore borrows another year's, exactly as the template does.

    That's why this searches across years rather than only this year's output
    folder: 2025's folder holds per-specialty, province-sheet workbooks that
    aren't this format at all, while the 2026 workbook already contains
    2025's figures.

    Prefers a "_updated" copy -- that's the working ground truth, kept apart
    from the pristine template the build copies its layout from.

    Set config.GROUND_TRUTH_FILE (or the OHB_GROUND_TRUTH environment
    variable) to bypass the search entirely.
    """
    explicit = os.environ.get("OHB_GROUND_TRUTH") or GROUND_TRUTH_FILE
    if explicit:
        return Path(explicit)
    y = _y(year)

    def _gt_names(yy: int) -> list[str]:
        return [
            f"{yy} Fee Comparisons v2_updated.xlsx",
            f"{yy} Fee Comparisons_updated.xlsx",
        ] + _template_names(yy)

    candidates = [output_dir(y) / n for n in _gt_names(y)]
    for other in available_years():
        if other != y:
            candidates += [output_dir(other) / n for n in _gt_names(other)]
    # Last: the templates folder. Moving the combined workbook there (rather
    # than copying it) would otherwise leave compare_fee_files with no
    # reference at all -- the template and the ground truth are often the
    # same file, differing only in whether the fee columns are filled in.
    td = templates_dir()
    candidates += [td / n for n in _gt_names(y)]
    candidates += _plausible_template_files(td, set(candidates))
    # Validated like the template: a per-specialty extract that merely
    # matched the name pattern isn't a usable reference workbook.
    for candidate in candidates:
        if candidate.exists() and _has_required_sheets(candidate):
            return candidate
    return output_dir(y) / f"{y} Fee Comparisons v2_updated.xlsx"


def mismatch_report(year: int = None) -> Path:
    """Where compare_fee_files.py writes its mismatch/missing/extra report."""
    return output_dir(_y(year)) / "Fee_Comparison_Mismatches.xlsx"


def mismatch_master(year: int = None) -> Path:
    """The reviewed copy of the mismatch report carrying the green
    highlights that highlight_matching_mismatches.py propagates."""
    return output_dir(_y(year)) / "Fee_Comparison_Mismatches_master.xlsx"


def fee_column_labels(year: int = None) -> set[str]:
    """The fee-column grouping labels in a Fee Comparison workbook's header.

    With no year given this covers BOTH years the build populates, so a
    comparison checks every fee column in the workbook rather than half of
    them. Pass a year to get just that year's pair."""
    if year is not None:
        return {f"{year} CDCP Fee", f"{year} PT Fee"}
    labels = set()
    for y in years():
        labels |= {f"{y} CDCP Fee", f"{y} PT Fee"}
    return labels


# Filename fragments (case-insensitive, matched against the file's stem) that
# force a PT fee guide to be accepted as a general SP source even though its
# name identifies no sub-specialty and does not say "SP"/"Specialist".
#
# SP resolution otherwise refuses such a file: one that names neither a
# sub-specialty nor specialist scope, and labels no row with a specialty,
# cannot say who its fees are for, and lending it to a specialty it never
# names is how ON_DA_Fee_Guide_2026 came to supply Ontario's OM/OP/RA rows.
# That rule is right for a decoy file and wrong for a province whose real
# specialist schedule simply isn't named to the convention -- this list is
# how you say "this particular file IS a legitimate specialist source"
# without weakening the rule everywhere else.
#
# Example:
#     SP_EXTRA_GENERAL_SOURCES = ["NS_Fee_Guide_2026", "NB_2026_Schedule"]
SP_EXTRA_GENERAL_SOURCES: list[str] = []
