"""
Show which PT fee guide file supplies a given code's fee, and what every other
file for that province would have said.

When a code comes out with an unexpected fee, the question is almost always
"which file did that number come from?" -- the build only reports the winning
source per province, and several files can cover the same code (spreadsheets
are read before csv, csv before docx, docx before pdf, and the first source to
supply a code wins). This walks the same discovery and ordering the build uses
but reports every file's answer separately, so a wrong value can be traced to
the file that produced it instead of guessed at.

Read-only: it loads guides and prints. It never writes anything.

Usage:
    python scripts/trace_fee_source.py <PT> <specialty> <code> [sub_specialty]

    python scripts/trace_fee_source.py QC GP 80672
    python scripts/trace_fee_source.py NS SP 81253 PA
    python scripts/trace_fee_source.py NB SP 56233 PR

<specialty> is the fee-guide folder: GP, SP, DH or DD.
<sub_specialty> only applies to SP (EN, PA, PE, PR, OS, OM, OP, OR, RA, AN)
and makes the lookup specialty-aware, exactly as the SP build is.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config
from fee_extraction import (
    CSV_SUFFIXES, DOC_SUFFIXES, PDF_SUFFIXES, SPREADSHEET_SUFFIXES,
    classify_file_specialty, discover_pt_files, load_pt_fees_from_files,
    normalize_code, sp_source_decisions,
)


def _tier(path: Path) -> str:
    """The order load_pt_fees_from_files reads formats in -- the first file
    that supplies a code wins, so this is what decides which value you get."""
    suffix = path.suffix.lower()
    for order, (label, suffixes) in enumerate(
        (("1 spreadsheet", SPREADSHEET_SUFFIXES), ("2 csv", CSV_SUFFIXES),
         ("3 docx", DOC_SUFFIXES), ("4 pdf", PDF_SUFFIXES)), start=1
    ):
        if suffix in suffixes:
            return label
    return "- other"


def main() -> None:
    if len(sys.argv) not in (4, 5):
        raise SystemExit(
            "Usage: python trace_fee_source.py <PT> <specialty> <code> [sub_specialty]\n"
            "  e.g. python trace_fee_source.py QC GP 80672\n"
            "       python trace_fee_source.py NS SP 81253 PA"
        )
    province, specialty, raw_code = sys.argv[1].upper(), sys.argv[2].upper(), sys.argv[3]
    sub_specialty = sys.argv[4].upper() if len(sys.argv) == 5 else None

    code = normalize_code(raw_code)
    if code is None:
        raise SystemExit(f"ERROR: {raw_code!r} is not a recognizable procedure code.")

    specialty_dir = config.pt_guides_dir() / specialty
    print(f"Tracing {config.YEAR} {province} {specialty} code {code}"
          + (f" (sub-specialty {sub_specialty})" if sub_specialty else ""))
    print(f"  Fee guide folder: {specialty_dir}")
    if not specialty_dir.exists():
        raise SystemExit(f"ERROR: folder does not exist: {specialty_dir}")

    files = discover_pt_files(specialty_dir, province)
    if not files:
        raise SystemExit(
            f"\nNo file in that folder is recognized as {province}'s. A guide is found\n"
            f"when its name STARTS with the province code (or an alias), or when it sits\n"
            f"in a per-province subfolder. Check the filename."
        )

    print(f"\n{len(files)} file(s) discovered for {province}, in the order they are read:\n")
    header = f"  {'read order':<14}{'file':<52}{'names specialty':<17}value for this code"
    print(header)
    print("  " + "-" * (len(header) - 2))

    for path in sorted(files, key=lambda f: (_tier(f), f.name)):
        named = ",".join(sorted(classify_file_specialty(path, province))) or "-"
        try:
            fees, _ = load_pt_fees_from_files(
                [path], {code}, verbose=False, target_specialty=sub_specialty
            )
            value = fees.get(code)
            shown = "(nothing)" if value is None else repr(value)
        except Exception as exc:  # a broken file shouldn't hide the others
            shown = f"ERROR: {exc}"
        print(f"  {_tier(path):<14}{path.name[:50]:<52}{named:<17}{shown}")

    if sub_specialty:
        skipped = sp_source_decisions(specialty_dir, province)[1]
        if skipped:
            print("\n  Excluded from SP resolution entirely:")
            for path, reason in skipped:
                print(f"    {path.name}: {reason}.")

    print("\n  The build takes the first non-empty value reading top to bottom.")
    print("  A file showing '(nothing)' didn't have the code, or had no readable fee for it.")


if __name__ == "__main__":
    main()
