"""
Last-resort OCR fee extractor, for PDF fee guides that have no extractable
text layer at all -- e.g. NL_DD_Fee_Guide_2026.pdf, whose text was flattened
to vector curves on export (CorelDRAW's "convert text to curves"), leaving
zero Tj/TJ text-showing operators and no /Font resource anywhere in the
file. load_fees_from_pdf (fee_extraction.py) can't help here no matter how
it reads the content stream, because there is no text to read -- only OCR,
which reads the *rendered page image* rather than any text encoding, can
recover anything.

load_pt_fees_from_files (fee_extraction.py) imports load_fees_from_pdf_via_ocr
from here and calls it automatically, but only as the very last fallback
tier, and only for whichever codes are still missing after every other
source -- spreadsheet, csv, docx, and normal (non-OCR) pdf text -- has
already been tried. That gating matters: OCR is slow (seconds per page) and
occasionally misreads digits (a "0" read as an "8", a column read out of
order), so it should only ever run for the rare file that genuinely has no
other way to resolve its fees, not as a first choice. The import is wrapped
in a try/except there, so a machine without the OCR dependencies installed
(see below) simply skips this tier with a warning instead of failing the
whole extraction run.

Requires, in addition to this project's usual dependencies:
  - the `tesseract-ocr` and `poppler-utils` system packages
  - the `pytesseract`, `pdf2image`, and `Pillow` pip packages

Can also be run directly for manual, one-off inspection of a single PDF (the
result should be spot-checked against the source file by eye before being
trusted -- this prints its findings for a person to review, not for another
script to consume):

    python scripts/ocr_pdf_fees.py <path_to_pdf> [known_codes_file]

    <path_to_pdf>       The PDF to OCR.
    [known_codes_file]  Optional text file, one procedure code per line. If
                         omitted, every bare 5-digit number the OCR pass
                         finds is treated as a candidate code -- fine for a
                         first look at a new file, but noisier than passing
                         the real CDCP code list for that specialty, since
                         some OCR misreads (or unrelated 5-digit numbers in
                         the prose) will otherwise be picked up too.
"""

import sys
from pathlib import Path

from pdf2image import convert_from_path
import pytesseract

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fee_extraction import extract_codes_from_text, _ANY_CODE_RE  # noqa: E402

DEFAULT_DPI = 200


def ocr_pdf_text(path: Path, dpi: int = DEFAULT_DPI) -> str:
    """Rasterize every page of `path` and OCR it, returning the pages'
    recognized text joined together. Tesseract does its own layout analysis
    on the rendered image (unlike pypdf's draw-order text extraction), so a
    table's rows come back in the right left-to-right, top-to-bottom order
    without needing pypdf's "layout" extraction-mode workaround."""
    pages = []
    images = convert_from_path(str(path), dpi=dpi)
    for page_num, image in enumerate(images, start=1):
        text = pytesseract.image_to_string(image)
        pages.append(text)
        print(f"  OCR'd page {page_num}/{len(images)} ({len(text)} chars)", file=sys.stderr)
    return "\n".join(pages)


def load_fees_from_pdf_via_ocr(path: Path, known_codes: set[str], dpi: int = DEFAULT_DPI) -> dict[str, float]:
    """OCR-based counterpart to fee_extraction.load_fees_from_pdf. Reuses the
    same (code, fee) text scanner (extract_codes_from_text) so OCR'd text is
    interpreted with exactly the same fee-token rules (currency formats,
    "I.C."/"c.s." no-fee markers, etc.) as every other PDF source, rather
    than a second, drifting copy of that logic."""
    text = ocr_pdf_text(path, dpi=dpi)
    return extract_codes_from_text(text, known_codes)


def main():
    if len(sys.argv) not in (2, 3):
        print("Usage: python ocr_pdf_fees.py <path_to_pdf> [known_codes_file]")
        sys.exit(1)

    pdf_path = Path(sys.argv[1])

    if len(sys.argv) == 3:
        codes_file = Path(sys.argv[2])
        known_codes = {
            line.strip() for line in codes_file.read_text().splitlines() if line.strip()
        }
        print(f"Loaded {len(known_codes)} known code(s) from {codes_file}", file=sys.stderr)
    else:
        print("No known_codes_file given -- OCR'ing once first to discover candidate codes.",
              file=sys.stderr)
        discovery_text = ocr_pdf_text(pdf_path)
        known_codes = set(_ANY_CODE_RE.findall(discovery_text))
        print(f"Found {len(known_codes)} candidate 5-digit code(s) in the OCR'd text.",
              file=sys.stderr)

    fees = load_fees_from_pdf_via_ocr(pdf_path, known_codes)

    for code in sorted(fees):
        print(f"{code}: {fees[code]}")
    print(f"\nResolved {len(fees)} / {len(known_codes)} known codes.", file=sys.stderr)


if __name__ == "__main__":
    main()