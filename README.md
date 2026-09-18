# Fee Comparison Builder

Builds one Excel workbook that puts **two rate years side by side** — for every
province and procedure code, what the CDCP pays and what the provincial fee
guide charges.

```
Sheet: DH · GP · QC GP · SP · QC SP · DD
Row:   province + procedure code
Cols:  2025 CDCP Fee | 2026 CDCP Fee | 2025 PT Fee | 2026 PT Fee
```

It reads the CDCP price files and the provincial fee guides (Excel, CSV, Word
or PDF) and writes the finished workbook. Nothing is typed in by hand.

---

# Part 1 — How to run it

**This is the whole job. Parts 2 and 3 are background only.**

## Step 1 — Install (once per computer)

Python 3.11 or newer is required.

```bash
pip install -r requirements.txt
```

## Step 2 — Put the files in the right folders

Everything lives under a folder called `Data`, next to `scripts`. One folder
per year:

```
Data/
├── templates/
│   └── <your template workbook>.xlsx      ← the blank workbook to copy the layout from
│
├── 2025/
│   ├── 2025_CDCP Fees/                    ← CDCP price files
│   └── 2025_Fee Guides/
│       ├── DH/                            ← dental hygienist guides
│       ├── GP/                            ← general practitioner guides
│       ├── SP/                            ← specialist guides
│       └── DD/                            ← denturist guides
│
└── 2026/
    ├── 2026_CDCP Fees/
    ├── 2026_Fee Guides/
    │   ├── DH/   GP/   SP/   DD/
    └── 2026_Output/                       ← the finished workbook appears here
```

`2026_Output` is created automatically. You never create it yourself.

### Naming the files

This is the one thing that has to be right.

**CDCP price files** — one per province, named with the year and the province:

```
2026 - CDCP PRICE FILE - ON.xlsx
2026 - CDCP PRICE FILE - BC.xlsx
```

Each file has a sheet per specialty (`DH`, `GP`, `SP`, `DD`) with columns
`Province`, `Specialty`, `Procedure Code`, `Provider Fee` (and
`Internal Lab Fee` on the `DD` sheet).

**Fee guides** — the filename must **start with the province code**:

```
Data/2026/2026_Fee Guides/GP/ON_GP_Fee_Guide_2026.pdf
Data/2026/2026_Fee Guides/DH/BC_DH_Fee_Guide_2026.xlsx
```

Anything after the province code is free text. Excel, CSV, Word and PDF all
work. (A folder named for the province works too — `GP/ON/anything.pdf`.)

**Specialist guides** — if a province has one file per specialty, put the
specialty in the name so the right codes come from the right file:

```
SP/ON_EN_Fee_Guide_2026.xlsx     ← Endodontics
SP/ON_PE_Fee_Guide_2026.xlsx     ← Periodontics
```

Recognised specialties: `EN` Endodontics · `OS` Oral Surgery ·
`PA` Paediatric · `PE` Periodontics · `PR` Prosthodontics · `OM` Oral Medicine ·
`OP` Oral Pathology · `OR` Orthodontics · `RA` Radiology · `AN` Anaesthesia.
Full words work as well as the two-letter codes ("Periodontics" is read the
same as "PE"). One file covering every specialty is fine too.

## Step 3 — Set the two years

Open `scripts/config.py`. Near the top:

```python
PAST_YEAR = 2025
CURRENT_YEAR = 2026
```

**These two numbers are the only thing you change from year to year.** For the
2026/2027 build:

```python
PAST_YEAR = 2026
CURRENT_YEAR = 2027
```

Every folder path, filename, column heading and the output filename follow
from them automatically.

You can also do it for one run without editing the file:

```bash
python scripts/build_fee_comparison.py --years 2026 2027
```

## Step 4 — Run it

```bash
python scripts/build_fee_comparison.py
```

It takes a few minutes. It prints what it is reading as it goes.

## Step 5 — Check the header before you walk away

The first lines tell you whether it found everything:

```
=== Building Fee Comparison: 2026 and 2027 ===
  2026 CDCP price files : Data/2026/2026_CDCP Fees
  2026 PT fee guides    : Data/2026/2026_Fee Guides
  2027 CDCP price files : Data/2027/2027_CDCP Fees
  2027 PT fee guides    : Data/2027/2027_Fee Guides
  Template              : Data/templates/Fee Comparisons Template.xlsx   [Data/templates]
  Output                : Data/2027/2027_Output/2027 Fee Comparisons - generated.xlsx
```

You may also see a `Header years` line such as `2025 -> 2026, 2026 -> 2027`.
That is normal: the template's own column headings are being relabelled to the
years you asked for. Nothing to do.

If a line ends with **`<- NOT FOUND`**, that folder is missing or misspelled.
Stop and fix the folder name — the build will still run, but that year's
columns will all come out `N/A`.

Below the header, one line per province shows which file supplied its fees:

```
  ON GP: ON_GP_Fee_Guide_2027.pdf (1204)
  BC GP: no PT fee guide found
```

`no PT fee guide found` means no file in that folder starts with `BC`.

## Step 6 — Open the result

```
Data/<current year>/<current year>_Output/<current year> Fee Comparisons - generated.xlsx
```

It is named for the newer year and contains **both** years.

---

## What `N/A` means

`N/A` is a real answer, not an error. It means **that source does not price
that code that year**.

Fees are never borrowed — not from the other year, not from another province,
not from a general guide to fill a specialist gap. Every number in the
workbook comes from that year's file for that province and that specialty. A
gap is shown as a gap on purpose, so nothing invented reaches the output.

Rows are the union of both years, so a code the CDCP only started listing in
the newer year still appears, with the older year shown for comparison.

## Known gaps

These are deliberate, and unchanged from previous builds:

- **Claim counts are written as `0`.** The CDCP files don't carry claim
  counts. The weighting formulas are written and ready, so dropping real
  counts in later makes the workbook calculate correctly with no code change.
- **QC GP / QC SP fee columns are plain numbers, not formulas.** The template
  had formulas pointing at linked workbooks we don't have; those are replaced
  with the extracted values.
- **GP's last column** was already a broken `#REF!` in the template. It is
  replaced with a working equivalent.

## If something goes wrong

| What you see | What to do |
|---|---|
| `no template workbook found` | Put the template workbook in `Data/templates/`. Any filename works, as long as it has the sheets `Claim Lines, DH, GP, QC GP, SP, QC SP, DD`. |
| `<- NOT FOUND` on a folder line | Check that year's folder name matches the layout above exactly. |
| `no PT fee guide found` for a province | Rename the guide so it **starts** with the province code (`ON_...`, `BC_...`). |
| A whole province is `N/A` | Same cause — the filename isn't starting with the province code. |
| `could not find column(s) ...` | That CDCP price file spells a column differently. The message names the file and lists the columns it did find. |
| `Permission denied` when saving | The output workbook is open in Excel. Close it and re-run. |

---

# Part 2 — Testing scripts

**You do not need these. Skip this section.**

These were used while the extraction was being built and checked. They compare
the generated workbook against a **hand-maintained "ground truth" workbook** —
a copy filled in by hand so the automated output could be verified row by row
against it.

**That ground-truth workbook is not part of this handover, and there is no
reason for you to create one.** The extraction has already been checked
against it; that work is finished. These scripts are kept only so the checking
can be repeated if the extraction logic is ever changed.

| Script | What it was for |
|---|---|
| `compare_fee_files.py` | Compares generated output against the ground-truth workbook and writes a mismatch report. |
| `highlight_matching_mismatches.py` | Carries "reviewed / accepted" green highlights from one mismatch report to the next. |
| `trace_fee_source.py` | For one code, shows which guide file supplied its fee and what every other file would have said. Useful for investigating a single surprising number. |
| `DH Full PDF Data Extraction (except qc) - demo.py` | An early standalone demo of PDF extraction. Superseded. |

Running any of these is optional and affects nothing. They never modify the
generated workbook.

---

# Part 3 — What each file does

You should not need to edit any of these. Listed so a future developer knows
where to look.

| File | Role |
|---|---|
| `config.py` | **The two year numbers**, and every folder path derived from them. The only file you normally edit. |
| `build_fee_comparison.py` | The script you run. Merges the two years and writes the workbook. |
| `fee_dataset.py` | Loads one year's fees for one sheet. Keeps "which year" a setting rather than a rewrite. |
| `fee_extraction.py` | Reads fees out of guides in any format — xlsx, csv, docx, pdf. |
| `cdcp_loader.py` | Reads the CDCP price files. |
| `sheet_builders.py` | Writes each sheet with the template's columns, styling and formulas. |
| `convert_ic_na_values.py` | One-off utility: rewrites `I.C.` and `#N/A` cells in a workbook to `N/A`. Run by hand if needed. |

### Two settings you will probably never touch

Both are in `config.py`, with fuller notes beside them:

- **`TEMPLATE_FILE`** — set this to a path to bypass the template search
  entirely. Normally unnecessary: a workbook in `Data/templates/` is found
  automatically under any filename.
- **`SP_EXTRA_GENERAL_SOURCES`** — a list of filename fragments. A specialist
  guide is only used for a specialty it names. If a province's genuine
  specialist schedule isn't named to that convention, list it here to allow
  it. Leave empty unless a specialist province comes out entirely `N/A`
  despite having a guide.
