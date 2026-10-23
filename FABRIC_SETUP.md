# Microsoft Fabric setup — no source-file changes required

This guide explains how to transfer and run the existing repository in Microsoft Fabric. **Do not import every `.py` file as a Notebook.** Keep the Python modules together as files, and create one Fabric notebook to launch the existing build. No edits to the repository's scripts or `requirements.txt` are needed.

## 1. Which file goes where?

Recommended code destination: **Lakehouse → Files → code → OHB-fee-extraction**. Preserve the `scripts/` folder rather than flattening its contents.

| Repository file | How to handle it in Fabric | Purpose / when to use |
| --- | --- | --- |
| `README.md` | Keep with the code, or read on GitHub. Do not import as a Notebook. | Input folder layout, file naming rules, and interpretation of results. |
| `FABRIC_SETUP.md` | Keep with the code, or read on GitHub. Do not import as a Notebook. | This setup guide. |
| `requirements.txt` | Upload to `Files/code/OHB-fee-extraction/requirements.txt`. Use it to install dependencies as described below. | Dependency specification; not executable code or a custom Python library. |
| `.gitignore` | Keep if transferring a full checkout; not needed to run the build. | Git ignore rules; not a notebook or dependency file. |
| `scripts/build_fee_comparison.py` | Upload unchanged to `Files/code/OHB-fee-extraction/scripts/`. Launch it from the notebook in section 5. | Main entry point: builds the two-year comparison workbook. |
| `scripts/config.py` | Upload unchanged into the same `scripts/` directory. Do not run separately. | Year defaults, data paths, and template/output configuration. Notebook environment variables and CLI arguments supply run-specific settings. |
| `scripts/cdcp_loader.py` | Upload unchanged into the same `scripts/` directory. Do not run separately. | Reads CDCP price files; imported by the build. |
| `scripts/fee_dataset.py` | Upload unchanged into the same `scripts/` directory. Do not run separately. | Loads and combines each year's fee datasets. |
| `scripts/fee_extraction.py` | Upload unchanged into the same `scripts/` directory. Do not run separately. | Extracts fee-guide data from supported file formats. |
| `scripts/sheet_builders.py` | Upload unchanged into the same `scripts/` directory. Do not run separately. | Creates workbook sheets and formatting. |
| `scripts/compare_fee_files.py` | Keep unchanged in `scripts/`. Optional; not part of the normal build. | Compares output with a separately supplied ground-truth workbook. See the original README before using. |
| `scripts/highlight_matching_mismatches.py` | Keep unchanged in `scripts/`. Optional; not part of the normal build. | Carries reviewed highlights between mismatch reports. |
| `scripts/trace_fee_source.py` | Keep unchanged in `scripts/`. Optional diagnostic tool. | Investigates which source supplied a fee. |
| `scripts/convert_ic_na_values.py` | Keep unchanged in `scripts/`. Optional, manual utility. | Rewrites certain workbook values. Do not run automatically as part of extraction. |

The repository currently contains scripts rather than a packaged Python project. Installing this repository with `pip install git+...` is not the setup method described here. The separate `file-processing` repository in the background guidance is not a replacement for these scripts.

## 2. Create the Fabric items and upload files

1. In your Fabric workspace, create or select a **Lakehouse**.
2. Create a **Notebook** using a Spark runtime with Python/PySpark cells and Python **3.11 or newer**. The extraction itself uses ordinary driver-side Python, not distributed Spark jobs.
3. Add the Lakehouse to the notebook and set it as the **default Lakehouse**. This guide depends on its mounted path, `/lakehouse/default/Files`. A standalone Python notebook without that mount is not the workflow shown here.
4. Under Lakehouse **Files**, create `code/OHB-fee-extraction/`. Upload `requirements.txt` and the complete `scripts/` directory, preserving the directory structure. If the upload interface flattens folders, create the folders explicitly and upload their contents into them.
5. Upload your own input data and template under **Files/Data/**. Those workbooks and fee guides are not included in this repository.

Expected layout, using 2025 and 2026 as examples:

```text
Files/
├── code/
│   └── OHB-fee-extraction/
│       ├── requirements.txt
│       └── scripts/
│           ├── build_fee_comparison.py
│           ├── config.py
│           ├── cdcp_loader.py
│           ├── fee_dataset.py
│           ├── fee_extraction.py
│           ├── sheet_builders.py
│           └── ...remaining scripts...
└── Data/
    ├── templates/
    │   └── <your template>.xlsx
    ├── 2025/
    │   ├── 2025_CDCP Fees/
    │   └── 2025_Fee Guides/
    │       ├── DH/
    │       ├── GP/
    │       ├── SP/
    │       └── DD/
    └── 2026/
        ├── 2026_CDCP Fees/
        └── 2026_Fee Guides/
            ├── DH/
            ├── GP/
            ├── SP/
            └── DD/
```

Use the exact input naming conventions in `README.md`. The template must contain these sheets: **Claim Lines, DH, GP, QC GP, SP, QC SP, DD**.

## 3. Get a Python environment working with requirements.txt

A **Fabric Environment** defines the runtime and shared libraries attached to your notebook. It is not a laptop-style virtual environment that you activate with `venv`. Choose **A** for repeatable shared/scheduled use, or **B** for a quick interactive setup; you do not need both.

The current `requirements.txt` specifies:

```text
openpyxl>=3.1.5
pypdf>=6.14.2
xlrd>=2.0.2
python-docx>=1.2.0
```

Use the repository file as the source of truth if these requirements change. These are minimum versions, not a lockfile; different installation dates may resolve different compatible versions.

### A. Shared Fabric Environment — recommended

1. Create an **Environment** item in the same workspace.
2. Select a supported Spark runtime with Python 3.11 or newer.
3. Open its **Libraries / Public libraries** configuration. Add the four packages from `requirements.txt`, choosing available versions that satisfy each `>=` constraint. If your Fabric UI provides an import-from-`requirements.txt` option, use it and review the resulting package selections; otherwise enter the packages manually. Do not upload `requirements.txt` as a wheel/custom library or assume that putting it in Environment Resources installs anything.
4. **Save and publish** the Environment. Wait for publishing to finish and resolve any dependency errors.
5. In the notebook's Environment selector, attach that Environment. Stop/restart the notebook session if necessary so it uses the published configuration.
6. Run the verification cell below before running the build.

Publishing may select concrete versions rather than retaining open-ended constraints. Record the successfully tested versions for repeatable scheduled execution. If package resolution fails, check runtime compatibility and available versions rather than silently dropping dependencies or lowering the repository's minimum versions.

### B. Install requirements into an interactive notebook session

After uploading `requirements.txt`, run this in its own Python/PySpark notebook cell:

```python
%pip install -r /lakehouse/default/Files/code/OHB-fee-extraction/requirements.txt
```

Use `%pip` in the notebook, not `!pip`, so installation targets the notebook's Python environment. This is session-scoped setup, not a replacement for publishing a shared Environment. If Fabric requests or performs a Python/session restart, rerun your configuration cells afterward. A new session may require installing again. Network/package access must be allowed by your workspace's policies.

### Verify the active interpreter and installed libraries

Run in a Python notebook cell after either setup option:

```python
import sys
from importlib.metadata import version

print('Python:', sys.version)
print('Executable:', sys.executable)
assert sys.version_info >= (3, 11), 'Select a runtime with Python 3.11 or newer.'

import openpyxl
import pypdf
import xlrd
import docx  # installed distribution is named python-docx

for package in ('openpyxl', 'pypdf', 'xlrd', 'python-docx'):
    print(f'{package}: {version(package)}')
```

Compare the printed versions with `requirements.txt`. Successful imports alone do not prove that minimum versions are met.

## 4. Notebook configuration cell

Copy the following into a Python cell in your Fabric notebook. These are **notebook settings**, not edits to `scripts/config.py`.

```python
from pathlib import Path

CODE_ROOT = Path('/lakehouse/default/Files/code/OHB-fee-extraction')
DATA_DIR = Path('/lakehouse/default/Files/Data')
PAST_YEAR = 2025
CURRENT_YEAR = 2026
# Leave None to use the existing template search, starting in Data/templates/.
TEMPLATE = None
# Example explicit override:
# TEMPLATE = Path('/lakehouse/default/Files/Data/templates/My Template.xlsx')

assert PAST_YEAR < CURRENT_YEAR, 'Choose two distinct years, oldest first.'
assert DATA_DIR.is_dir(), f'Upload Data/ and attach the default Lakehouse: {DATA_DIR}'
for name in ('build_fee_comparison.py', 'config.py', 'cdcp_loader.py',
             'fee_dataset.py', 'fee_extraction.py', 'sheet_builders.py'):
    assert (CODE_ROOT / 'scripts' / name).is_file(), f'Missing script: {name}'
if TEMPLATE is not None:
    assert Path(TEMPLATE).is_file(), f'Missing template: {TEMPLATE}'
```

## 5. Run the existing build without modifying it

Copy this into the next Python notebook cell:

```python
import os
import subprocess
import sys

# Use a copied environment so these overrides do not change other notebook work.
env = os.environ.copy()
env['OHB_DATA_DIR'] = str(DATA_DIR)
env.pop('OHB_TEMPLATE', None)
if TEMPLATE is not None:
    env['OHB_TEMPLATE'] = str(TEMPLATE)

command = [
    sys.executable, '-u',
    str(CODE_ROOT / 'scripts' / 'build_fee_comparison.py'),
    '--years', str(PAST_YEAR), str(CURRENT_YEAR),
]
subprocess.run(command, env=env, check=True)
```

This launches the original script with the notebook's Python interpreter, explicit years, and Lakehouse paths. A fresh process avoids notebook command-line arguments and cached configuration from earlier runs. Logs stream into the notebook. `check=True` raises an error on a nonzero exit, so a failed build is not silently treated as successful.

Review the startup banner and per-province logs. **Missing year/source folders are warnings, not necessarily build failures**: the existing code can produce N/A values for absent sources. A successful process does not prove that all expected fee guides were uploaded.

For the standard layout, the output is:

```text
Files/Data/<current year>/<current year>_Output/<current year> Fee Comparisons - generated.xlsx
```

The exact `Saved` path in the logs is authoritative because the scripts also recognize legacy folder layouts. Open Lakehouse Files to locate/download the workbook. This workflow writes Excel files, not Delta tables.

## 6. Built-in Resources, Environment Resources, or Lakehouse?

| Approach | When to choose it | How to use these files |
| --- | --- | --- |
| Notebook Built-in Resources | A notebook-specific copy of small supporting code. | Upload `scripts/` and `requirements.txt`, preserving the tree. Change `CODE_ROOT` to the actual local resource path provided by Fabric. Each notebook needs its own copy. |
| Environment Resources | Shared code for notebooks attached to the same Environment. | Upload the same tree and set `CODE_ROOT` to its actual accessible local resource path. Sharing source files does not install Python packages; configure Libraries separately. |
| Lakehouse Files | Persistent code/data, or a Git checkout when supported by the runtime/filesystem. | Use the paths in this guide. Manual upload works without Git; a cloned checkout can retain `.git`. |

Keep input/output data in the Lakehouse whichever code-storage option you select. Resource paths vary by runtime and notebook type: obtain the actual path from Fabric's resource explorer rather than guessing a mount location. For session installation from a different code location, change the path in the `%pip` cell too.

Uploading files does not create a Git working copy. Environment Resources are not equivalent to a Git checkout. If cloning into Lakehouse Files, first verify Git availability and filesystem support; use your organization's approved authentication method for private repositories. Do not embed tokens in notebooks, clone URLs, or committed files.

Separate branches alone do not isolate users sharing a single checkout. Use separate working directories per user/run and coordinate Git changes. Likewise, builds for the same current year can overwrite the same output workbook: serialize execution or give each run a separate data root with the required inputs/template.

## 7. Troubleshooting

| Symptom | What to check |
| --- | --- |
| `/lakehouse/default/Files` does not exist | Attach and select a default Lakehouse; use a notebook/runtime supporting that mount; restart the session if needed. |
| `ModuleNotFoundError: openpyxl`, `pypdf`, `xlrd`, or `docx` | Publish/attach the dependency Environment and restart, or run the `%pip` requirements cell in this session. `docx` comes from `python-docx`. |
| `ModuleNotFoundError: config` or another repository module | Preserve the complete `scripts/` folder and launch the entry point shown above. Installing requirements does not install repository modules. |
| No matching distribution / dependency conflict | Check the package version constraints, runtime compatibility, package network access, and Environment publication errors. |
| Template missing or unusable | Upload a real template with all seven required sheets; optionally set `TEMPLATE` in the notebook. |
| Missing-source warnings or unexpected N/A values | Check exact year folders, specialty folders, and province-prefixed filenames against the original README. |
| Permission or save error | Check Lakehouse access, destination permissions, and concurrent builds writing the same workbook. |

The optional comparison/diagnostic utilities are not needed for normal extraction. Read their usage and supply their own required inputs before launching them; the hand-maintained ground-truth workbook is not supplied.

## 8. Reference documentation

Fabric interfaces can change; consult Microsoft documentation for runtime-specific Environment and notebook behavior:

- [Fabric environments overview](https://learn.microsoft.com/en-us/fabric/data-engineering/create-and-use-environment)
- [Manage Python libraries in Fabric notebooks](https://learn.microsoft.com/en-us/fabric/data-engineering/library-management)
- [Use a notebook in Fabric](https://learn.microsoft.com/en-us/fabric/data-engineering/how-to-use-notebook)

This is a documentation-only setup guide. The steps require a Fabric workspace and your actual data/template for end-to-end validation; no Fabric execution is implied by adding this file.
