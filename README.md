# BOM SMT Comparator (Python)

Desktop app that checks an SMT feeder setup sheet against an engineering BOM. It runs locally in Python 3.12 with a PySide6 window. Files stay on your PC.

![BOM SMT Comparator overview](docs/app-overview.png)

![Discrepancy results](docs/app-results.png)

![Dark mode](docs/app-dark.png)

## What it does

The app reads reference designators from a BOM workbook and from an SMT setup sheet, then lists differences:

- **Missing in SMT** — in the BOM, not on the setup sheet
- **Missing in BOM** — on the setup sheet, not in the BOM (including DNP parts that were loaded)
- **MPN Mismatch** — same designator, different part number
- **Duplicate in BOM** / **Duplicate in SMT** — the same designator listed more than once

Optional checks:

- Learn extra prefixes from the BOM
- Verify manufacturer part numbers on the setup-sheet lines
- Exclude BOM rows marked DNP / not fitted
- Columnar setup sheet — only for machine dumps with a fixed designator column (tab / 2+ spaces / `|`)

You can search or filter the table (click a summary chip to filter by type) and export a discrepancy report to Excel. **Dark mode** is a shop-floor theme. **Ctrl+Enter** compares; **Ctrl+E** exports.

## Python vs JavaScript

There is a matching browser app: [BomCompare_Javascript](https://github.com/sran4/BomCompare_Javascript). Both use the same comparison rules (missing, MPN mismatch, duplicates, DNP). Pick one:

| | Python (this repo) | JavaScript |
| --- | --- | --- |
| **Best when** | Daily use on a line PC you control | Sharing a file or a quick look with no install |
| **How you run it** | `python bom_smt_comparator.py` after a one-time venv | Double-click `BOM_SMT_Comparator.html` in Chrome or Edge |
| **Install** | Python 3.12 + `pip install -r requirements.txt` | None |
| **Sample on start** | No — you choose files | Yes — built-in demo with a “Sample data” banner |
| **PDF / Excel** | pdfplumber (and pypdf fallback), pandas/openpyxl | PDF.js, SheetJS, ExcelJS in the page |
| **Settings** | Remembered in Windows (window size, prefixes, dark mode) | Remembered in the browser (localStorage) |

**Python — pros**

- Stays up as a normal Windows window (drop files, shortcuts, export).
- Header auto-detect and PDF text extract were solid on the shared test BOMs.
- No wait for large in-page spreadsheet libraries.

**Python — cons**

- Must install Python and packages on each PC.
- No demo comparison until you pick files.
- Not something you email as a single double-click file.

**JavaScript — pros**

- Zero install; one HTML file or the `web/` folder.
- Opens with a sample so new users see chips, filters, and a real table immediately.
- Easy to put on a share or GitHub for anyone with a browser.

**JavaScript — cons**

- The single-file HTML is very large (libraries are inlined).
- Relies on the browser; a blocked or slow script load can delay the first compare.
- Weaker fit as the always-on tool on an SMT station.

Use **this Python app** on the machine that compares production BOMs. Use **JavaScript** when someone only needs to try the tool or cannot install Python.

## What you need

- Windows (tested with Python 3.12)
- Python 3.12

Supported inputs:

- **SMT setup sheet:** `.txt`, `.log`, or a text-based `.pdf` (not a scanned image)
- **BOM:** `.xlsx`, `.xls`, or `.xlsm`

## How to run

If you **downloaded a ZIP on Windows**, follow **[INSTALL_WINDOWS.md](INSTALL_WINDOWS.md)** or double-click **`Run.bat`**.

From a terminal in this folder:

```powershell
python -m venv venv
.\venv\Scripts\python.exe -m pip install -r requirements.txt
.\venv\Scripts\python.exe bom_smt_comparator.py
```

If the virtual environment already exists, skip the first two lines and run:

```powershell
.\venv\Scripts\python.exe bom_smt_comparator.py
```

## How to compare files

1. Browse or drop the **SMT setup sheet** (left panel, first field).
2. Browse or drop the **BOM Excel** file. Sheet, header row, designator, MPN, and description columns are filled in automatically when possible.
3. Header row **Auto-detect** is usually enough. Change **Sheet** or columns if your workbook uses unusual names.
4. Adjust **SMT prefixes** if your designators are not the usual `R`, `C`, `U`, and so on.
5. Click **Compare Files** (or press Ctrl+Enter).
6. Filter or search the discrepancy table, or click a colored chip. Click **Export to Excel** (or press Ctrl+E) to save a report.

## Notes

- Password-protected workbooks and PDFs are not supported. Save an unprotected copy first.
- Image-only PDFs need OCR first, or export the setup sheet as text from the SMT software.
- Comparison runs on this machine. Your BOM and setup files are not uploaded anywhere.
