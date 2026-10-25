# Install on Windows (Python app)

Use this after you download the ZIP from GitHub or clone the repo. Files stay on this PC.

## 1. Install Python 3.12 (once per PC)

1. Open https://www.python.org/downloads/
2. Download **Python 3.12**.
3. Run the installer.
4. Check **Add python.exe to PATH**.
5. Click **Install Now**.

To check: open Command Prompt and run `py -3.12 --version`. You should see `Python 3.12.x`.

## 2. Unpack the app

1. If you downloaded a ZIP from GitHub, right-click it → **Extract All…**
2. Open the folder that contains `bom_smt_comparator.py` and `Run.bat`.

## 3. Run it

**Easiest:** double-click `Run.bat`.

The first time it creates a `venv` folder and installs packages (a few minutes). After that it only starts the window.

If Windows SmartScreen blocks the `.bat`: **More info** → **Run anyway**.

**Or** in Command Prompt, in this folder:

```powershell
py -3.12 -m venv venv
.\venv\Scripts\python.exe -m pip install -r requirements.txt
.\venv\Scripts\python.exe bom_smt_comparator.py
```

## 4. Optional desktop shortcut

1. Right-click `Run.bat` → **Show more options** → **Send to** → **Desktop (create shortcut)**.
2. Rename the shortcut if you want.

## After it opens

1. Browse or drop the SMT setup sheet (`.txt`, `.log`, or text PDF).
2. Browse or drop the BOM Excel file.
3. Click **Compare Files**.

Password-protected Excel/PDF files are not supported. Scanned (image-only) PDFs need OCR first.

## If it fails

- `py -3.12` not found → Python 3.12 is not installed, or PATH was not checked.
- `pip` errors about SSL → try another network, or run `Run.bat` again.
- Window does not appear → run `Run.bat` from Command Prompt so you can read the error.
