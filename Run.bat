@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>&1
if errorlevel 1 (
  echo Python launcher "py" was not found.
  echo Install Python 3.12 from https://www.python.org/downloads/
  echo and check "Add python.exe to PATH".
  pause
  exit /b 1
)

if not exist "venv\Scripts\python.exe" (
  echo Creating Python 3.12 virtual environment...
  py -3.12 -m venv venv
  if errorlevel 1 (
    echo Could not create venv with Python 3.12. Trying "py -3"...
    py -3 -m venv venv
  )
  if errorlevel 1 (
    echo Failed to create venv. See INSTALL_WINDOWS.md.
    pause
    exit /b 1
  )
)

echo Installing / updating packages...
"venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo pip install failed. See INSTALL_WINDOWS.md.
  pause
  exit /b 1
)

echo Starting BOM SMT Comparator...
"venv\Scripts\python.exe" bom_smt_comparator.py
if errorlevel 1 pause
