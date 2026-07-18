@echo off
REM Double-click this to import every filled X-4 PDF in this folder into
REM Field_Rep_Database.xlsx. Keep it in the same folder as import_visits.py.

cd /d "%~dp0"

echo ==============================================
echo   IATSE 479 - Field Visit Importer
echo ==============================================
echo.

REM find Python (either "python" or the "py" launcher)
set "PY=python"
where python >nul 2>nul || set "PY=py"
%PY% --version >nul 2>nul
if errorlevel 1 (
  echo Python is not installed yet.
  echo Get it from https://python.org/downloads and TICK "Add Python to PATH" during install,
  echo then double-click this file again.
  echo.
  pause
  exit /b 1
)

echo Checking the two required libraries (first run only)...
%PY% -m pip install --quiet pypdf openpyxl

echo Importing PDFs in this folder...
echo.
%PY% import_visits.py *.pdf

echo.
echo Finished. Your database is: Field_Rep_Database.xlsx
pause
