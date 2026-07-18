#!/bin/bash
# Double-click this file to import every filled X-4 PDF in this folder into
# Field_Rep_Database.xlsx. Keep it in the same folder as import_visits.py.

cd "$(cd "$(dirname "$0")" && pwd)" || exit 1

echo "=============================================="
echo "  IATSE 479 - Field Visit Importer"
echo "=============================================="
echo

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 is not installed yet."
  echo "A macOS window may pop up offering to install the 'Command Line Developer Tools' - click Install,"
  echo "let it finish, then double-click this file again."
  xcode-select --install 2>/dev/null
  echo; read -r -p "Press Return to close." _; exit 1
fi

echo "Checking the two required libraries (first run only)..."
python3 -m pip install --quiet --user pypdf openpyxl 2>/dev/null || \
  python3 -m pip install --quiet pypdf openpyxl

echo "Importing PDFs in this folder..."
echo
python3 import_visits.py *.pdf

echo
echo "Finished. Your database is: Field_Rep_Database.xlsx"
read -r -p "Press Return to close this window." _
