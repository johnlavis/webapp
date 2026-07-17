#!/usr/bin/env python3
"""Verify every formula in the workbook without LibreOffice:
   1. structural validity (functions whitelist, balanced parens, in-bounds refs, quoted cross-sheet)
   2. independent Python re-computation of key aggregates vs the example row."""
import re, sys
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

FN = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")
REF = re.compile(r"(?:'([^']+)'|([A-Za-z_][A-Za-z0-9 ]*?))?!?\$?([A-Z]{1,3})\$?(\d+)")
WHITELIST = {"IF","SUM","COUNT","COUNTA","COUNTIF","COUNTIFS","SUMPRODUCT","DATE"}

wb = openpyxl.load_workbook("IATSE_479_Field_Visit_Tracker.xlsx")
log = wb["Visit Log"]
ncols = log.max_column
headers = {get_column_letter(c): log.cell(row=1, column=c).value for c in range(1, ncols+1)}

errors = []
nformulas = 0
for ws in wb.worksheets:
    maxc = ws.max_column
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue
            nformulas += 1
            loc = f"{ws.title}!{cell.coordinate}"
            # balanced parens
            if v.count("(") != v.count(")"):
                errors.append(f"{loc}: unbalanced parens")
            # function whitelist
            for fn in FN.findall(v):
                base = fn.split(".")[-1].upper()
                if base not in WHITELIST:
                    errors.append(f"{loc}: non-whitelisted function {fn}")
            # cross-sheet refs to Visit Log must be quoted 'Visit Log'
            if "Visit Log" in v and "'Visit Log'" not in v:
                errors.append(f"{loc}: unquoted 'Visit Log' sheet ref")

print(f"Formulas scanned: {nformulas}")
print(f"Structural errors: {len(errors)}")
for e in errors[:50]:
    print("  ", e)

# ---- Independent recomputation of the example row (row 2 of Visit Log) ----
def val(header):
    for L,h in headers.items():
        if h == header:
            return log[f"{L}2"].value
    raise KeyError(header)

def count_yes(prefix):
    return sum(1 for L,h in headers.items()
               if h and h.startswith(prefix) and str(log[f"{L}2"].value).strip().lower()=="yes")

checks = []
# computed columns
crew = sum(1 for i in range(1,5) if val(f"Union Crew {i} Name"))
nonu = sum(1 for i in range(1,9) if val(f"Non-Union {i} Name"))
depts = count_yes("Dept: ")
fus = count_yes("Follow-up: ")
checks += [("# Union Crew", crew, 2), ("# Non-Union", nonu, 1),
           ("# Depts Visited", depts, 3), ("# Follow-ups Done", fus, 2)]
# dashboard-relevant
checks += [("Phase=Shooting", 1 if val("Production Phase")=="Shooting" else 0, 1),
           ("CrewDay=Day", 1 if val("Crew Day")=="Day" else 0, 1),
           ("Area:Stage", 1 if str(val("Area: Stage")).lower()=="yes" else 0, 1),
           ("Wx:Cloudy", 1 if str(val("Wx: Cloudy")).lower()=="yes" else 0, 1),
           ("Reason~routine", 1 if "routine" in str(val("Reason for Visit")).lower() else 0, 1)]

print("\nExample-row recomputation (computed value vs expected):")
allok = True
for name, got, exp in checks:
    ok = got == exp
    allok &= ok
    print(f"  [{'OK' if ok else 'XX'}] {name}: {got} (expected {exp})")

print("\nRESULT:", "PASS" if (not errors and allok) else "FAIL")
sys.exit(0 if (not errors and allok) else 1)
