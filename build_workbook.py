#!/usr/bin/env python3
"""Build the IATSE Local 479 Field Visit Report tracker workbook."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.comments import Comment

MAXROW = 300    # table / validation / dashboard span (up to ~300 visits)
WORKROWS = 60   # rows we pre-style & pre-fill with formulas (Table auto-extends the rest)

# ---- palette ----
NAVY   = "1F3864"
S1     = "2E5496"  # Section 1 - Production Visit Details
S2     = "1E7145"  # Section 2 - Union Crew
S3     = "9C4A00"  # Section 3 - Non-Union
S4     = "6B2E8F"  # Section 4 - Apparel
S5     = "8A1C2B"  # Section 5 - Production Notes
CALC   = "555555"  # computed
WHITE  = "FFFFFF"
INPUTBG = "FFFDE7"  # pale yellow input tint
thin = Side(style="thin", color="D0D0D0")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

DEPTS = ["Art","Construction","Costumes","Craft Service","Electrical","Greens","Grip",
         "Locations","Medic","Paint","Plaster","Props","Set Dec","Set Teacher","Sound","SPFX","Video"]
SIZES = ["Small","Medium","Large","X-Large","2X-Large","3X-Large","4X-Large"]

# validation formula strings (quoted comma list)
VAL = {
    "phase": '"Prep,Shooting,Wrap"',
    "crewday": '"Day,Night,Split"',
    "stewardstatus": '"On Production,Off Production"',
    "role": '"Key,2nd,3rd"',
    "yes": '"Yes"',
    "dept": '"' + ",".join(DEPTS) + '"',
    "size": '"' + ",".join(SIZES) + '"',
}

# column spec: (header, section_color, validation_key or None, width, number_format or None, comment)
cols = []
def add(h, sec, val=None, w=16, fmt=None, cmt=None):
    cols.append((h, sec, val, w, fmt, cmt))

# --- identity / frozen block ---
add("Visit ID", NAVY, None, 12, None, "Unique ID for each visit, e.g. 2026-001. One row = one set visit.")
add("Date of Visit", NAVY, None, 13, "yyyy-mm-dd", "Enter a real date (e.g. 2026-07-15) so monthly stats work.")
add("Production Title", NAVY, None, 26)
add("Field Representative", NAVY, None, 20)
# --- Section 1 ---
add("L479 No.", S1, None, 13)
add("Date Production Notified", S1, None, 15, "yyyy-mm-dd")
add("Production Phase", S1, "phase", 15, None, "Prep / Shooting / Wrap")
add("Arrival Time", S1, None, 12)
add("Departure Time", S1, None, 12)
add("Reason for Visit", S1, None, 24, None, "routine / issue")
add("Crew Call", S1, None, 12)
add("Crew Day", S1, "crewday", 11, None, "Day / Night / Split")
for a in ["Stage","Location","Mill","Lockup","Other"]:
    add(f"Area: {a}", S1, "yes", 10, None, "Enter Yes if this area was visited.")
add("Visit Notes", S1, None, 30)
for w_ in ["Sunny","Cloudy","Rain","Snow","Lightning","Winds","Extreme Cold","Extreme Heat"]:
    add(f"Wx: {w_}", S1, "yes", 10, None, "Enter Yes if this weather condition applied.")
# --- Section 2: Union Crew ---
add("Steward 1 Name", S2, None, 20)
add("Steward 1 Status", S2, "stewardstatus", 15)
add("Steward 1 Start Date", S2, None, 14, "yyyy-mm-dd")
add("Steward 2 Name", S2, None, 20)
add("Steward 2 Status", S2, "stewardstatus", 15)
add("Steward 2 Start Date", S2, None, 14, "yyyy-mm-dd")
for i in range(1,5):
    add(f"Union Crew {i} Name", S2, None, 20)
    add(f"Union Crew {i} Role", S2, "role", 10, None, "Key / 2nd / 3rd")
    add(f"Union Crew {i} Notes/Issues", S2, None, 24)
add("Section 2 Notes", S2, None, 28)
for d in DEPTS:
    add(f"Dept: {d}", S2, "yes", 11, None, "Enter Yes if this department was visited.")
# --- Section 3: Non-Union ---
for i in range(1,9):
    add(f"Non-Union {i} Name", S3, None, 20)
    add(f"Non-Union {i} Dept", S3, "dept", 14)
add("Section 3 Notes", S3, None, 28)
# --- Section 4: Apparel ---
for i in range(1,6):
    add(f"Apparel {i} Name", S4, None, 20)
    add(f"Apparel {i} Phone", S4, None, 14)
    add(f"Apparel {i} Size", S4, "size", 12)
# --- Section 5: Production Notes ---
add("UPM/Producer 1", S5, None, 20)
add("UPM/Producer 2", S5, None, 20)
add("Production Coordinator", S5, None, 20)
for f_ in ["IATSE Project Info Sheet","Current Crewlist","479 Roster Provided",
           "Availability List Provided","Lunch Reporting Schedule"]:
    add(f"Follow-up: {f_}", S5, "yes", 13, None, "Enter Yes if this follow-up was completed.")
add("Section 5 Notes", S5, None, 30)
# --- computed columns ---
add("# Union Crew", CALC, None, 12, "0")
add("# Non-Union", CALC, None, 12, "0")
add("# Depts Visited", CALC, None, 13, "0")
add("# Follow-ups Done", CALC, None, 14, "0")

# header index map
hidx = {h: i+1 for i,(h,*_) in enumerate(cols)}
def L(h):  # column letter for header
    return get_column_letter(hidx[h])

wb = openpyxl.Workbook()

# =========================================================================
#  SHEET: Visit Log
# =========================================================================
ws = wb.active
ws.title = "Visit Log"
ws.sheet_view.showGridLines = False

for ci,(h,sec,val,w,fmt,cmt) in enumerate(cols, start=1):
    c = ws.cell(row=1, column=ci, value=h)
    c.font = Font(name="Arial", bold=True, color=WHITE, size=9)
    c.fill = PatternFill("solid", fgColor=sec)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border
    ws.column_dimensions[get_column_letter(ci)].width = w
    if cmt:
        c.comment = Comment(cmt, "Form")
ws.row_dimensions[1].height = 42

# computed-column formulas + input styling for every data row
crew_name_cells = lambda r: [f"{L(f'Union Crew {i} Name')}{r}" for i in range(1,5)]
nonu_name_cells = lambda r: [f"{L(f'Non-Union {i} Name')}{r}" for i in range(1,9)]
dept_first, dept_last = L("Dept: Art"), L("Dept: Video")
fu_first, fu_last = L("Follow-up: IATSE Project Info Sheet"), L("Follow-up: Lunch Reporting Schedule")

# Light styling for a working band of rows (Table style handles the rest)
for r in range(2, WORKROWS+1):
    for ci,(h,sec,val,w,fmt,cmt) in enumerate(cols, start=1):
        cell = ws.cell(row=r, column=ci)
        cell.font = Font(name="Arial", size=9)
        if sec == CALC:
            cell.fill = PatternFill("solid", fgColor="EFEFEF")

# Number formats (date + computed cols) and computed formulas across the FULL span - cheap, no fills
fmt_cols = [(h, fmt) for (h,sec,val,w,fmt,cmt) in cols if fmt]
for r in range(2, MAXROW+1):
    for h, fmt in fmt_cols:
        ws.cell(row=r, column=hidx[h]).number_format = fmt
    g = f'{L("Production Title")}{r}=""'  # blank when the row has no production title
    ws[f"{L('# Union Crew')}{r}"]     = f'=IF({g},"",COUNTA(' + ",".join(crew_name_cells(r)) + "))"
    ws[f"{L('# Non-Union')}{r}"]      = f'=IF({g},"",COUNTA(' + ",".join(nonu_name_cells(r)) + "))"
    ws[f"{L('# Depts Visited')}{r}"]  = f'=IF({g},"",COUNTIF({dept_first}{r}:{dept_last}{r},"Yes"))'
    ws[f"{L('# Follow-ups Done')}{r}"]= f'=IF({g},"",COUNTIF({fu_first}{r}:{fu_last}{r},"Yes"))'

# data validations
dv_objs = {}
for key, f in VAL.items():
    dv = DataValidation(type="list", formula1=f, allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv)
    dv_objs[key] = dv
for h,sec,val,w,fmt,cmt in cols:
    if val:
        col = L(h)
        dv_objs[val].add(f"{col}2:{col}{MAXROW}")

# example row (row 2) - realistic values
ex = {
    "Visit ID":"2026-001","Date of Visit":"2026-07-15","Production Title":"Northern Lights (S2)",
    "Field Representative":"J. Doe","L479 No.":"250415-001","Date Production Notified":"2026-07-10",
    "Production Phase":"Shooting","Arrival Time":"9:30 AM","Departure Time":"1:15 PM",
    "Reason for Visit":"routine","Crew Call":"7:00 AM","Crew Day":"Day",
    "Area: Stage":"Yes","Area: Location":"Yes","Wx: Cloudy":"Yes","Wx: Winds":"Yes",
    "Visit Notes":"All quiet. Stage 4 base camp.",
    "Steward 1 Name":"A. Smith","Steward 1 Status":"On Production","Steward 1 Start Date":"2026-06-01",
    "Union Crew 1 Name":"R. Lee","Union Crew 1 Role":"Key","Union Crew 1 Notes/Issues":"None",
    "Union Crew 2 Name":"M. Ford","Union Crew 2 Role":"2nd",
    "Dept: Grip":"Yes","Dept: Electrical":"Yes","Dept: Set Dec":"Yes",
    "Non-Union 1 Name":"T. Nguyen","Non-Union 1 Dept":"Construction",
    "Apparel 1 Name":"R. Lee","Apparel 1 Phone":"555-0142","Apparel 1 Size":"Large",
    "UPM/Producer 1":"C. Ramos","Production Coordinator":"D. Park",
    "Follow-up: Current Crewlist":"Yes","Follow-up: 479 Roster Provided":"Yes",
    "Section 5 Notes":"Requested updated crew list; provided 479 roster.",
}
for h,v in ex.items():
    ws.cell(row=2, column=hidx[h]).value = v
    ws.cell(row=2, column=hidx[h]).font = Font(name="Arial", size=9, italic=True, color="9A6A00")
# tag example row
tag = ws.cell(row=2, column=hidx["Visit ID"])
tag.comment = Comment("EXAMPLE ROW - overwrite or delete. Copy the pattern for new visits.", "Form")

ws.freeze_panes = "E2"  # freeze header row + first 4 identity columns

# Excel table over full capacity so it auto-formats/filters
last_col = get_column_letter(len(cols))
tbl = Table(displayName="VisitLog", ref=f"A1:{last_col}{MAXROW}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True,
                                    showColumnStripes=False, showFirstColumn=False, showLastColumn=False)
ws.add_table(tbl)

# =========================================================================
#  SHEET: Dashboard
# =========================================================================
db = wb.create_sheet("Dashboard")
db.sheet_view.showGridLines = False
db.column_dimensions["A"].width = 2
db.column_dimensions["B"].width = 30
db.column_dimensions["C"].width = 12
db.column_dimensions["D"].width = 3
db.column_dimensions["E"].width = 30
db.column_dimensions["F"].width = 12

def title_cell(cell, text, color=NAVY, size=14):
    c = db[cell]; c.value = text
    c.font = Font(name="Arial", bold=True, color=WHITE, size=size)
    c.fill = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

def section_hdr(cell, text, color):
    c = db[cell]; c.value = text
    c.font = Font(name="Arial", bold=True, color=WHITE, size=10)
    c.fill = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

def label(cell, text):
    c = db[cell]; c.value = text
    c.font = Font(name="Arial", size=10)
def valf(cell, formula, bold=False, fmt="0"):
    c = db[cell]; c.value = formula
    c.font = Font(name="Arial", size=10, bold=bold, color=NAVY)
    c.number_format = fmt
    c.alignment = Alignment(horizontal="center")

db.merge_cells("B2:F2")
title_cell("B2", "IATSE Local 479  –  Field Visit Report  –  Stats Dashboard")
db.merge_cells("B3:F3")
db["B3"] = "Live stats – recalculates automatically as you add rows in the Visit Log."
db["B3"].font = Font(name="Arial", size=9, italic=True, color="666666")

vlog = "'Visit Log'"
dcol = f"{vlog}!${L('Date of Visit')}$2:${L('Date of Visit')}${MAXROW}"
pcol = f"{vlog}!${L('Production Title')}$2:${L('Production Title')}${MAXROW}"
idcol = f"{vlog}!${L('Visit ID')}$2:${L('Visit ID')}${MAXROW}"

# ---- KPI band ----
section_hdr("B5", "TOTALS", NAVY)
db.merge_cells("B5:C5")
label("B6","Total Visits Logged");           valf("C6", f"=COUNTA({idcol})", bold=True)
label("B7","Productions Tracked");            valf("C7", f"=SUMPRODUCT(({pcol}<>\"\")/COUNTIF({pcol},{pcol}&\"\"))", bold=True)
label("B8","Total Union Crew Logged");        valf("C8", f"=SUM({vlog}!${L('# Union Crew')}$2:${L('# Union Crew')}${MAXROW})", bold=True)
label("B9","Total Non-Union Logged");         valf("C9", f"=SUM({vlog}!${L('# Non-Union')}$2:${L('# Non-Union')}${MAXROW})", bold=True)
label("B10","Follow-ups Completed");          valf("C10", f"=SUM({vlog}!${L('# Follow-ups Done')}$2:${L('# Follow-ups Done')}${MAXROW})", bold=True)

# ---- Production Phase ----
section_hdr("E5", "BY PRODUCTION PHASE", S1)
db.merge_cells("E5:F5")
phcol = f"{vlog}!${L('Production Phase')}$2:${L('Production Phase')}${MAXROW}"
for i,p in enumerate(["Prep","Shooting","Wrap"]):
    label(f"E{6+i}", p); valf(f"F{6+i}", f'=COUNTIF({phcol},"{p}")')

# ---- Crew Day ----
section_hdr("E10", "BY CREW DAY", S1)
db.merge_cells("E10:F10")
cdcol = f"{vlog}!${L('Crew Day')}$2:${L('Crew Day')}${MAXROW}"
for i,p in enumerate(["Day","Night","Split"]):
    label(f"E{11+i}", p); valf(f"F{11+i}", f'=COUNTIF({cdcol},"{p}")')

# ---- Reason ----
section_hdr("B12", "BY REASON FOR VISIT", S1)
db.merge_cells("B12:C12")
rcol = f"{vlog}!${L('Reason for Visit')}$2:${L('Reason for Visit')}${MAXROW}"
label("B13","Routine (contains 'routine')"); valf("C13", f'=COUNTIF({rcol},"*routine*")')
label("B14","Issue (contains 'issue')");     valf("C14", f'=COUNTIF({rcol},"*issue*")')

# ---- Areas Visited ----
r0 = 17
section_hdr(f"B{r0-1}", "AREAS VISITED", S1); db.merge_cells(f"B{r0-1}:C{r0-1}")
for i,a in enumerate(["Stage","Location","Mill","Lockup","Other"]):
    col = L(f"Area: {a}")
    label(f"B{r0+i}", a); valf(f"C{r0+i}", f'=COUNTIF({vlog}!${col}$2:${col}${MAXROW},"Yes")')

# ---- Weather ----
section_hdr(f"E{r0-1}", "WEATHER CONDITIONS", S1); db.merge_cells(f"E{r0-1}:F{r0-1}")
wx = ["Sunny","Cloudy","Rain","Snow","Lightning","Winds","Extreme Cold","Extreme Heat"]
for i,a in enumerate(wx):
    col = L(f"Wx: {a}")
    label(f"E{r0+i}", a); valf(f"F{r0+i}", f'=COUNTIF({vlog}!${col}$2:${col}${MAXROW},"Yes")')

# ---- Departments visited (frequency) ----
dr = 27
section_hdr(f"B{dr-1}", "DEPARTMENTS VISITED (frequency)", S2); db.merge_cells(f"B{dr-1}:C{dr-1}")
half = (len(DEPTS)+1)//2
for i,d in enumerate(DEPTS):
    col = L(f"Dept: {d}")
    if i < half:
        cellL, cellV = f"B{dr+i}", f"C{dr+i}"
    else:
        cellL, cellV = f"E{dr+(i-half)}", f"F{dr+(i-half)}"
    label(cellL, d); valf(cellV, f'=COUNTIF({vlog}!${col}$2:${col}${MAXROW},"Yes")')
section_hdr(f"E{dr-1}", "DEPARTMENTS (cont.)", S2); db.merge_cells(f"E{dr-1}:F{dr-1}")

# ---- Follow-ups completion ----
fr = dr + half + 1
section_hdr(f"B{fr-1}", "FOLLOW-UP COMPLETION", S5); db.merge_cells(f"B{fr-1}:C{fr-1}")
fus = ["IATSE Project Info Sheet","Current Crewlist","479 Roster Provided",
       "Availability List Provided","Lunch Reporting Schedule"]
for i,f_ in enumerate(fus):
    col = L(f"Follow-up: {f_}")
    label(f"B{fr+i}", f_); valf(f"C{fr+i}", f'=COUNTIF({vlog}!${col}$2:${col}${MAXROW},"Yes")')

# ---- Monthly trend (selectable year) ----
mr = fr
section_hdr(f"E{mr-1}", "VISITS BY MONTH", S3); db.merge_cells(f"E{mr-1}:F{mr-1}")
db[f"E{mr}"] = "Year:"; db[f"E{mr}"].font = Font(name="Arial", size=10, bold=True)
db[f"F{mr}"] = 2026
db[f"F{mr}"].font = Font(name="Arial", size=10, bold=True, color="0000FF")
db[f"F{mr}"].fill = PatternFill("solid", fgColor="FFFF00")
db[f"F{mr}"].number_format = "0"
db[f"F{mr}"].alignment = Alignment(horizontal="center")
yr = f"$F${mr}"
months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
for i,m in enumerate(months):
    row = mr+1+i
    label(f"E{row}", m)
    valf(f"F{row}",
         f'=COUNTIFS({dcol},">="&DATE({yr},{i+1},1),{dcol},"<"&DATE({yr},{i+2},1))')

# ---- By Production (auto-count; user lists production names) ----
pr = mr + 14
section_hdr(f"B{pr-1}", "VISITS BY PRODUCTION  (type a title in col B)", S1)
db.merge_cells(f"B{pr-1}:C{pr-1}")
db[f"B{pr}"] = "Production Title"; db[f"C{pr}"] = "Visits"
for c in (f"B{pr}", f"C{pr}"):
    db[c].font = Font(name="Arial", bold=True, size=9, color=WHITE)
    db[c].fill = PatternFill("solid", fgColor="808080")
    db[c].alignment = Alignment(horizontal="center")
# pre-fill one from example, formulas for 20 slots
db[f"B{pr+1}"] = ex["Production Title"]
db[f"B{pr+1}"].font = Font(name="Arial", size=10, italic=True, color="9A6A00")
for i in range(20):
    row = pr+1+i
    bcell = f"B{row}"
    db[bcell].fill = PatternFill("solid", fgColor=INPUTBG)
    db[bcell].border = border
    if i>0:
        db[bcell].font = Font(name="Arial", size=10)
    db[f"C{row}"] = f'=IF({bcell}="","",COUNTIF({pcol},{bcell}))'
    db[f"C{row}"].font = Font(name="Arial", size=10, color=NAVY, bold=True)
    db[f"C{row}"].number_format = "0"
    db[f"C{row}"].alignment = Alignment(horizontal="center")
    db[f"C{row}"].border = border

# =========================================================================
#  SHEET: Instructions
# =========================================================================
ins = wb.create_sheet("Instructions")
ins.sheet_view.showGridLines = False
ins.column_dimensions["A"].width = 2
ins.column_dimensions["B"].width = 100
ins.merge_cells("B2:B2")
ins["B2"] = "How to use this workbook"
ins["B2"].font = Font(name="Arial", bold=True, size=16, color=NAVY)
lines = [
    ("", ""),
    ("1. Fill out the ", "Visit Log"),
    ("   • One row = one set visit. Start on row 3 (row 2 is a shaded EXAMPLE you can overwrite or delete).", ""),
    ("   • The first 4 columns (Visit ID, Date, Production, Field Rep) stay frozen as you scroll right.", ""),
    ("   • Columns are colour-coded by the 5 form sections, matching the paper X-4 form:", ""),
    ("        Navy = ID   • Blue = §1 Visit Details   • Green = §2 Union Crew   • Orange = §3 Non-Union   • Purple = §4 Apparel   • Red = §5 Production Notes", ""),
    ("   • Cells with a dropdown arrow (Phase, Crew Day, Role, Dept, Size, and all Yes fields) – pick from the list.", ""),
    ("   • Check-box items from the form (Areas Visited, Weather, Departments, Follow-ups) are 'Yes' columns – type/pick Yes when it applies, leave blank otherwise.", ""),
    ("   • Dates: enter as real dates (e.g. 2026-07-15) so the monthly chart works.", ""),
    ("   • Grey columns on the far right (# Union Crew, # Non-Union, # Depts, # Follow-ups) fill in automatically – don't type in them.", ""),
    ("", ""),
    ("2. Watch the ", "Dashboard"),
    ("   • Every stat recalculates automatically the moment you add or edit a Visit Log row. No buttons, no macros.", ""),
    ("   • Totals, phase/crew-day/reason breakdowns, areas, weather, department frequency, and follow-up completion are all automatic.", ""),
    ("   • VISITS BY MONTH: change the yellow Year cell to see any year.", ""),
    ("   • VISITS BY PRODUCTION: type production titles in column B and the visit count fills in beside it.", ""),
    ("", ""),
    ("3. Capacity", ""),
    ("   • Built for up to 500 visits. Fill rows top-to-bottom. Need more? Ask and it can be extended.", ""),
    ("", ""),
    ("Source form: IATSE Studio Mechanics Local 479 – Field Visit Report Form X-4 (2026). Every one of the form's 102 fields maps to a column here.", ""),
]
r = 4
for a,b in lines:
    c = ins.cell(row=r, column=2)
    if b:
        c.value = a + b
        c.font = Font(name="Arial", size=11, bold=True, color=NAVY)
    else:
        c.value = a
        c.font = Font(name="Arial", size=11 if not a.startswith(("1.","2.","3.")) else 12,
                      bold=a.startswith(("1.","2.","3.","Source")))
    c.alignment = Alignment(wrap_text=True, vertical="top")
    r += 1

# order sheets: Instructions, Visit Log, Dashboard
wb.move_sheet("Instructions", -(wb.sheetnames.index("Instructions")))
wb.active = wb.sheetnames.index("Visit Log")

out = "IATSE_479_Field_Visit_Tracker.xlsx"
wb.save(out)
print("saved", out, "columns:", len(cols))
