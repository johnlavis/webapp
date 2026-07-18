#!/usr/bin/env python3
"""Page-style Field Visit Report + auto-updating Field Rep Summary (no macros).
   - Visit Log: enter one visit per row (dropdowns).
   - Visit Form: page-styled X-4, auto-populates from the visit you pick; Save-as-PDF for employer.
   - Field Rep Summary: the 5 tracked fields (Production, Location, Date, Departments, Proof), auto-updated.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.comments import Comment

MAXROW = 50
WORKROWS = 50

NAVY="1F3864"; S1="2E5496"; S2="1E7145"; S3="9C4A00"; S4="6B2E8F"; S5="8A1C2B"; CALC="555555"
WHITE="FFFFFF"; INPUTBG="FFFDE7"; SUMBG="EAF3EA"
thin=Side(style="thin",color="B0B0B0"); med=Side(style="medium",color="1F3864")
border=Border(left=thin,right=thin,top=thin,bottom=thin)

DEPTS=["Art","Construction","Costumes","Craft Service","Electrical","Greens","Grip","Locations",
       "Medic","Paint","Plaster","Props","Set Dec","Set Teacher","Sound","SPFX","Video"]
SIZES=["Small","Medium","Large","X-Large","2X-Large","3X-Large","4X-Large"]
AREAS=["Stage","Location","Mill","Lockup","Other"]
WX=["Sunny","Cloudy","Rain","Snow","Lightning","Winds","Extreme Cold","Extreme Heat"]
FOLLOWUPS=["IATSE Project Info Sheet","Current Crewlist","479 Roster Provided",
           "Availability List Provided","Lunch Reporting Schedule"]

VAL={"phase":'"Prep,Shooting,Wrap"',"crewday":'"Day,Night,Split"',
     "stewardstatus":'"On Production,Off Production"',"role":'"Key,2nd,3rd"',"yes":'"Yes"',
     "dept":'"'+",".join(DEPTS)+'"',"size":'"'+",".join(SIZES)+'"'}

cols=[]
def add(h,sec,val=None,w=16,fmt=None,cmt=None): cols.append((h,sec,val,w,fmt,cmt))

add("Visit ID",NAVY,None,12,None,"Unique ID per visit, e.g. 2026-001. One row = one set visit.")
add("Date of Visit",NAVY,None,13,"yyyy-mm-dd","Enter a real date (2026-07-15) so monthly stats work.")
add("Production Title",NAVY,None,26)
add("Location Visited",NAVY,None,24,None,"Where you physically went, e.g. 'Stage 4, Pinewood' or a location address.")
add("Field Representative",NAVY,None,20)
# Section 1
add("L479 No.",S1,None,13)
add("Date Production Notified",S1,None,15,"yyyy-mm-dd")
add("Production Phase",S1,"phase",15,None,"Prep / Shooting / Wrap")
add("Arrival Time",S1,None,12); add("Departure Time",S1,None,12)
add("Reason for Visit",S1,None,22,None,"routine / issue")
add("Crew Call",S1,None,12); add("Crew Day",S1,"crewday",11,None,"Day / Night / Split")
for a in AREAS: add(f"Area: {a}",S1,"yes",10,None,"Enter Yes if this area was visited.")
add("Visit Notes",S1,None,30)
add("Proof of Visit",S1,None,30,None,"How you documented the visit, e.g. 'Emailed production'.")
for w_ in WX: add(f"Wx: {w_}",S1,"yes",10,None,"Enter Yes if this weather applied.")
# Section 2
add("Steward 1 Name",S2,None,20); add("Steward 1 Status",S2,"stewardstatus",15); add("Steward 1 Start Date",S2,None,14,"yyyy-mm-dd")
add("Steward 2 Name",S2,None,20); add("Steward 2 Status",S2,"stewardstatus",15); add("Steward 2 Start Date",S2,None,14,"yyyy-mm-dd")
for i in range(1,5):
    add(f"Union Crew {i} Name",S2,None,20); add(f"Union Crew {i} Role",S2,"role",10,None,"Key / 2nd / 3rd"); add(f"Union Crew {i} Notes/Issues",S2,None,24)
add("Section 2 Notes",S2,None,28)
for d in DEPTS: add(f"Dept: {d}",S2,"yes",11,None,"Enter Yes if you saw this department.")
# Section 3
for i in range(1,9): add(f"Non-Union {i} Name",S3,None,20); add(f"Non-Union {i} Dept",S3,"dept",14)
add("Section 3 Notes",S3,None,28)
# Section 4
for i in range(1,6): add(f"Apparel {i} Name",S4,None,20); add(f"Apparel {i} Phone",S4,None,14); add(f"Apparel {i} Size",S4,"size",12)
# Section 5
add("UPM/Producer 1",S5,None,20); add("UPM/Producer 2",S5,None,20); add("Production Coordinator",S5,None,20)
for f_ in FOLLOWUPS: add(f"Follow-up: {f_}",S5,"yes",13,None,"Enter Yes if completed.")
add("Section 5 Notes",S5,None,30)
# computed
add("# Union Crew",CALC,None,12,"0"); add("# Non-Union",CALC,None,12,"0")
add("# Depts Visited",CALC,None,13,"0"); add("# Follow-ups Done",CALC,None,14,"0")

hidx={h:i+1 for i,(h,*_) in enumerate(cols)}
def L(h): return get_column_letter(hidx[h])

wb=openpyxl.Workbook()

# =========================== Visit Log ===========================
ws=wb.active; ws.title="Visit Log"; ws.sheet_view.showGridLines=False
for ci,(h,sec,val,w,fmt,cmt) in enumerate(cols,start=1):
    c=ws.cell(row=1,column=ci,value=h)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=9)
    c.fill=PatternFill("solid",fgColor=sec)
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=border
    ws.column_dimensions[get_column_letter(ci)].width=w
    if cmt: c.comment=Comment(cmt,"Form")
ws.row_dimensions[1].height=42

crew_cells=lambda r:[f"{L(f'Union Crew {i} Name')}{r}" for i in range(1,5)]
nonu_cells=lambda r:[f"{L(f'Non-Union {i} Name')}{r}" for i in range(1,9)]
d0,d1=L("Dept: Art"),L("Dept: Video"); f0,f1=L("Follow-up: IATSE Project Info Sheet"),L("Follow-up: Lunch Reporting Schedule")

for r in range(2,WORKROWS+1):
    for ci,(h,sec,val,w,fmt,cmt) in enumerate(cols,start=1):
        cell=ws.cell(row=r,column=ci); cell.font=Font(name="Arial",size=9)
        if sec==CALC: cell.fill=PatternFill("solid",fgColor="EFEFEF")
fmt_cols=[(h,fmt) for (h,sec,val,w,fmt,cmt) in cols if fmt]
for r in range(2,MAXROW+1):
    for h,fmt in fmt_cols: ws.cell(row=r,column=hidx[h]).number_format=fmt
    g=f'{L("Production Title")}{r}=""'
    ws[f"{L('# Union Crew')}{r}"]=f'=IF({g},"",COUNTA('+",".join(crew_cells(r))+"))"
    ws[f"{L('# Non-Union')}{r}"]=f'=IF({g},"",COUNTA('+",".join(nonu_cells(r))+"))"
    ws[f"{L('# Depts Visited')}{r}"]=f'=IF({g},"",COUNTIF({d0}{r}:{d1}{r},"Yes"))'
    ws[f"{L('# Follow-ups Done')}{r}"]=f'=IF({g},"",COUNTIF({f0}{r}:{f1}{r},"Yes"))'

dvs={}
for k,f in VAL.items():
    dv=DataValidation(type="list",formula1=f,allow_blank=True); ws.add_data_validation(dv); dvs[k]=dv
for h,sec,val,w,fmt,cmt in cols:
    if val: c=L(h); dvs[val].add(f"{c}2:{c}{MAXROW}")

ex={"Visit ID":"2026-001","Date of Visit":"2026-07-15","Production Title":"Northern Lights (S2)",
 "Location Visited":"Stage 4, Pinewood Studios","Field Representative":"J. Doe","L479 No.":"250415-001",
 "Date Production Notified":"2026-07-10","Production Phase":"Shooting","Arrival Time":"9:30 AM",
 "Departure Time":"1:15 PM","Reason for Visit":"routine","Crew Call":"7:00 AM","Crew Day":"Day",
 "Area: Stage":"Yes","Area: Location":"Yes","Wx: Cloudy":"Yes","Wx: Winds":"Yes",
 "Visit Notes":"All quiet. Stage 4 base camp.","Proof of Visit":"Emailed production",
 "Steward 1 Name":"A. Smith","Steward 1 Status":"On Production","Steward 1 Start Date":"2026-06-01",
 "Union Crew 1 Name":"R. Lee","Union Crew 1 Role":"Key","Union Crew 1 Notes/Issues":"None",
 "Union Crew 2 Name":"M. Ford","Union Crew 2 Role":"2nd",
 "Dept: Grip":"Yes","Dept: Electrical":"Yes","Dept: Set Dec":"Yes",
 "Non-Union 1 Name":"T. Nguyen","Non-Union 1 Dept":"Construction",
 "Apparel 1 Name":"R. Lee","Apparel 1 Phone":"555-0142","Apparel 1 Size":"Large",
 "UPM/Producer 1":"C. Ramos","Production Coordinator":"D. Park",
 "Follow-up: Current Crewlist":"Yes","Follow-up: 479 Roster Provided":"Yes",
 "Section 5 Notes":"Requested updated crew list; provided 479 roster."}
for h,v in ex.items():
    cc=ws.cell(row=2,column=hidx[h]); cc.value=v; cc.font=Font(name="Arial",size=9,italic=True,color="9A6A00")
ws.cell(row=2,column=hidx["Visit ID"]).comment=Comment("EXAMPLE ROW - overwrite or delete.","Form")
ws.freeze_panes="F2"
lastcol=get_column_letter(len(cols))
tbl=Table(displayName="VisitLog",ref=f"A1:{lastcol}{MAXROW}")
tbl.tableStyleInfo=TableStyleInfo(name="TableStyleLight1",showRowStripes=True)
ws.add_table(tbl)

# =========================== Field Rep Summary ===========================
fr=wb.create_sheet("Field Rep Summary"); fr.sheet_view.showGridLines=False
sum_cols=["Visit ID","Production","Location Visited","Date Visited","Departments Seen","Proof of Visit"]
widths=[12,26,26,13,40,26]
for i,(h,wd) in enumerate(zip(sum_cols,widths),start=1):
    c=fr.cell(row=1,column=i,value=h)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",fgColor=S2)
    c.alignment=Alignment(horizontal="center",vertical="center"); c.border=border
    fr.column_dimensions[get_column_letter(i)].width=wd
fr.row_dimensions[1].height=22
fr.freeze_panes="A2"

def depts_formula(r):
    parts="&".join([f'IF(\'Visit Log\'!{L(f"Dept: {d}")}{r}="Yes",", {d}","")' for d in DEPTS])
    return f'=IF(\'Visit Log\'!{L("Production Title")}{r}="","",IF(({parts})="","",MID({parts},3,300)))'

for r in range(2,MAXROW+1):
    g=f'\'Visit Log\'!{L("Production Title")}{r}=""'
    fr.cell(row=r,column=1).value=f'=IF({g},"",\'Visit Log\'!{L("Visit ID")}{r})'
    fr.cell(row=r,column=2).value=f'=IF({g},"",\'Visit Log\'!{L("Production Title")}{r})'
    fr.cell(row=r,column=3).value=f'=IF({g},"",\'Visit Log\'!{L("Location Visited")}{r})'
    fr.cell(row=r,column=4).value=f'=IF({g},"",\'Visit Log\'!{L("Date of Visit")}{r})'
    fr.cell(row=r,column=5).value=depts_formula(r)
    fr.cell(row=r,column=6).value=f'=IF({g},"",\'Visit Log\'!{L("Proof of Visit")}{r})'
    for cc in range(1,7):
        cell=fr.cell(row=r,column=cc); cell.font=Font(name="Arial",size=10); cell.border=border
        if cc==4: cell.number_format="yyyy-mm-dd"
        if r%2==0: cell.fill=PatternFill("solid",fgColor=SUMBG)

# small totals block to the right
fr.column_dimensions["H"].width=24; fr.column_dimensions["I"].width=12
idc=f"'Field Rep Summary'!$A$2:$A${MAXROW}"; pc=f"'Field Rep Summary'!$B$2:$B${MAXROW}"
fr["H1"]="TRACKER TOTALS"; fr["H1"].font=Font(name="Arial",bold=True,color=WHITE,size=10)
fr["H1"].fill=PatternFill("solid",fgColor=NAVY); fr["I1"].fill=PatternFill("solid",fgColor=NAVY)
fr["H2"]="Total Visits Logged"; fr["I2"]=f"=COUNTA({idc})"
fr["H3"]="Productions Tracked";  fr["I3"]=f'=SUMPRODUCT(({pc}<>"")/COUNTIF({pc},{pc}&""))'
for rr in (2,3):
    fr[f"H{rr}"].font=Font(name="Arial",size=10); fr[f"I{rr}"].font=Font(name="Arial",size=10,bold=True,color=NAVY)
    fr[f"I{rr}"].alignment=Alignment(horizontal="center"); fr[f"I{rr}"].number_format="0"

# =========================== Visit Form (page-style, auto-populate) ===========================
vf=wb.create_sheet("Visit Form"); vf.sheet_view.showGridLines=False
for col,w in zip("ABCDEFGH",[13,13,12,13,12,12,12,12]): vf.column_dimensions[col].width=w
PICK="$D$3"; IDR="'Visit Log'!$A$2:$A$"+str(MAXROW)
def V(h):
    return f'=IFERROR(INDEX(\'Visit Log\'!${L(h)}$2:${L(h)}${MAXROW},MATCH({PICK},{IDR},0)),"")'
def CB(h,label):
    idx=f'IFERROR(INDEX(\'Visit Log\'!${L(h)}$2:${L(h)}${MAXROW},MATCH({PICK},{IDR},0)),"")'
    return f'=IF({idx}="Yes","[X] {label}","[  ] {label}")'
def depts_pick():
    return f'=IFERROR(INDEX(\'Field Rep Summary\'!$E$2:$E${MAXROW},MATCH({PICK},{IDR},0)),"")'

lab=Font(name="Arial",size=7.5,color="666666")
valf=Font(name="Arial",size=10,color="000000")
h_font=Font(name="Arial",bold=True,size=9,color=WHITE)
def merge(rng): vf.merge_cells(rng)
def put(anchor,value,font=None,fill=None,align=None,wrap=False,bord=True,size=None):
    c=vf[anchor]; c.value=value
    c.font=font or Font(name="Arial",size=10)
    if fill: c.fill=PatternFill("solid",fgColor=fill)
    c.alignment=align or Alignment(horizontal="left",vertical="center",wrap_text=wrap,indent=1)
    if bord: c.border=border
def band(cell,text,color):
    put(cell,text,font=h_font,fill=color,align=Alignment(horizontal="left",vertical="center",indent=1),bord=False)
def field(label_anchor,label_text,val_anchor,val_formula,merge_val=None,merge_lab=None):
    if merge_lab: merge(merge_lab)
    put(label_anchor,label_text,font=lab,align=Alignment(horizontal="left",vertical="bottom",indent=1),bord=False)
    if merge_val: merge(merge_val)
    put(val_anchor,val_formula,font=valf,fill=INPUTBG,align=Alignment(horizontal="left",vertical="center",wrap_text=True,indent=1))

# ---- header ----
merge("A1:H1"); put("A1","IATSE Studio Mechanics Local 479",font=Font(name="Arial",bold=True,size=14,color=WHITE),
                    fill=NAVY,align=Alignment(horizontal="center",vertical="center"),bord=False)
merge("A2:H2"); put("A2","Field Visit Report  -  Form X-4 (2026)",font=Font(name="Arial",bold=True,size=11,color=NAVY),
                    align=Alignment(horizontal="center"),bord=False)
put("B3","Show Visit ID:",font=Font(name="Arial",bold=True,size=9),align=Alignment(horizontal="right"),bord=False)
put("D3","2026-001",font=Font(name="Arial",bold=True,size=11,color="0000FF"),fill="FFFF00",
    align=Alignment(horizontal="center"))
merge("E3:H3"); put("E3","<- pick a visit; the whole form fills in automatically",
                    font=Font(name="Arial",italic=True,size=8,color="888888"),bord=False,
                    align=Alignment(horizontal="left",indent=1))
dvpick=DataValidation(type="list",formula1=IDR,allow_blank=True); vf.add_data_validation(dvpick); dvpick.add("D3")

r=5
band(f"A{r}","SECTION 1  -  PRODUCTION VISIT DETAILS",S1); merge(f"A{r}:H{r}"); r+=1
field(f"A{r}","FIELD REPRESENTATIVE",f"A{r+1}",V("Field Representative"),merge_lab=f"A{r}:C{r}",merge_val=f"A{r+1}:C{r+1}")
field(f"D{r}","DATE OF VISIT",f"D{r+1}",V("Date of Visit"),merge_lab=f"D{r}:E{r}",merge_val=f"D{r+1}:E{r+1}")
field(f"F{r}","L479 No.",f"F{r+1}",V("L479 No."),merge_lab=f"F{r}:H{r}",merge_val=f"F{r+1}:H{r+1}")
r+=2
field(f"A{r}","PRODUCTION",f"A{r+1}",V("Production Title"),merge_lab=f"A{r}:E{r}",merge_val=f"A{r+1}:E{r+1}")
field(f"F{r}","PHASE",f"F{r+1}",V("Production Phase"),merge_lab=f"F{r}:H{r}",merge_val=f"F{r+1}:H{r+1}")
r+=2
field(f"A{r}","LOCATION VISITED",f"A{r+1}",V("Location Visited"),merge_lab=f"A{r}:E{r}",merge_val=f"A{r+1}:E{r+1}")
field(f"F{r}","CREW DAY",f"F{r+1}",V("Crew Day"),merge_lab=f"F{r}:H{r}",merge_val=f"F{r+1}:H{r+1}")
r+=2
field(f"A{r}","REASON FOR VISIT",f"A{r+1}",V("Reason for Visit"),merge_lab=f"A{r}:C{r}",merge_val=f"A{r+1}:C{r+1}")
field(f"D{r}","ARRIVAL",f"D{r+1}",V("Arrival Time"),merge_lab=f"D{r}:E{r}",merge_val=f"D{r+1}:E{r+1}")
field(f"F{r}","DEPARTURE / CREW CALL",f"F{r+1}",'=IFERROR(INDEX(\'Visit Log\'!$'+L("Departure Time")+'$2:$'+L("Departure Time")+f'${MAXROW},MATCH({PICK},{IDR},0)),"")&"  /  "&'+'IFERROR(INDEX(\'Visit Log\'!$'+L("Crew Call")+'$2:$'+L("Crew Call")+f'${MAXROW},MATCH({PICK},{IDR},0)),"")',merge_lab=f"F{r}:H{r}",merge_val=f"F{r+1}:H{r+1}")
r+=2
# Areas + Weather
put(f"A{r}","AREAS VISITED:",font=lab,align=Alignment(horizontal="left",vertical="center",indent=1),bord=False)
merge(f"A{r}:B{r}")
acols=["C","D","E","F","G"]
for i,a in enumerate(AREAS): put(f"{acols[i]}{r}",CB(f"Area: {a}",a),font=Font(name="Arial",size=9),bord=False,align=Alignment(horizontal="left"))
r+=1
put(f"A{r}","WEATHER:",font=lab,align=Alignment(horizontal="left",vertical="center",indent=1),bord=False); merge(f"A{r}:B{r}")
wcols=["C","D","E","F","G","H"]
for i,wname in enumerate(WX[:4]): put(f"{wcols[i]}{r}",CB(f"Wx: {wname}",wname),font=Font(name="Arial",size=8),bord=False,align=Alignment(horizontal="left"))
r+=1
put(f"A{r}","",bord=False); merge(f"A{r}:B{r}")
for i,wname in enumerate(WX[4:]): put(f"{wcols[i]}{r}",CB(f"Wx: {wname}",wname),font=Font(name="Arial",size=8),bord=False,align=Alignment(horizontal="left"))
r+=2

band(f"A{r}","SECTION 2  -  UNION CREW & DEPARTMENTS SEEN",S2); merge(f"A{r}:H{r}"); r+=1
field(f"A{r}","DEPARTMENTS SEEN",f"A{r+1}",depts_pick(),merge_lab=f"A{r}:H{r}",merge_val=f"A{r+1}:H{r+1}")
r+=2
field(f"A{r}","STEWARD 1 (name / status)",f"A{r+1}",'=IFERROR(INDEX(\'Visit Log\'!$'+L("Steward 1 Name")+'$2:$'+L("Steward 1 Name")+f'${MAXROW},MATCH({PICK},{IDR},0)),"")&"   -   "&'+'IFERROR(INDEX(\'Visit Log\'!$'+L("Steward 1 Status")+'$2:$'+L("Steward 1 Status")+f'${MAXROW},MATCH({PICK},{IDR},0)),"")',merge_lab=f"A{r}:D{r}",merge_val=f"A{r+1}:D{r+1}")
field(f"E{r}","STEWARD 2 (name / status)",f"E{r+1}",'=IFERROR(INDEX(\'Visit Log\'!$'+L("Steward 2 Name")+'$2:$'+L("Steward 2 Name")+f'${MAXROW},MATCH({PICK},{IDR},0)),"")&"   -   "&'+'IFERROR(INDEX(\'Visit Log\'!$'+L("Steward 2 Status")+'$2:$'+L("Steward 2 Status")+f'${MAXROW},MATCH({PICK},{IDR},0)),"")',merge_lab=f"E{r}:H{r}",merge_val=f"E{r+1}:H{r+1}")
r+=2
for i in range(1,5):
    field(f"A{r}",f"UNION CREW {i} (name / role)",f"A{r+1}",'=IFERROR(INDEX(\'Visit Log\'!$'+L(f"Union Crew {i} Name")+'$2:$'+L(f"Union Crew {i} Name")+f'${MAXROW},MATCH({PICK},{IDR},0)),"")&"   ("&'+'IFERROR(INDEX(\'Visit Log\'!$'+L(f"Union Crew {i} Role")+'$2:$'+L(f"Union Crew {i} Role")+f'${MAXROW},MATCH({PICK},{IDR},0)),"")&")"',merge_lab=f"A{r}:C{r}",merge_val=f"A{r+1}:C{r+1}")
    field(f"D{r}","NOTES / ISSUES",f"D{r+1}",V(f"Union Crew {i} Notes/Issues"),merge_lab=f"D{r}:H{r}",merge_val=f"D{r+1}:H{r+1}")
    r+=2

band(f"A{r}","SECTION 5  -  PRODUCTION NOTES & PROOF OF VISIT",S5); merge(f"A{r}:H{r}"); r+=1
field(f"A{r}","UPM / PRODUCER",f"A{r+1}",'=IFERROR(INDEX(\'Visit Log\'!$'+L("UPM/Producer 1")+'$2:$'+L("UPM/Producer 1")+f'${MAXROW},MATCH({PICK},{IDR},0)),"")&"   "&'+'IFERROR(INDEX(\'Visit Log\'!$'+L("UPM/Producer 2")+'$2:$'+L("UPM/Producer 2")+f'${MAXROW},MATCH({PICK},{IDR},0)),"")',merge_lab=f"A{r}:D{r}",merge_val=f"A{r+1}:D{r+1}")
field(f"E{r}","PRODUCTION COORDINATOR",f"E{r+1}",V("Production Coordinator"),merge_lab=f"E{r}:H{r}",merge_val=f"E{r+1}:H{r+1}")
r+=2
# PROOF - prominent
put(f"A{r}","PROOF OF VISIT",font=Font(name="Arial",bold=True,size=9,color=WHITE),fill=S5,
    align=Alignment(horizontal="left",vertical="center",indent=1)); merge(f"A{r}:B{r+1}")
merge(f"C{r}:H{r+1}"); put(f"C{r}",V("Proof of Visit"),font=Font(name="Arial",bold=True,size=11,color="1E7145"),
    fill=INPUTBG,align=Alignment(horizontal="left",vertical="center",indent=1))
r+=2
field(f"A{r}","VISIT NOTES",f"A{r+1}",V("Visit Notes"),merge_lab=f"A{r}:H{r}",merge_val=f"A{r+1}:H{r+1}")
vf.row_dimensions[r+1].height=34
r+=2
merge(f"A{r}:H{r}"); put(f"A{r}","For Field Representative use. IATSE Studio Mechanics Local 479 - Field Visit Report Form X-4.",
    font=Font(name="Arial",italic=True,size=7.5,color="888888"),align=Alignment(horizontal="center"),bord=False)
LASTROW=r

# print setup -> clean single-page PDF
vf.print_area=f"A1:H{LASTROW}"
vf.page_setup.orientation="portrait"
vf.page_setup.fitToWidth=1; vf.page_setup.fitToHeight=1
vf.sheet_properties.pageSetUpPr=PageSetupProperties(fitToPage=True)
vf.page_margins.left=vf.page_margins.right=0.4
vf.page_margins.top=vf.page_margins.bottom=0.5

# =========================== Instructions ===========================
ins=wb.create_sheet("Instructions"); ins.sheet_view.showGridLines=False
ins.column_dimensions["A"].width=2; ins.column_dimensions["B"].width=104
ins["B2"]="How this workbook works"; ins["B2"].font=Font(name="Arial",bold=True,size=16,color=NAVY)
lines=[
 "",
 "This file has NO macros, so it opens and prints anywhere with no security prompts - safe to send to an employer.",
 "",
 ("1. Enter each visit in the ","Visit Log"),
 "   - One row per set visit. Row 2 is a shaded EXAMPLE - overwrite or delete it.",
 "   - Give every visit a unique Visit ID (e.g. 2026-001, 2026-002...). That ID is how the form finds the visit.",
 "   - Cells with a dropdown arrow (Phase, Crew Day, Role, Dept, Size, all Yes fields) - pick from the list.",
 "   - Fill the three tracked fields especially: Location Visited, the Dept: columns you saw (Yes), and Proof of Visit.",
 "",
 ("2. Print / save a visit from the ","Visit Form"),
 "   - Pick a Visit ID in the yellow cell (top). The whole page-style form fills in automatically for that visit.",
 "   - File > Save As (or Export) > PDF  ->  you get a clean one-page Field Visit Report for your employer.",
 "   - Change the Visit ID to print a different visit. Nothing to retype.",
 "",
 ("3. Your tracked stats live in the ","Field Rep Summary"),
 "   - Updates automatically from the Visit Log. Columns: Visit ID, Production, Location Visited, Date Visited,",
 "     Departments Seen (auto-listed from the Yes boxes), and Proof of Visit.",
 "   - Totals (visits logged, productions tracked) are on the right.",
 "",
 "Note: this is a standalone workbook. When your own field-rep sheet on Google Drive is ready, share it and I'll",
 "align these columns to it. I have NOT touched your Drive.",
 "",
 "Source: IATSE Studio Mechanics Local 479 - Field Visit Report Form X-4 (2026).",
]
rr=4
for item in lines:
    c=ins.cell(row=rr,column=2)
    if isinstance(item,tuple):
        c.value=item[0]+item[1]; c.font=Font(name="Arial",size=12,bold=True,color=NAVY)
    else:
        c.value=item; c.font=Font(name="Arial",size=11,bold=item.startswith(("This file","Note:","Source")))
    c.alignment=Alignment(wrap_text=True,vertical="top"); rr+=1

# order: Instructions, Visit Form, Field Rep Summary, Visit Log
order=["Instructions","Visit Form","Field Rep Summary","Visit Log"]
wb._sheets.sort(key=lambda s: order.index(s.title))
wb.active=order.index("Visit Form")

out="IATSE_479_Field_Visit_Form_drive.xlsx"; wb.save(out)
print("saved",out,"| Log cols:",len(cols),"| form last row:",LASTROW)
