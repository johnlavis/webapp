#!/usr/bin/env python3
"""
Import filled IATSE Local 479 X-4 Field Visit PDFs into a Master-Database-format
CSV that matches John's 'Set Visit Master Database' columns (plus X-4-only extras).

Usage (Mac):   python3 import_to_masterdb.py *.pdf
Output:        Set_Visit_Import_MasterDB.csv  (append-only; re-imports are skipped)

The output columns match the Master Database so rows can be pasted straight into it.
Nothing here touches your existing Drive files.
"""
import csv, os, sys, glob, datetime
from pypdf import PdfReader

OUT="Set_Visit_Import_MasterDB.csv"
MASTER=["Visit ID","Date","Year","Production","Production Code Name","Production Company",
        "Contract","Season","Unit","Studio","City","Notes","Issues Discussed","Source Document",
        "Source URL","Source Line","Confidence Level","Review Flag","Raw Entry","Production Location"]
EXTRA=["Departments Seen (X-4)","Areas Visited (X-4)","Proof of Visit (X-4)","Production Phase (X-4)",
       "Crew Day (X-4)","Arrival (X-4)","Departure (X-4)","Field Representative (X-4)","Source PDF (X-4)"]
COLS=MASTER+EXTRA
AREA={"Visited":"Stage","Location":"Location","Mill":"Mill","Lockup":"Lockup","Other Location":"Other"}

def txt(f,k):
    v=f.get(k,{}).get("/V"); return "" if v is None else str(v).strip()
def on(f,k):
    v=f.get(k,{}).get("/V")
    if v is None: return None
    s=str(v); return None if s in ("/Off","Off","") else s.lstrip("/")

def next_id(existing):
    n=0
    for r in existing:
        vid=r.get("Visit ID","")
        if vid.startswith("SV-") and vid[3:].isdigit(): n=max(n,int(vid[3:]))
    return n

def row_from_pdf(path, vid):
    f=PdfReader(path).get_fields() or {}
    date=txt(f,"Date of Visit"); yr=""
    for fmt in ("%Y-%m-%d","%m/%d/%Y","%m/%d/%y"):
        try: yr=str(datetime.datetime.strptime(date,fmt).year); break
        except: pass
    areas=[lbl for fld,lbl in AREA.items() if on(f,fld)]
    depts=sorted([on(f,k) for k in f if k.startswith("DeptVisit") and on(f,k)])
    prod=txt(f,"Production Title")
    r={c:"" for c in COLS}
    r.update({"Visit ID":vid,"Date":date,"Year":yr,"Production":prod,"Production Code Name":prod,
        "Unit":", ".join(areas),"Notes":txt(f,"Visit Notes if any"),
        "Source Document":"IATSE 479 X-4 Field Visit Form","Confidence Level":"Confirmed",
        "Raw Entry":f"{date} {prod} - {', '.join(depts)}".strip(),"Production Location":", ".join(areas),
        "Departments Seen (X-4)":", ".join(depts),"Areas Visited (X-4)":", ".join(areas),
        "Proof of Visit (X-4)":"Emailed production","Production Phase (X-4)":on(f,"ProductionPhase") or "",
        "Crew Day (X-4)":on(f,"CrewDay") or "","Arrival (X-4)":txt(f,"Arrival Time"),
        "Departure (X-4)":txt(f,"Departure Time"),"Field Representative (X-4)":txt(f,"Field Representative"),
        "Source PDF (X-4)":os.path.basename(path)})
    return r

def main(argv):
    args=argv[1:] or sorted(glob.glob("*.pdf"))
    paths=[]
    for a in args: paths+= sorted(glob.glob(a)) if any(c in a for c in "*?[") else [a]
    paths=[p for p in paths if p.lower().endswith(".pdf") and os.path.exists(p)]
    existing=[]
    if os.path.exists(OUT):
        with open(OUT,newline="") as fh: existing=list(csv.DictReader(fh))
    seen={r.get("Source PDF (X-4)") for r in existing}
    nid=next_id(existing); added=0
    for p in paths:
        if os.path.basename(p) in seen: print("skip:",os.path.basename(p)); continue
        nid+=1; existing.append(row_from_pdf(p,f"SV-{nid:05d}")); added+=1
        print("imported:",os.path.basename(p))
    with open(OUT,"w",newline="") as fh:
        w=csv.DictWriter(fh,fieldnames=COLS); w.writeheader(); w.writerows(existing)
    print(f"Done. Added {added}. {OUT} now has {len(existing)} rows.")

if __name__=="__main__": sys.exit(main(sys.argv))
