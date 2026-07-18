#!/usr/bin/env python3
"""Render a sample FILLED Field Visit Report to PDF (what the employer receives)."""
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor

NAVY=HexColor("#1F3864"); S1=HexColor("#2E5496"); S2=HexColor("#1E7145"); S5=HexColor("#8A1C2B")
INK=HexColor("#000000"); GREY=HexColor("#666666"); YEL=HexColor("#FFFDE7"); GREEN=HexColor("#1E7145")

# ---- example visit ----
D=dict(vid="2026-001",prod="Northern Lights (S2)",loc="Stage 4, Pinewood Studios",date="2026-07-15",
       rep="J. Doe",l479="250415-001",phase="Shooting",arr="9:30 AM",dep="1:15 PM",call="7:00 AM",
       crewday="Day",reason="routine",
       areas={"Stage":1,"Location":1,"Mill":0,"Lockup":0,"Other":0},
       wx={"Sunny":0,"Cloudy":1,"Rain":0,"Snow":0,"Lightning":0,"Winds":1,"Extreme Cold":0,"Extreme Heat":0},
       depts="Electrical, Grip, Set Dec",
       stew1="A. Smith  -  On Production",stew2="",
       crew=[("R. Lee (Key)","None"),("M. Ford (2nd)",""),("",""),("","")],
       upm="C. Ramos",coord="D. Park",proof="Emailed production",
       notes="All quiet. Stage 4 base camp.")

W,H=letter
c=canvas.Canvas("Sample_Field_Visit_Report.pdf",pagesize=letter)
M=42; x0=M; x1=W-M; y=H-M

def band(y,text,color,h=18):
    c.setFillColor(color); c.rect(x0,y-h,x1-x0,h,fill=1,stroke=0)
    c.setFillColor(HexColor("#FFFFFF")); c.setFont("Helvetica-Bold",9)
    c.drawString(x0+6,y-h+5,text); return y-h

def field(x,w,y,label,value,h=26,vfont=("Helvetica",10),vcolor=INK,fill=YEL):
    c.setFillColor(fill); c.rect(x,y-h,w,h,fill=1,stroke=0)
    c.setStrokeColor(HexColor("#B0B0B0")); c.setLineWidth(0.5); c.rect(x,y-h,w,h,fill=0,stroke=1)
    c.setFillColor(GREY); c.setFont("Helvetica",6.5); c.drawString(x+4,y-8,label.upper())
    c.setFillColor(vcolor); c.setFont(*vfont); c.drawString(x+5,y-21,str(value))
    return h

# ---- header ----
c.setFillColor(NAVY); c.rect(x0,y-30,x1-x0,30,fill=1,stroke=0)
c.setFillColor(HexColor("#FFFFFF")); c.setFont("Helvetica-Bold",15)
c.drawCentredString(W/2,y-20,"IATSE Studio Mechanics Local 479")
y-=30
c.setFillColor(NAVY); c.setFont("Helvetica-Bold",11)
c.drawCentredString(W/2,y-14,"Field Visit Report  -  Form X-4 (2026)")
y-=20
c.setFillColor(INK); c.setFont("Helvetica-Bold",9)
c.drawString(x0+2,y-12,"Visit ID:")
c.setFillColor(HexColor("#FFF200")); c.rect(x0+52,y-16,70,15,fill=1,stroke=0)
c.setFillColor(HexColor("#0000FF")); c.setFont("Helvetica-Bold",11); c.drawString(x0+58,y-13,D["vid"])
y-=22

# ---- Section 1 ----
y=band(y,"SECTION 1  -  PRODUCTION VISIT DETAILS",S1)
w3=(x1-x0-8)/3
field(x0,w3,y,"Field Representative",D["rep"]); field(x0+w3+4,w3,y,"Date of Visit",D["date"]); field(x0+2*(w3+4),w3,y,"L479 No.",D["l479"]); y-=30
wA=(x1-x0-4)*0.66; wB=(x1-x0-4)-wA
field(x0,wA,y,"Production",D["prod"],vfont=("Helvetica-Bold",11)); field(x0+wA+4,wB,y,"Phase",D["phase"]); y-=30
field(x0,wA,y,"Location Visited",D["loc"],vfont=("Helvetica-Bold",11),vcolor=NAVY); field(x0+wA+4,wB,y,"Crew Day",D["crewday"]); y-=30
w4=(x1-x0-12)/4
field(x0,w4,y,"Reason for Visit",D["reason"]); field(x0+w4+4,w4,y,"Arrival",D["arr"])
field(x0+2*(w4+4),w4,y,"Departure",D["dep"]); field(x0+3*(w4+4),w4,y,"Crew Call",D["call"]); y-=30
# areas
c.setFillColor(GREY); c.setFont("Helvetica-Bold",7.5); c.drawString(x0+2,y-10,"AREAS VISITED:")
c.setFillColor(INK); c.setFont("Helvetica",9); ax=x0+90
for a,on in D["areas"].items():
    c.drawString(ax,y-10,("[X] " if on else "[  ] ")+a); ax+=78
y-=16
c.setFillColor(GREY); c.setFont("Helvetica-Bold",7.5); c.drawString(x0+2,y-10,"WEATHER:")
c.setFillColor(INK); c.setFont("Helvetica",8.5); wx0=x0+90; i=0
for w_,on in D["wx"].items():
    xx=wx0+(i%4)*118;
    if i%4==0 and i>0: y-=13
    c.drawString(xx,y-10,("[X] " if on else "[  ] ")+w_); i+=1
y-=20

# ---- Section 2 ----
y=band(y,"SECTION 2  -  UNION CREW & DEPARTMENTS SEEN",S2)
field(x0,x1-x0,y,"Departments Seen",D["depts"],vfont=("Helvetica-Bold",11),vcolor=GREEN); y-=30
half=(x1-x0-4)/2
field(x0,half,y,"Steward 1 (name / status)",D["stew1"]); field(x0+half+4,half,y,"Steward 2 (name / status)",D["stew2"]); y-=30
for i,(nm,note) in enumerate(D["crew"],1):
    field(x0,half,y,f"Union Crew {i} (name / role)",nm); field(x0+half+4,half,y,"Notes / Issues",note); y-=28
y-=2

# ---- Section 5 ----
y=band(y,"SECTION 5  -  PRODUCTION NOTES & PROOF OF VISIT",S5)
field(x0,half,y,"UPM / Producer",D["upm"]); field(x0+half+4,half,y,"Production Coordinator",D["coord"]); y-=30
# proof prominent
c.setFillColor(S5); c.rect(x0,y-30,110,30,fill=1,stroke=0)
c.setFillColor(HexColor("#FFFFFF")); c.setFont("Helvetica-Bold",9); c.drawString(x0+6,y-19,"PROOF OF VISIT")
c.setFillColor(YEL); c.rect(x0+110,y-30,x1-x0-110,30,fill=1,stroke=0)
c.setStrokeColor(HexColor("#B0B0B0")); c.setLineWidth(0.5); c.rect(x0+110,y-30,x1-x0-110,30,fill=0,stroke=1)
c.setFillColor(GREEN); c.setFont("Helvetica-Bold",12); c.drawString(x0+118,y-19,D["proof"]); y-=32
field(x0,x1-x0,y,"Visit Notes",D["notes"],h=34); y-=38

c.setFillColor(GREY); c.setFont("Helvetica-Oblique",7.5)
c.drawCentredString(W/2,y-6,"For Field Representative use.  IATSE Studio Mechanics Local 479 - Field Visit Report Form X-4.")

c.showPage(); c.save()
print("saved Sample_Field_Visit_Report.pdf")
