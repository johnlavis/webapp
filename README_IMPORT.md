# Field Visit PDF → Excel — Setup Guide (Windows & Mac)

You fill out the **X-4 PDF** in your PDF app (Acrobat Reader, Edge, or Mac Preview),
save one file per visit, then import it so each visit becomes a row in
`Field_Rep_Database.xlsx`.

## Put these in one folder (e.g. a "Field Visits" folder)
- `Import Visits.bat`        ← Windows: double-click this to import
- `Import Visits.command`    ← Mac: double-click this to import
- `import_visits.py`         ← the importer both launchers run
- `Field_Rep_Database.xlsx`  ← your database (delete the 2 demo rows first)
- your filled PDFs

---

## Windows

**One-time — install Python**
1. Go to https://python.org/downloads and install Python 3.
2. On the FIRST install screen, **tick "Add Python to PATH"**, then Install.

**Import (easy way)**
1. Put your filled PDFs in the folder.
2. Double-click **`Import Visits.bat`**.
   - First time: if "Windows protected your PC" appears, click **More info → Run anyway**.
   - First run also installs two small libraries (needs internet once).

**Import (typing, optional)** — open the folder, click the address bar, type `cmd`, Enter:
```
pip install pypdf openpyxl        # one time only
python import_visits.py *.pdf     # import every PDF in the folder
```

---

## Mac

**One-time — install libraries** (Macs already have Python 3). Open Terminal (Cmd+Space → "Terminal"):
```
pip3 install pypdf openpyxl
```

**Import (easy way)**
1. Put your filled PDFs in the folder.
2. Double-click **`Import Visits.command`**.
   - First time: if it says *"unidentified developer"*, right-click the file → **Open** → **Open**.

**Import (typing, optional)** — right-click the folder in Finder → *New Terminal at Folder*:
```
python3 import_visits.py *.pdf
```

---

## Both platforms

Importing the same file name twice is skipped, so keep all your PDFs in the folder and
just re-run the launcher whenever you add new ones.

### What gets tracked (from the PDF's fields and checkmarks)

| Column | Comes from |
|---|---|
| Date of Visit | "Date of Visit" text field |
| Production | "Production Title" text field |
| Location Visited | the **Areas Visited** checkboxes you ticked (Stage / Location / Mill / Lockup / Other) |
| Departments Seen | the **Depts Visited** checkboxes you ticked |
| Proof of Visit | defaults to "Emailed production" (the X-4 has no email field — edit in Excel if needed) |
| plus | Field Rep, Phase, Reason, Arrival, Departure, Source PDF, Imported date |

---

## No setup at all?
Drop your filled PDFs into our chat and I'll run the import and send back the updated
`Field_Rep_Database.xlsx`.
