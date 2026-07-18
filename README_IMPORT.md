# Field Visit PDF → Excel — How to Import Your Visits

You fill out the **X-4 PDF** in your PDF app (Acrobat Reader, Mac Preview, or Edge),
save one file per visit, then run the importer to add each visit as a row in
`Field_Rep_Database.xlsx`.

Keep these together in one folder:
- `import_visits.py`  (the importer)
- `Field_Rep_Database.xlsx`  (your database — delete the 2 demo rows first)
- your filled PDFs

---

## One-time setup

### Mac
1. Open **Terminal** (Cmd+Space, type "Terminal").
2. Check Python:  `python3 --version`  (Macs have it; if missing, get it from python.org).
3. Install the two libraries:
   ```
   pip3 install pypdf openpyxl
   ```

### Windows
1. Install **Python** from https://python.org/downloads — on the first screen,
   TICK **"Add Python to PATH"**, then Install.
2. Open **Command Prompt** (Start → type "cmd").
3. Install the two libraries:
   ```
   pip install pypdf openpyxl
   ```

---

## Every time you want to import

1. Put your filled PDFs in the same folder as `import_visits.py`.
2. Open Terminal / Command Prompt in that folder:
   - **Mac:** in Finder, right-click the folder → *New Terminal at Folder*.
   - **Windows:** open the folder, click the address bar, type `cmd`, press Enter.
3. Run one of these:
   ```
   python import_visits.py visit_2026-07-15.pdf        # one file
   python import_visits.py *.pdf                        # every PDF in the folder
   python import_visits.py                              # every PDF in a .\FilledForms subfolder
   ```
   (On Mac use `python3` instead of `python`.)

`Field_Rep_Database.xlsx` updates with a new row per visit. Importing the same
file name twice is skipped, so you can keep all your PDFs in one folder and just
re-run it whenever you add new ones.

---

## What gets tracked (from the PDF's fields and checkmarks)

| Column | Comes from |
|---|---|
| Date of Visit | "Date of Visit" text field |
| Production | "Production Title" text field |
| Location Visited | the **Areas Visited** checkboxes you ticked (Stage / Location / Mill / Lockup / Other) |
| Departments Seen | the **Depts Visited** checkboxes you ticked |
| Proof of Visit | defaults to "Emailed production" (the X-4 has no email field — edit in Excel if needed) |
| plus | Field Rep, Phase, Reason, Arrival, Departure, Source PDF, Imported date |

---

## No Python? Send them to me instead
Drop your filled PDFs into our chat and I'll run the import and send back the
updated `Field_Rep_Database.xlsx`. No setup on your end.
