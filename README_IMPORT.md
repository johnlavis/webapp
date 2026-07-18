# Field Visit PDF → Excel — Windows Guide

You fill out the **X-4 PDF** in Acrobat Reader (or Edge), save one file per visit,
then import it so each visit becomes a row in `Field_Rep_Database.xlsx`.

## Put these in one folder (e.g. a "Field Visits" folder)
- `Import Visits.bat`        ← double-click this to import
- `import_visits.py`         ← the importer it runs
- `Field_Rep_Database.xlsx`  ← your database (delete the 2 demo rows first)
- your filled PDFs

---

## One-time: install Python
1. Go to https://python.org/downloads and install Python 3.
2. On the FIRST install screen, **tick "Add Python to PATH"**, then click Install.
   (This one checkbox is what makes the double-click work.)

---

## The easy way — double-click

1. Put your filled PDFs in the folder.
2. Double-click **`Import Visits.bat`**.
   - **First time only:** if a blue "Windows protected your PC" box appears, click
     **More info → Run anyway** (it's your own file).
   - The first run quietly installs two small libraries (needs internet once).
3. A window shows what it imported, then says "Press any key to continue."

`Field_Rep_Database.xlsx` now has a new row per visit. Importing the same file name
again is skipped, so you can keep all your PDFs in the folder and just re-run it when
you add new ones.

---

## The manual way (if you prefer typing)

Open the folder, click the address bar, type `cmd`, press Enter, then:

```
pip install pypdf openpyxl        # one time only
python import_visits.py *.pdf     # import every PDF in the folder
```

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

## No setup at all?
Drop your filled PDFs into our chat and I'll run the import and send back the updated
`Field_Rep_Database.xlsx`.

(A Mac version, `Import Visits.command`, is included too if you ever switch.)
