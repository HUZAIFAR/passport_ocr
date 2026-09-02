# Passport Name Extractor

Put a folder of passport scans in, get an Excel sheet out with each person's
**Last Name, First Name, Passport Number, issuing country and Expiry Date**.
Expired passports are coloured red, passports with less than 7 months left yellow.

> **New here? Open `README - MAC.txt` or `README - WINDOWS.txt`** for numbered,
> click-by-click instructions. This file is the fuller reference.

- Works with PDF, JPG, PNG and TIFF scans. Indian, UK, US and any other passport that
  has the two lines of `<<<<` machine-readable text at the bottom (they all do).
- Runs **completely offline on your own computer**. Nothing is uploaded anywhere.
- Every value comes from the machine-readable zone (MRZ) and is checked two ways: the MRZ's
  own check digits, and the printed fields on the page. Anything that does not check out
  is highlighted so you can look at those few by hand.

---

## The short version

1. Install **Python** (Windows users also install **Tesseract**).
2. Double-click **`setup_mac.command`** or **`setup_windows.bat`**. Once only.
3. Drop the passport files into the **`passports`** folder and double-click
   **`run_mac.command`** or **`run_windows.bat`**. Excel opens with the results.

Want to test it first? Copy the fake passports from **`sample_passports`** into
`passports` and run. They should all come out as `high` confidence.

---

## What is in this folder

| File | What it is |
|---|---|
| `README - MAC.txt`, `README - WINDOWS.txt` | **Start here.** Step-by-step instructions in plain text. |
| `FEATURES - what it does.txt` | Everything the program does, handles and does not do. |
| `setup_mac.command` / `setup_windows.bat` | **One-time setup.** Double-click once to install what the program needs. |
| `run_mac.command` / `run_windows.bat` | **Run.** Double-click every time you want to process the folder. |
| `passports/` | Put your passport files in here. Sub-folders are fine. |
| `sample_passports/` | A few **fake** test passports to try the program with. |
| `passport_names.xlsx` / `.csv` | The result. Created next to the script when you run it. |
| `passport_ocr.py` | The program itself. You never need to open it. |
| `requirements.txt` | List of Python add-ons the setup installs. |

---

## How it reads a passport

1. **MRZ.** Every passport has two lines of machine-readable text at the bottom of the data page:

   ```
   P<INDSHARMA<<RAHUL<KUMAR<<<<<<<<<<<<<<<<<<<<
   J8369854<4IND9001014M2601017<<<<<<<<<<<<<<<8
   ```

   Line 1 holds the surname, then `<<`, then the given names. Line 2 holds the passport
   number, the issuing country, date of birth, sex and the **expiry date**, each protected by a
   check digit. The names, number, country and expiry date in the spreadsheet all come from here.
2. **Cross-check.** The printed *Surname*, *Given Name(s)*, *Passport No.* and *Date of Expiry*
   fields are read as well. When they agree with the MRZ the value is marked verified / `high`.
   When they disagree, the row is flagged and the Notes column shows both readings.
3. **Validity.** The expiry date is compared with today's date: before today = **EXPIRED**,
   less than 7 months away = **EXPIRES SOON**.
4. The MRZ is read at several magnifications and the readings are compared, sideways and
   upside-down scans are rotated automatically, and a PDF that already contains text is read directly.

---

## Reading the results

| Column | Meaning |
|---|---|
| File / Page | Which scan the row came from. A PDF with several passports gives several rows. |
| Last Name (Surname) / First Name / Middle | From the MRZ. Hyphens appear as spaces and apostrophes are dropped, exactly as the MRZ encodes them (`OKONKWO SMITH`, `DSOUZA`). |
| Passport Number | From MRZ line 2, validated by its check digit and the printed number. |
| Passport From | Issuing country, e.g. `India (IND)`. |
| Expiry Date | From MRZ line 2, validated by its check digit and the printed date. |
| Validity | `EXPIRED`, `EXPIRES SOON (<7 months)`, `OK` or `UNKNOWN - check`. |
| Months Left | Months from today until expiry (negative = expired). |
| Name Confidence | `high` = MRZ and printed name fields agree. `medium` = MRZ only, looks clean. `low` / `none` = check by hand. |
| Passport No. Check / Expiry Check | `verified (...)` or the reason it needs checking. |
| Read From / Notes | Where values came from and, for flagged rows, exactly what was unclear. |
| MRZ line | The raw machine-readable name line, to compare with the scan. |

**Colours** (also on the *Legend* sheet inside the Excel file):

| Colour | Meaning |
|---|---|
| Whole row **red** | Passport has expired. |
| Whole row **yellow** | Less than 7 months of validity left. |
| Name cells **blue** | Name from the MRZ only (printed fields not readable). Quick glance. |
| Name cells **orange** | Name needs checking: MRZ and printed fields disagree, or the read is incomplete. |
| Name cells **grey** | Nothing readable (blank page, photo only, very poor scan). |
| Passport No. / Expiry cells **orange** | That value needs checking: check digit failed, printed field differs, or unreadable. |

Tip: use the filter arrows on *Validity* and *Name Confidence* to see the problem rows first.

---

## Tips for good results

- Scan at **300 dpi** or better and make sure the **whole data page** is in the scan,
  especially the two `<<<<` lines at the very bottom.
- Sideways or upside-down scans are handled automatically. Phone photos work; flat scans work better.
- About 1 to 3 seconds per file.

---

## Running from the command line (optional)

```bash
python3 passport_ocr.py /path/to/folder -o results.xlsx
```

| Option | Meaning |
|---|---|
| `folder` (first argument) | Folder of scans, or a single file. Default: `passports` next to the script. |
| `-o results.xlsx` | Where to write the Excel file. A `.csv` with the same name is written too. |
| `--engine tesseract` | Force Tesseract. Default: Apple Vision on Mac, Tesseract elsewhere. |
| `--today 2027-01-15` | Judge validity as of another date (for testing or planning ahead). |
| `--verbose` | Print the text the OCR saw for every flagged row. |
| `--dpi 400` | Render PDF pages at a higher resolution (slower). |

On Mac the setup uses Apple's built-in text recognition, so Tesseract is not needed there.
If you prefer Tesseract on a Mac: `brew install tesseract`, then add `--engine tesseract`.

---

## If something goes wrong

| Message / problem | What to do |
|---|---|
| *Python was not found* / *python3: command not found* | Install Python (see the .txt guide). On Windows, tick **Add python.exe to PATH** in the installer. |
| *Tesseract OCR is not installed* (Windows) | Install Tesseract and keep the default folder. |
| Mac: *cannot be opened because it is from an unidentified developer* | Right-click the file > **Open**, or System Settings > Privacy & Security > **Open Anyway**. |
| *Could not write passport_names.xlsx* | The file is open in Excel. Close it and run again. |
| *No PDF or image files found* | The `passports` folder is empty. |
| Lots of orange cells | Scans are probably low quality or cropped. Rescan at 300 dpi with the bottom `<<<<` lines visible. |
| A single-name passport holder | Shows the name in Last Name with First Name empty, flagged so you can decide how to record it. |

---

## Privacy

Everything happens on the computer running the script. No internet connection is used
while processing (only the one-time setup downloads the Python add-ons). When you are done,
you can simply delete the `passports` folder and the Excel file.
