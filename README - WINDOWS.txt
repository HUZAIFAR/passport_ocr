=====================================================================
   PASSPORT NAME EXTRACTOR  -  STEP BY STEP FOR WINDOWS
=====================================================================

WHAT IT DOES
------------
You put passport scans (PDF or photos) into a folder. The program reads
every one and creates an Excel sheet with:

    Last Name | First Name | Passport Number | Passport From | Expiry Date

Expired passports are coloured RED. Passports with less than 7 months of
validity left are coloured YELLOW. Anything the program is not 100% sure
about is highlighted so you can check it.

Everything runs on your own PC. Nothing is uploaded to the internet.
(Full list of features:  FEATURES - what it does.txt)


BEFORE YOU START
----------------
  * You need: a Windows 10 or 11 PC, an internet connection (for the
    one-time setup only), about 15 minutes, and the file  passport_ocr.zip .
  * Do the steps in order. Steps 1 to 4 are done ONCE. After that you
    only ever do Steps 5 and 6.
  * Do not rename or move the files inside the  passport_ocr  folder.


=====================================================================
STEP 1 - INSTALL PYTHON            (about 5 minutes, only once)
=====================================================================
  1. Open this link:                https://www.python.org/downloads/
  2. Click the big yellow button    "Download Python 3.xx"
  3. Open the downloaded file. It is in your Downloads folder and is
     called something like   python-3.13.x-amd64.exe
  4. IMPORTANT: on the very first screen of the installer, at the
     bottom, TICK the box      [x] Add python.exe to PATH
  5. Click  "Install Now" . Click  Yes  if Windows asks for permission.
  6. Click  Close  when it says "Setup was successful".


=====================================================================
STEP 2 - INSTALL TESSERACT         (about 3 minutes, only once)
=====================================================================
  Tesseract is the free program that reads the text in the scans.

  1. Open this link:   https://github.com/UB-Mannheim/tesseract/wiki
  2. Under the heading "The latest installers" click the link that
     looks like      tesseract-ocr-w64-setup-5.x.x.exe      to download it.
  3. Open the downloaded file. Click  Yes  if Windows asks for permission.
  4. Click  OK  (language), then  Next ,  I Agree ,  Next ,  Next ,
     Next ,  Install , then  Next  and  Finish .
     Do NOT change the install folder. It must stay
          C:\Program Files\Tesseract-OCR


=====================================================================
STEP 3 - PUT THE FOLDER ON YOUR DESKTOP
=====================================================================
  1. Right-click  passport_ocr.zip  >  Extract All...  >  Extract .
  2. A folder called  passport_ocr  appears. Drag it onto your Desktop.
  3. Open it. You should see these files:
        setup_windows.bat      run_windows.bat      passport_ocr.py
        passports (folder)     sample_passports (folder)
        README - WINDOWS.txt   FEATURES - what it does.txt   ...


=====================================================================
STEP 4 - RUN THE SETUP             (about 2 minutes, only once)
=====================================================================
  1. Double-click  setup_windows.bat .
  2. If Windows shows a blue box "Windows protected your PC":
        click  "More info"  then  "Run anyway" .
  3. A black window opens and installs a few things (you will see text
     scrolling). Wait until it says     Setup complete!
     then press any key to close it.


=====================================================================
STEP 5 - ADD THE PASSPORTS
=====================================================================
  Copy all the passport files (PDF, JPG, PNG, TIFF) into the folder
  called  passports  inside  passport_ocr .
  Sub-folders inside  passports  are fine.
  (There is a small text file already in there - you can delete it.)


=====================================================================
STEP 6 - RUN IT
=====================================================================
  1. Double-click  run_windows.bat .
     (Same "More info" > "Run anyway" if Windows asks.)
  2. A black window shows one line per passport as it works.
     It takes about 1 to 3 seconds per passport.
  3. When it is finished, Excel opens the results file by itself.
     Press any key to close the black window.

  Where is the result?  Inside the  passport_ocr  folder:
        passport_names.xlsx    (Excel)
        passport_names.csv     (same data, plain text)
  If Excel does not open by itself, just double-click passport_names.xlsx.

  Running it again: close Excel first (the file must not be open),
  then repeat Step 6. It re-reads everything in the passports folder
  and rewrites passport_names.xlsx.


=====================================================================
TRY IT FIRST  (optional but recommended)
=====================================================================
  The folder  sample_passports  contains 6 FAKE passports.
  Copy them into  passports  and run Step 6. All six should show
  Name Confidence = high (two of them are deliberately expired and
  will be red). Then delete them from  passports  again.


=====================================================================
READING THE EXCEL FILE
=====================================================================
  Whole row RED       passport has EXPIRED
  Whole row YELLOW    less than 7 months of validity left
  Blue name cells     name came from the machine-readable lines only;
                      looks clean, but have a quick look
  Orange cells        please check this value against the scan.
                      The Notes column says exactly what was unclear.
  Grey name cells     nothing could be read (blank page / very bad scan)

  Name Confidence "high" = the machine-readable lines AND the printed
  name on the passport agree letter for letter.

  What to do with a flagged (orange) row:
     open that scan, compare it with the Notes column, and correct the
     cell in Excel by hand if needed. Usually only a few rows.

  Tip: click the small filter arrow on "Validity" or "Name Confidence"
  to show only the problem rows.

  A second sheet called "Legend" inside the Excel file repeats this.


=====================================================================
IF SOMETHING GOES WRONG
=====================================================================
  "Python was not found"  (or the Microsoft Store opens instead)
        -> Do Step 1 again and make sure  "Add python.exe to PATH"
           is TICKED. If Python is already installed: open the
           installer again > Modify > Next > tick "Add Python to
           environment variables" > Install.

  "Tesseract OCR is not installed"
        -> Do Step 2 again and keep the default folder.

  "Windows protected your PC"
        -> Click "More info" then "Run anyway".

  "Could not write passport_names.xlsx"
        -> The file is open in Excel. Close it and run again.

  "No PDF or image files found"
        -> The  passports  folder is empty. Do Step 5.

  "Input folder/file not found"
        -> The  passports  folder was renamed or moved. It must be
           called  passports  and sit next to run_windows.bat.

  Many orange cells
        -> The scans are poor or cut off. Rescan at 300 dpi and make
           sure the two lines of  <<<<  text at the bottom of the
           passport page are fully visible.

  Anything else: ask the person who sent you this folder.
