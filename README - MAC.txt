=====================================================================
   PASSPORT NAME EXTRACTOR  -  STEP BY STEP FOR MAC
=====================================================================

WHAT IT DOES
------------
You put passport scans (PDF or photos) into a folder. The program reads
every one and creates an Excel sheet with:

    Last Name | First Name | Passport Number | Passport From |
    Nationality | Expiry Date

Expired passports are coloured RED. Passports with less than 7 months of
validity left are coloured YELLOW. Anything the program is not 100% sure
about is highlighted so you can check it.

Everything runs on your own Mac. Nothing is uploaded to the internet.
(Full list of features:  FEATURES - what it does.txt)


BEFORE YOU START
----------------
  * You need: a Mac, an internet connection (for the one-time setup
    only), about 10 minutes, and the file  passport_ocr.zip .
  * Do the steps in order. Steps 1 to 3 are done ONCE. After that you
    only ever do Steps 4 and 5.
  * Do not rename or move the files inside the  passport_ocr  folder.


=====================================================================
STEP 1 - INSTALL PYTHON            (about 5 minutes, only once)
=====================================================================
  (Already have Python 3.9 or newer? Skip to Step 2.)

  1. Open this link in Safari:      https://www.python.org/downloads/
  2. Click the big yellow button    "Download Python 3.xx"
  3. Open the downloaded file. It is in your Downloads folder and is
     called something like   python-3.13.x-macos11.pkg
  4. Click  Continue > Continue > Continue > Agree > Install.
     Type your Mac password if it asks.
  5. Click  Close  when it says "The installation was successful".
     (A Finder window may open - you can close it.)


=====================================================================
STEP 2 - PUT THE FOLDER ON YOUR DESKTOP
=====================================================================
  1. Double-click  passport_ocr.zip . A folder called  passport_ocr  appears
     next to it.
  2. Drag that  passport_ocr  folder onto your Desktop.
  3. Open it. You should see these files:
        setup_mac.command      run_mac.command      passport_ocr.py
        passports (folder)     sample_passports (folder)
        README - MAC.txt       FEATURES - what it does.txt   ...


=====================================================================
STEP 3 - RUN THE SETUP             (about 2 minutes, only once)
=====================================================================

  >>> IF YOU RECEIVED THIS FOLDER BY WHATSAPP, EMAIL OR AIRDROP,
  >>> USE METHOD B. IT ALWAYS WORKS AND TAKES 20 SECONDS.

  ---------------------------------------------------------------
  METHOD A - double-click
  ---------------------------------------------------------------
  1. RIGHT-click  setup_mac.command  and choose  Open .
     (Right-click = click with two fingers on the trackpad, or hold
      the Control key while clicking.)

  2. If the Mac says "cannot be opened because it is from an
     unidentified developer" or "Apple could not verify ...":
        a. Click  Done  (or  OK ).
        b. Apple menu (top-left) > System Settings.
        c. Click  Privacy & Security  in the left list, then scroll
           down on the right until you see a line about
           setup_mac.command with an  "Open Anyway"  button.
        d. Click  Open Anyway , confirm with your password or Touch ID,
           then click  Open  in the next box.

  3. If instead it says  "setup_mac.command is damaged and can't be
     opened. You should move it to the Trash"  -  the file is NOT
     damaged and there will be NO "Open Anyway" button anywhere.
     WhatsApp / Mail put a block on the file. Click  Cancel
     (never "Move to Trash") and use METHOD B below.

  4. A black Terminal window opens. If the Mac asks
     "Terminal would like to access files in your Desktop folder",
     click  OK / Allow .

  5. The window installs a few things (you will see text scrolling).
     Wait until it says      Setup complete!      then press Enter.

  ---------------------------------------------------------------
  METHOD B - one line in Terminal  (fixes the block, always works)
  ---------------------------------------------------------------
  1. Press  Cmd + Space , type  Terminal , press Enter.
  2. Copy this WHOLE line, paste it into the Terminal window
     (Cmd + V), and press Enter:

          cd ~/Desktop/passport_ocr && xattr -cr . && chmod +x *.command && bash setup_mac.command

     (If the Mac asks to access your Desktop folder, click OK.)
  3. Wait until it says  Setup complete!  then press Enter.

  That command removes the download block and makes the files
  runnable. After it, double-clicking works normally too.

  It says "no such file or directory"?  The folder is not on your
  Desktop, or it is not called  passport_ocr . Put it on the Desktop
  (Step 2) and try again.


=====================================================================
STEP 4 - ADD THE PASSPORTS
=====================================================================
  Copy all the passport files (PDF, JPG, PNG, TIFF) into the folder
  called  passports  inside  passport_ocr .
  Sub-folders inside  passports  are fine.
  (There is a small text file already in there - you can delete it.)


=====================================================================
STEP 5 - RUN IT
=====================================================================
  1. RIGHT-click  run_mac.command  and choose  Open .
     (Blocked, or you used Method B in Step 3? Then run it the same
      way instead - Terminal, paste this line, press Enter:

          cd ~/Desktop/passport_ocr && bash run_mac.command      )
  2. A black window shows one line per passport as it works.
     It takes about 1 to 3 seconds per passport.
  3. When it is finished, Excel opens the results file by itself.
     Press Enter to close the black window.

  Where is the result?  Inside the  passport_ocr  folder:
        passport_names.xlsx    (Excel)
        passport_names.csv     (same data, plain text)
  If Excel does not open by itself, just double-click passport_names.xlsx.
  (No Excel? It also opens in Numbers or Google Sheets.)

  Terminal alternative:

          cd ~/Desktop/passport_ocr && bash run_mac.command

  Running it again: close Excel first (the file must not be open),
  then repeat Step 5. It re-reads everything in the passports folder
  and rewrites passport_names.xlsx.


=====================================================================
TRY IT FIRST  (optional but recommended)
=====================================================================
  The folder  sample_passports  contains 6 FAKE passports.
  Copy them into  passports  and run Step 5. All six should show
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

  Dates: the expiry date comes from the machine-readable lines at the
  bottom of the passport, which write dates unambiguously. Only when
  those cannot be read does the program fall back to the printed date,
  and it then reads it the way the issuing country writes it (India,
  Pakistan, UK, France, Australia: day/month/year - USA: month/day/year).
  If the country cannot be worked out, the row is flagged instead of
  guessed, with both possible dates in the Notes.

  A second sheet called "Legend" inside the Excel file repeats this.


=====================================================================
IF SOMETHING GOES WRONG
=====================================================================
  "Python 3 was not found"
        -> Do Step 1 again.

  "cannot be opened because it is from an unidentified developer" /
  "Apple could not verify ..."
        -> System Settings > Privacy & Security > Open Anyway.
           See Step 3, Method A.

  "setup_mac.command is damaged and can't be opened.
   You should move it to the Trash"
        -> The file is fine. WhatsApp / Mail blocked it, and for this
           message there is NO "Open Anyway" button anywhere.
           Click Cancel (never "Move to Trash") and use Step 3,
           Method B - the one Terminal line clears the block.

  "Could not write passport_names.xlsx"
        -> The file is open in Excel. Close it and run again.

  "No PDF or image files found"
        -> The  passports  folder is empty. Do Step 4.

  "Input folder/file not found"
        -> The  passports  folder was renamed or moved. It must be
           called  passports  and sit next to run_mac.command.

  Grey name cells / "not readable"
        -> The passport was too small in the scan for the <<<< lines to be
           read. Rescan with the passport filling more of the page, or scan
           the sheet at 300-600 dpi. As a guide the passport should fill at
           least about a quarter of the page width.

  Many orange cells
        -> The scans are poor or cut off. Rescan at 300 dpi and make
           sure the two lines of  <<<<  text at the bottom of the
           passport page are fully visible.

  Anything else: ask the person who sent you this folder.
