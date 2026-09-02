@echo off
REM ------------------------------------------------------------------
REM  Passport Name Extractor - ONE-TIME SETUP FOR WINDOWS
REM  Double-click this file once. It installs the Python packages the
REM  script needs into a private folder called ".venv".
REM ------------------------------------------------------------------
cd /d "%~dp0"
echo === Passport Name Extractor - Windows setup ===
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found.
  echo Install it from  https://www.python.org/downloads/
  echo IMPORTANT: on the first screen of the installer TICK "Add python.exe to PATH".
  echo Then double-click this file again.
  echo.
  pause
  exit /b 1
)
python --version

if not exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
  echo.
  echo WARNING: Tesseract OCR was not found in  C:\Program Files\Tesseract-OCR
  echo          Install it from  https://github.com/UB-Mannheim/tesseract/wiki
  echo          ^(download the latest "tesseract-ocr-w64-setup-*.exe", keep the default folder^)
  echo          The script cannot read passports without it.
  echo.
)

echo Creating a private Python environment (.venv) ...
python -m venv .venv
if errorlevel 1 (
  echo Could not create the .venv folder. Re-install Python from https://www.python.org/downloads/ and retry.
  pause
  exit /b 1
)

echo Installing packages - this can take a few minutes the first time ...
.venv\Scripts\python -m pip install --upgrade pip >nul 2>nul
.venv\Scripts\python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Installation failed - see the messages above. Check your internet connection and try again.
  pause
  exit /b 1
)

if not exist passports mkdir passports
echo.
echo ------------------------------------------------------------------
echo  Setup complete!
echo  1. Put your passport PDFs / scans into the "passports" folder.
echo  2. Double-click "run_windows.bat".
echo ------------------------------------------------------------------
pause
