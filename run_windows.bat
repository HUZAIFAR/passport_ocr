@echo off
REM ------------------------------------------------------------------
REM  Passport Name Extractor - RUN (WINDOWS)
REM  Double-click this file. It reads every PDF / image in the
REM  "passports" folder and writes passport_names.xlsx (+ .csv).
REM ------------------------------------------------------------------
cd /d "%~dp0"

set PY=.venv\Scripts\python.exe
if not exist "%PY%" (
  echo Setup has not been run yet - double-click "setup_windows.bat" first.
  echo ^(Trying with the system Python anyway ...^)
  set PY=python
)

"%PY%" passport_ocr.py passports -o passport_names.xlsx
set STATUS=%errorlevel%

echo.
if "%STATUS%"=="0" (
  if exist passport_names.xlsx (
    echo Opening passport_names.xlsx ...
    start "" passport_names.xlsx
  )
) else (
  echo Something went wrong - read the messages above.
)
pause
