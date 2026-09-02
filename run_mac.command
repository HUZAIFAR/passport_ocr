#!/bin/bash
# ------------------------------------------------------------------
#  Passport Name Extractor - RUN (MAC)
#  Double-click this file. It reads every PDF / image in the
#  "passports" folder and writes passport_names.xlsx (+ .csv).
# ------------------------------------------------------------------
cd "$(dirname "$0")"

PY=./.venv/bin/python
if [ ! -x "$PY" ]; then
  echo "Setup has not been run yet - double-click 'setup_mac.command' first."
  echo "(Trying with the system Python anyway ...)"
  PY=python3
fi

"$PY" passport_ocr.py passports -o passport_names.xlsx
STATUS=$?

echo
if [ $STATUS -eq 0 ] && [ -f passport_names.xlsx ]; then
  echo "Opening passport_names.xlsx ..."
  open passport_names.xlsx
else
  echo "Something went wrong - read the messages above."
fi
read -p "Press Enter to close this window."
