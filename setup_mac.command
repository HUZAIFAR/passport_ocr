#!/bin/bash
# ------------------------------------------------------------------
#  Passport Name Extractor - ONE-TIME SETUP FOR MAC
#  Double-click this file once. It installs the Python packages the
#  script needs into a private folder called ".venv" (nothing else on
#  your Mac is changed).
# ------------------------------------------------------------------
cd "$(dirname "$0")"
echo "=== Passport Name Extractor - Mac setup ==="
echo

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 was not found."
  echo "Please install it from  https://www.python.org/downloads/  then double-click this file again."
  echo
  read -p "Press Enter to close this window."
  exit 1
fi

echo "Using: $(python3 --version)"
echo "Creating a private Python environment (.venv) ..."
python3 -m venv .venv
if [ $? -ne 0 ]; then
  echo "Could not create the .venv folder. Please install Python from https://www.python.org/downloads/ and retry."
  read -p "Press Enter to close this window."
  exit 1
fi

echo "Installing packages - this can take a few minutes the first time ..."
./.venv/bin/python -m pip install --upgrade pip >/dev/null 2>&1
./.venv/bin/python -m pip install -r requirements.txt
if [ $? -ne 0 ]; then
  echo
  echo "Installation failed - see the messages above. Check your internet connection and try again."
  read -p "Press Enter to close this window."
  exit 1
fi

mkdir -p passports
echo
echo "------------------------------------------------------------------"
echo " Setup complete!"
echo " 1. Put your passport PDFs / scans into the 'passports' folder."
echo " 2. Double-click 'run_mac.command'."
echo "------------------------------------------------------------------"
read -p "Press Enter to close this window."
