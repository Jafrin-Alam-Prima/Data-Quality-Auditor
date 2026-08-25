#!/usr/bin/env bash
# Excel Data Quality Auditor - start the app on macOS or Linux.
set -e
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  echo "Creating a virtual environment..."
  python3 -m venv .venv
  source .venv/bin/activate
  python -m pip install --upgrade pip >/dev/null
  python -m pip install -r requirements.txt
else
  source .venv/bin/activate
fi

echo "Starting the Excel Data Quality Auditor..."
streamlit run app.py
