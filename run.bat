@echo off
REM Excel Data Quality Auditor - start the app on Windows.
cd /d "%~dp0"

if not exist ".venv" (
    echo Creating a virtual environment...
    python -m venv .venv
    call .venv\Scripts\activate.bat
    echo Installing the required packages...
    python -m pip install --upgrade pip >nul
    python -m pip install -r requirements.txt
) else (
    call .venv\Scripts\activate.bat
)

echo.
echo Starting the Excel Data Quality Auditor...
echo The app will open in your web browser.
echo Close this window to stop it.
echo.
streamlit run app.py
pause
