@echo off
cd /d "%~dp0"
python -m streamlit run app.py
if errorlevel 1 (
  echo.
  echo Could not start the app. Install dependencies first with:
  echo   python -m pip install -r requirements.txt
  echo.
  pause
)
