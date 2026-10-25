@echo off
cd /d "%~dp0"
title KeRRA Daily Report

where python >nul 2>nul
if errorlevel 1 (
  echo Python is not installed or not on PATH.
  echo Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH", then run this again.
  pause
  exit /b 1
)

rem Stop Streamlit asking for an e-mail address on first launch
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit"
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
  echo [general]> "%USERPROFILE%\.streamlit\credentials.toml"
  echo email = "">> "%USERPROFILE%\.streamlit\credentials.toml"
)

if not exist ".venv\Scripts\activate.bat" (
  echo First run: setting up. This needs internet once and takes a few minutes...
  python -m venv .venv
  call ".venv\Scripts\activate.bat"
  python -m pip install --upgrade pip
  pip install -r requirements.txt
  if errorlevel 1 (
    echo Setup failed. Check your internet connection and try again.
    pause
    exit /b 1
  )
) else (
  call ".venv\Scripts\activate.bat"
)

echo Starting the app. Keep this window open while you work; close it to stop the app.
streamlit run app.py --browser.gatherUsageStats false
pause
