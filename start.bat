@echo off
title advizeo ML Satz Converter

echo.
echo  ============================================================
echo   advizeo ML Satz Converter
echo  ============================================================
echo.

:: Change to the script's directory
cd /d "%~dp0"

:: Try to find Python in common locations
set PYTHON=
for %%P in (
  python
  "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
  "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
  "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
  "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
  "C:\Python313\python.exe"
  "C:\Python312\python.exe"
  "C:\Python311\python.exe"
) do (
  if not defined PYTHON (
    %%P --version >nul 2>&1
    if not errorlevel 1 set PYTHON=%%P
  )
)

if not defined PYTHON (
  echo  FEHLER / ERROR: Python nicht gefunden.
  echo  Please install Python 3.10+ from https://python.org
  pause
  exit /b 1
)

echo  Python gefunden: %PYTHON%
echo.
echo  Pruefe Abhaengigkeiten / Checking dependencies...
%PYTHON% -m pip show flask >nul 2>&1
if errorlevel 1 (
  echo  Installiere Flask und openpyxl...
  %PYTHON% -m pip install flask openpyxl --quiet
  if errorlevel 1 (
    echo  FEHLER beim Installieren der Abhaengigkeiten.
    pause
    exit /b 1
  )
)

echo  Starte Server auf http://localhost:5000
echo  Zum Beenden: Ctrl+C
echo.
start "" http://localhost:5000
%PYTHON% app.py

pause
