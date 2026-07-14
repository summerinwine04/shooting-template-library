@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Visual Template Library

echo ============================================
echo   Visual Template Library - starting...
echo   (Keep this window open; close it to stop)
echo ============================================
echo.

rem pick a Python launcher: prefer "python", fall back to "py"
set "PY=python"
%PY% -c "import sys" >nul 2>nul || set "PY=py"

rem check python exists at all
%PY% -c "import sys" >nul 2>nul
if errorlevel 1 (
  echo [!] Python not found. Please install Python 3 from https://www.python.org/downloads/
  echo     During install, tick "Add Python to PATH".
  echo.
  pause
  exit /b 1
)

rem ensure openpyxl is installed
%PY% -c "import openpyxl" >nul 2>nul || (
  echo Installing dependency: openpyxl ...
  %PY% -m pip install openpyxl
)

rem start the server (opens the browser automatically)
%PY% "%~dp0serve.py"

echo.
echo Server stopped. Press any key to close this window.
pause >nul
