@echo off
rem ===================================================================
rem  Chinese numeral to Arabic numeral converter - launcher
rem  Finds the .py file next to this .bat and runs it with a GUI.
rem  (This file is intentionally ASCII-only to avoid codepage trouble.)
rem ===================================================================
setlocal
cd /d "%~dp0"

set "SCRIPT="
for %%F in ("%~dp0*.py") do if not defined SCRIPT set "SCRIPT=%%~fF"
if not defined SCRIPT (
  echo.
  echo [ERROR] The converter .py file was not found in this folder.
  echo         Keep this .bat together with the .py file.
  echo.
  pause
  exit /b 1
)

set "RUN="
if not defined RUN ( where pythonw.exe >nul 2>nul && set "RUN=pythonw.exe" )
if not defined RUN ( where pyw.exe >nul 2>nul && set "RUN=pyw.exe" )
if not defined RUN ( where python.exe >nul 2>nul && set "RUN=python.exe" )
if not defined RUN ( where py.exe >nul 2>nul && set "RUN=py.exe" )
if not defined RUN (
  echo.
  echo [ERROR] Python 3 was not found on this computer.
  echo         Install it from https://www.python.org/downloads/
  echo         and tick "Add python.exe to PATH" during setup.
  echo.
  pause
  exit /b 1
)

start "" "%RUN%" "%SCRIPT%"
exit /b 0
