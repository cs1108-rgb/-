@echo off
setlocal
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js was not found. Please install Node.js 24 LTS from https://nodejs.org
  echo Then close this window and run this file again.
  pause
  exit /b 1
)
node "%~dp0windows-launcher.cjs" --setup
set "MEASURE_EXIT=%ERRORLEVEL%"
if not "%MEASURE_EXIT%"=="0" echo Operation failed. Read the error message above.
pause
exit /b %MEASURE_EXIT%
