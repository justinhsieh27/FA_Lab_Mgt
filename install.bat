@echo off
setlocal EnableExtensions

cd /d "%~dp0"

where node >nul 2>&1
if errorlevel 1 (
  echo Node.js 22+ and npm are required. Please install Node.js before continuing.
  exit /b 1
)

where npm >nul 2>&1
if errorlevel 1 (
  echo Node.js 22+ and npm are required. Please install Node.js before continuing.
  exit /b 1
)

echo Installing FA Lab Case Management dependencies...
call npm install
if errorlevel 1 (
  echo Installation failed.
  exit /b 1
)

echo Installation complete. Start the system with: run.bat
